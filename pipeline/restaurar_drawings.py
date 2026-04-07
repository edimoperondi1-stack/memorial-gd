"""
restaurar_drawings.py
---------------------
Restaura drawings, imagens e shapes que o openpyxl remove ao salvar.

Após o openpyxl preencher e salvar o template, esta função:
  1. Copia os arquivos xl/drawings/* e xl/media/* do template original
  2. Restaura os xl/worksheets/_rels/* (referências drawing ↔ sheet)
  3. Restaura as tags <drawing> nos XMLs dos sheets
  4. Atualiza [Content_Types].xml para registrar os novos parts

Isso preserva o diagrama unifilar, logos e demais shapes do template original.
"""

import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


def _map_sheet_names_to_files(z: zipfile.ZipFile) -> dict:
    """Retorna um dict {nome_aba: 'xl/worksheets/sheetX.xml'} lendo workbook.xml e _rels."""
    mapa = {}
    try:
        wb_xml = z.read("xl/workbook.xml")
        rels_xml = z.read("xl/_rels/workbook.xml.rels")
        
        # Pega a Rels do workbook
        root_rels = ET.fromstring(rels_xml)
        ns_rels = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
        rels = {}
        for rel in root_rels.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
            rels[rel.attrib["Id"]] = rel.attrib["Target"]

        # Pega as Sheets do workbook
        root_wb = ET.fromstring(wb_xml)
        ns_wb = {"ns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
        
        for sheet in root_wb.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"):
            nome = sheet.attrib.get("name")
            rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if nome and rid and rid in rels:
                target = rels[rid]
                if target.startswith("worksheets/"):
                    target = "xl/" + target
                elif target.startswith("/xl/"):
                    target = target[1:]
                elif "worksheets/" in target:
                    target = "xl/worksheets/" + target.split("/")[-1]
                mapa[nome] = target
    except Exception as e:
        print(f"Erro ao mapear sheets: {e}")
    return mapa


def restaurar_drawings(caminho_preenchido: str, caminho_template: str) -> str:
    """
    Restaura drawings/shapes do template original no arquivo preenchido pelo openpyxl.

    Args:
        caminho_preenchido: xlsx salvo pelo openpyxl (COM dados, SEM drawings).
        caminho_template: xlsx original (SEM dados, COM drawings).

    Returns:
        Caminho do arquivo restaurado (substitui o preenchido in-place).
    """
    tmp_path = tempfile.mktemp(suffix=".xlsx")

    with zipfile.ZipFile(caminho_template, "r") as z_tpl, \
         zipfile.ZipFile(caminho_preenchido, "r") as z_src, \
         zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as z_out:

        # Catalogar arquivos do template
        tpl_files = set(z_tpl.namelist())
        src_files = set(z_src.namelist())

        prefixos_restaurar = (
            "xl/drawings/",
            "xl/media/",
            "xl/comments",       # xl/comments1.xml, etc.
            "xl/printerSettings/",
            "xl/theme/",         # ESSENCIAL PARA O LIBREOFFICE RENDERIZAR SHAPES
        )

        # Mapear abas pelo nome para cruzar arquivos
        tpl_map = _map_sheet_names_to_files(z_tpl)
        src_map = _map_sheet_names_to_files(z_src)
        
        # Inverter para src_file -> tpl_file
        src_file_to_tpl_file = {}
        for nome_aba, src_target in src_map.items():
            if nome_aba in tpl_map:
                src_file_to_tpl_file[src_target] = tpl_map[nome_aba]

        # Arquivos que vamos copiar forçadamente do template
        arquivos_do_template = set()
        for f in tpl_files:
            for p in prefixos_restaurar:
                if f.startswith(p):
                    arquivos_do_template.add(f)
                    break

        # Rels dos sheets: Copiar todos do template pois contêm referências aos drawings,
        # MAS com o nome do arquivo src correspondente!
        # Se tpl tem xl/worksheets/_rels/sheet4.xml.rels, mas no src ele se chama sheet2.xml,
        # vamos salvar como xl/worksheets/_rels/sheet2.xml.rels.
        rels_mapeados = {} # tpl_rel_file -> src_rel_file -> content
        for src_file, tpl_file in src_file_to_tpl_file.items():
            tpl_rel = tpl_file.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels"
            src_rel = src_file.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels"
            if tpl_rel in tpl_files:
                rels_mapeados[src_rel] = z_tpl.read(tpl_rel)
                arquivos_do_template.add(src_rel)

        # Processar todos os arquivos do source
        arquivos_escritos = set()

        for f in src_files:
            if f in arquivos_do_template:
                # Este arquivo vai ser substituído pelo do template
                continue

            if f == "[Content_Types].xml":
                # Vamos reconstruir o Content_Types
                continue

            if f.startswith("xl/worksheets/sheet") and f.endswith(".xml") and "_rels" not in f:
                # Descobrir o nome interno desta aba (ex: sheet12.xml) e seu nome oficial
                tpl_file_corr = src_file_to_tpl_file.get(f)
                nome_oficial_aba = next((nom for nom, targ in src_map.items() if targ == f), "")

                if tpl_file_corr and tpl_file_corr in tpl_files:
                    # IGNORAR Restauração de Desenhos na Aba DU-SOLAR
                    # Motivo: Usaremos uma marca d'água perfeitamente sobreposta no Passo 4.
                    if "DU-SOLAR" in nome_oficial_aba.upper():
                        content = z_src.read(f).decode("utf-8")
                    else:
                        content = _restaurar_drawing_tag(
                            z_src.read(f).decode("utf-8"),
                            z_tpl.read(tpl_file_corr).decode("utf-8")
                        )
                else:
                    content = z_src.read(f).decode("utf-8")

                # Garantir xmlns:r em TODOS os sheets (openpyxl pode remover
                # mas deixar r:id em <hyperlinks> ou outros elementos)
                content = _ensure_r_namespace(content)
                    
                z_out.writestr(f, content)
                arquivos_escritos.add(f)
                continue

            # Copiar do source (openpyxl) se não for ser sobrescrito pelo Rels mapeado
            if f not in rels_mapeados:
                z_out.writestr(f, z_src.read(f))
                arquivos_escritos.add(f)

        # Escrever Rels mapeados
        for src_rel, content in rels_mapeados.items():
            z_out.writestr(src_rel, content)
            arquivos_escritos.add(src_rel)

        # Copiar arquivos do template (drawings, media)
        for f in arquivos_do_template:
            if f not in arquivos_escritos:
                z_out.writestr(f, z_tpl.read(f))
                arquivos_escritos.add(f)

        # Copiar arquivos do template que não existem no source
        # (podem ter sido removidos pelo openpyxl)
        for f in tpl_files:
            if f not in arquivos_escritos:
                for p in prefixos_restaurar:
                    if f.startswith(p):
                        z_out.writestr(f, z_tpl.read(f))
                        arquivos_escritos.add(f)

        # Reconstruir [Content_Types].xml
        content_types = _reconstruir_content_types(
            z_src.read("[Content_Types].xml").decode("utf-8"),
            z_tpl.read("[Content_Types].xml").decode("utf-8"),
            arquivos_escritos,
        )
        z_out.writestr("[Content_Types].xml", content_types)

    # Substituir o arquivo original
    shutil.move(tmp_path, caminho_preenchido)

    # Log: quantos arquivos de drawing e media foram restaurados
    n_drawings = sum(1 for f in arquivos_do_template if f.startswith("xl/drawings/"))
    n_media = sum(1 for f in arquivos_do_template if f.startswith("xl/media/"))
    print(f"  [restaurar_drawings] {n_drawings} drawings + {n_media} media files restaurados do template")

    return caminho_preenchido


def _ensure_r_namespace(src: str) -> str:
    """Garante que xmlns:r está declarado no <worksheet> tag raiz.
    
    openpyxl remove este namespace do root ao salvar, mas pode deixar
    declarações inline em <hyperlink xmlns:r="...">. Precisamos do namespace
    no root para que <drawing r:id="..."/>, <legacyDrawing r:id="..."/>
    e outros elementos que usam r: funcionem.
    """
    # Verificar se xmlns:r está no <worksheet> root tag especificamente
    ws_tag_match = re.search(r'<worksheet\s[^>]*>', src)
    if ws_tag_match and 'xmlns:r=' in ws_tag_match.group(0):
        return src  # Já tem no root — OK
    
    r_ns = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    return re.sub(
        r'(<worksheet\s+xmlns="[^"]*")',
        r'\1 ' + r_ns,
        src,
        count=1,
    )


def _restaurar_drawing_tag(sheet_xml_src: str, sheet_xml_tpl: str | None) -> str:
    """
    Se o template tinha uma tag <drawing r:id="..."/> ou <legacyDrawing .../> neste sheet,
    restaurá-la no XML do source na posição correta (antes de pageMargins/pageSetup).
    """
    if sheet_xml_tpl is None:
        return sheet_xml_src

    drawing_match = re.search(r'<drawing\s+r:id="([^"]+)"\s*/>', sheet_xml_tpl)
    legacy_match = re.search(r'<legacyDrawing\s+r:id="([^"]+)"\s*/>', sheet_xml_tpl)

    if not drawing_match and not legacy_match:
        return sheet_xml_src

    def insert_tag(src: str, tag: str) -> str:
        # A especificação OOXML exige que <drawing> e <legacyDrawing> venham DEPOIS
        # de tags como <pageMargins> e <pageSetup>, e ANTES das abaixo:
        for end_tag in ["<picture", "<oleObjects", "<controls", "<webPublishItems", "<tableParts", "<extLst", "</worksheet>"]:
            if end_tag in src:
                return src.replace(end_tag, tag + end_tag, 1)
        return src

    if drawing_match:
        src_has = re.search(r'<drawing\s+r:id="[^"]+"\s*/>', sheet_xml_src)
        if not src_has:
            sheet_xml_src = insert_tag(sheet_xml_src, drawing_match.group(0))

    if legacy_match:
        src_has = re.search(r'<legacyDrawing\s+r:id="[^"]+"\s*/>', sheet_xml_src)
        if not src_has:
            sheet_xml_src = insert_tag(sheet_xml_src, legacy_match.group(0))

    return sheet_xml_src


def _reconstruir_content_types(ct_src: str, ct_tpl: str, arquivos_no_zip: set) -> str:
    """
    Reconstrói [Content_Types].xml combinando registros do source e template.
    Garante que todos os drawings e media estão registrados.
    """
    # Extrair todos os Overrides e Defaults do template
    tpl_overrides = dict(re.findall(
        r'<Override\s+PartName="([^"]+)"\s+ContentType="([^"]+)"\s*/>', ct_tpl
    ))
    tpl_defaults = dict(re.findall(
        r'<Default\s+Extension="([^"]+)"\s+ContentType="([^"]+)"\s*/>', ct_tpl
    ))

    # Extrair do source
    src_overrides = dict(re.findall(
        r'<Override\s+PartName="([^"]+)"\s+ContentType="([^"]+)"\s*/>', ct_src
    ))
    src_defaults = dict(re.findall(
        r'<Default\s+Extension="([^"]+)"\s+ContentType="([^"]+)"\s*/>', ct_src
    ))

    # Combinar: priorizar source, adicionar faltantes do template
    all_defaults = {**tpl_defaults, **src_defaults}
    all_overrides = {**src_overrides}  # source tem prioridade

    # Adicionar overrides do template para arquivos que existem no zip
    for part_name, content_type in tpl_overrides.items():
        # Normalizar: PartName começa com /
        file_in_zip = part_name.lstrip("/")
        if file_in_zip in arquivos_no_zip and part_name not in all_overrides:
            all_overrides[part_name] = content_type

    # Garantir defaults para extensões de media
    if "png" not in all_defaults:
        all_defaults["png"] = "image/png"
    if "jpeg" not in all_defaults:
        all_defaults["jpeg"] = "image/jpeg"
    if "vml" not in all_defaults:
        all_defaults["vml"] = "application/vnd.openxmlformats-officedocument.vmlDrawing"

    # Montar XML
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    parts.append('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')

    for ext, ct in sorted(all_defaults.items()):
        parts.append(f'<Default Extension="{ext}" ContentType="{ct}"/>')

    for pn, ct in sorted(all_overrides.items()):
        parts.append(f'<Override PartName="{pn}" ContentType="{ct}"/>')

    parts.append('</Types>')
    return "\n".join(parts)
