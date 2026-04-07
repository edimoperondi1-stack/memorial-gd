"""
lo_export_pdf.py
----------------
Script Python executado DENTRO do LibreOffice via macro.
Recebe parâmetros via variáveis de ambiente.

Fluxo:
  1. Abre o documento
  2. Remove as abas que NÃO estão na lista de abas visíveis
  3. Exporta para PDF
  4. Fecha SEM salvar (preserva o xlsx original)
"""
import os
import sys


def export_pdf():
    """Macro LibreOffice Basic que exporta PDF com apenas as abas selecionadas."""
    # Essa função é chamada como macro StarBasic, não diretamente em Python.
    # Veja _get_macro_content() abaixo para o código que realmente roda.
    pass


# O código abaixo gera a macro StarBasic que será instalada no LibreOffice.
def get_macro_content(abas_visiveis: list, caminho_pdf: str) -> str:
    """
    Gera o conteúdo da macro LibreOffice Basic.

    A macro:
      1. Deleta todas as abas que NÃO estão na lista (mais confiável que ocultar)
      2. Exporta para PDF via storeToURL
      3. Fecha sem salvar (o xlsx original permanece intacto)
    """
    from pathlib import Path
    pdf_url = Path(caminho_pdf).absolute().as_uri()

    # Criar lista de abas a manter
    # IMPORTANTE: nomes de abas com < > & precisam ser escapados para XML
    # pois o .xba é um arquivo XML. O LibreOffice desescapa ao carregar a macro.
    import xml.sax.saxutils
    abas_basic = "\n".join(
        f'        aKeep({i}) = "{xml.sax.saxutils.escape(a)}"'
        for i, a in enumerate(abas_visiveis)
    )

    return f'''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
Sub ExportPDF()
    Dim oDoc As Object
    Dim oSheets As Object
    Dim nSheets As Long
    Dim i As Long
    Dim j As Long
    Dim sName As String
    Dim bKeep As Boolean

    oDoc = ThisComponent
    oSheets = oDoc.getSheets()

    ' Lista de abas a manter
    Dim aKeep({len(abas_visiveis) - 1}) As String
{abas_basic}

    ' Primeiro: tornar todas as abas visiveis (necessario antes de deletar)
    nSheets = oSheets.getCount()
    For i = 0 To nSheets - 1
        oSheets.getByIndex(i).isVisible = True
    Next i

    ' Deletar abas que nao estao na lista (de tras pra frente para nao mudar indices)
    For i = nSheets - 1 To 0 Step -1
        sName = oSheets.getByIndex(i).getName()
        bKeep = False
        For j = 0 To UBound(aKeep)
            If sName = aKeep(j) Then
                bKeep = True
                Exit For
            End If
        Next j
        If Not bKeep Then
            oSheets.removeByName(sName)
        End If
    Next i

    ' Exportar para PDF
    Dim aArgs(1) As New com.sun.star.beans.PropertyValue
    aArgs(0).Name = "FilterName"
    aArgs(0).Value = "calc_pdf_Export"
    aArgs(1).Name = "Overwrite"
    aArgs(1).Value = True

    oDoc.storeToURL("{pdf_url}", aArgs())

    ' Fechar sem salvar (nao modifica o xlsx original)
    oDoc.setModified(False)
    oDoc.close(True)
End Sub
</script:module>'''
