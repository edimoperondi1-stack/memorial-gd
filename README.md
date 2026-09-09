---
title: Memorial GD Solar
emoji: ☀
colorFrom: yellow
colorTo: red
sdk: docker
app_port: 8080
pinned: false
---

# Memorial GD Solar

Pipeline de geracao de documentos tecnicos para projetos de Geracao Distribuida (GD) solar, conforme requisitos da Energisa (MT).

## Como usar

Acesse a interface web e preencha o formulario com os dados do projeto. O sistema gera automaticamente o memorial de calculo e os documentos necessarios.

## Ferramentas

- **Dimensionamento de Strings:** Calcula a configuracao de strings fotovoltaicas com verificacoes de MPPT, Voc corrigida para T_min e Isc_total para dimensionamento de protecoes.
- **Dimensionamento de Condutores CA:** Calcula a bitola minima de condutores conforme NBR 5410 pelos criterios de corrente admissivel e queda de tensao.
- **Localizacao do Projeto:** Mapa de satelite para enquadrar o local. Gera um PNG com a coordenada UTM e a bussola desenhadas, pronto para colar na prancha do DWG. A banda do fuso e calculada a partir da latitude. Requer as chaves do Google Maps (ver Configuracao).

## Stack

- Python 3.11 + LibreOffice headless (conversao XLSX para PDF)
- Frontend HTML/JS puro (sem framework)
- Backend: SimpleHTTPRequestHandler (stdlib Python)

## Configuracao

A ferramenta de localizacao usa a API do Google Maps e precisa de duas chaves,
definidas em Settings > Variables and secrets do Space:

| Variavel | Para que serve | Restricao recomendada |
|---|---|---|
| `GOOGLE_MAPS_BROWSER_KEY` | Maps JavaScript API, usada pelo mapa no navegador. E necessariamente publica: aparece no HTML da pagina. | Referrer HTTP, limitado ao dominio do Space |
| `GOOGLE_MAPS_API_KEY` | Maps Static API, usada pelo servidor para gerar o PNG. Nunca chega ao navegador. | Restricao por API |

Sem elas a pagina `/localizacao` carrega e explica o que falta, sem quebrar o resto do app.

## Desenvolvimento local

```bash
docker compose up --build
# Acesse http://localhost:8080
```

## Notas

O armazenamento no Hugging Face Spaces e ephemero: arquivos gerados (PDFs, XLSXs de clientes) sao perdidos ao reiniciar o container. O design da aplicacao e compativel com isso pois os arquivos sao baixados imediatamente apos a geracao.

O arquivo `MEMORIAL_GD_v4-22022022 (54).xlsx` e o template base do pipeline — nao e dado de cliente — e necessario para o build da imagem Docker.
