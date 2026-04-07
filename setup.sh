#!/bin/bash
# ────────────────────────────────────────────────────────────────────────
# Memorial GD — Script de Setup para Deploy
# ────────────────────────────────────────────────────────────────────────
# Uso:
#   chmod +x setup.sh
#   ./setup.sh
#
# Pré-requisitos:
#   - Docker e Docker Compose instalados
#   - Ou: Python 3.10+, LibreOffice, pip
# ────────────────────────────────────────────────────────────────────────

set -e

echo "=========================================="
echo "  Memorial GD — Setup"
echo "=========================================="

# Verificar se Docker está disponível
if command -v docker &>/dev/null; then
    echo ""
    echo "[1/3] Docker encontrado. Construindo imagem..."
    docker build -t memorial-gd .

    echo ""
    echo "[2/3] Criando diretório de dados persistentes..."
    mkdir -p data
    # Copiar equipamentos.json para o volume se não existir
    if [ ! -f data/equipamentos.json ]; then
        cp pipeline/api/equipamentos.json data/equipamentos.json
    fi

    echo ""
    echo "[3/3] Iniciando container..."
    docker compose up -d

    echo ""
    echo "=========================================="
    echo "  Memorial GD rodando em:"
    echo "  http://localhost:8080"
    echo ""
    echo "  Comandos úteis:"
    echo "    docker compose logs -f      # ver logs"
    echo "    docker compose down          # parar"
    echo "    docker compose up -d         # reiniciar"
    echo "=========================================="

else
    echo ""
    echo "Docker não encontrado. Tentando setup local..."
    echo ""

    # Verificar Python
    if ! command -v python3 &>/dev/null; then
        echo "ERRO: Python 3 não encontrado. Instale com: sudo apt install python3 python3-pip"
        exit 1
    fi

    # Verificar LibreOffice
    if ! command -v soffice &>/dev/null; then
        echo "ERRO: LibreOffice não encontrado. Instale com: sudo apt install libreoffice-calc"
        exit 1
    fi

    echo "[1/2] Instalando dependências Python..."
    pip3 install -r requirements.txt

    echo ""
    echo "[2/2] Iniciando servidor..."
    echo ""
    echo "=========================================="
    echo "  Memorial GD rodando em:"
    echo "  http://localhost:8080"
    echo "=========================================="
    echo ""
    python3 pipeline/api/server.py 8080
fi
