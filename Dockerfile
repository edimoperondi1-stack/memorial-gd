# ────────────────────────────────────────────────────────────────────────
# Memorial GD — Pipeline de Geração de Documentos (Energisa)
# ────────────────────────────────────────────────────────────────────────
# Base: Python 3.11 slim + LibreOffice headless
# Build: docker build -t memorial-gd .
# Run:   docker run -p 8080:8080 memorial-gd
# ────────────────────────────────────────────────────────────────────────

FROM python:3.11-slim

# Evitar prompts interativos durante a instalação
ENV DEBIAN_FRONTEND=noninteractive

# ── Instalar LibreOffice, gcc (para socket shim) e dependências ───────
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    libreoffice-core \
    fonts-liberation \
    fonts-dejavu-core \
    gcc \
    libc6-dev \
    coreutils \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# ── Configurar ambiente LibreOffice headless ──────────────────────────
ENV SAL_USE_VCLPLUGIN=svp
ENV HOME=/app

# ── Copiar código ─────────────────────────────────────────────────────
WORKDIR /app

# Requirements primeiro (cache de camada Docker)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Template Excel (arquivo base do memorial)
COPY "MEMORIAL_GD_v4-22022022 (54).xlsx" .

# Pipeline Python
COPY pipeline/ pipeline/

# Scripts auxiliares do LibreOffice (soffice helper, recalc)
# O pipeline referencia /sessions/.../skills/xlsx/scripts/ — vamos
# copiar para um local interno e ajustar o path
COPY lo_scripts/ /app/lo_scripts/

# ── Inicializar LibreOffice (cria diretórios de config) ───────────────
RUN soffice --headless --terminate_after_init 2>/dev/null || true
RUN mkdir -p /app/.config/libreoffice/4/user/basic/Standard

# ── Expor porta ───────────────────────────────────────────────────────
EXPOSE 8080

# ── Comando de inicialização ──────────────────────────────────────────
CMD ["python", "pipeline/api/server.py", "8080"]
