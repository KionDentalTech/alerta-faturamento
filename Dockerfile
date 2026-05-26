FROM python:3.11-slim

# Timezone São Paulo
ENV TZ=America/Sao_Paulo
RUN apt-get update && apt-get install -y --no-install-recommends \
        cron tzdata \
    && ln -fs /usr/share/zoneinfo/America/Sao_Paulo /etc/localtime \
    && dpkg-reconfigure -f noninteractive tzdata \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Dependências Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Código e config
COPY scripts/ ./scripts/
COPY config/  ./config/

# Pastas de saída e dados (dados virão via volume)
RUN mkdir -p /app/dados /app/saidas/logs

# Cron: segunda a sexta, 10h horário de Brasília
COPY crontab /etc/cron.d/kion-alerta
RUN chmod 0644 /etc/cron.d/kion-alerta \
    && crontab /etc/cron.d/kion-alerta

# Manter logs do cron visíveis via docker logs
RUN touch /var/log/cron.log

CMD ["sh", "-c", "cron && tail -f /app/saidas/logs/cron.log 2>/dev/null || tail -f /dev/null"]
