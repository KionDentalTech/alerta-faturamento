#!/bin/bash
# ============================================================
#  Kion Dental — Setup inicial do servidor Linux
#  Executar UMA VEZ como root no servidor de produção
# ============================================================

set -e

echo "==> Instalando dependências do sistema..."
apt-get update && apt-get install -y docker.io docker-compose-v2 cifs-utils

echo "==> Criando pasta de montagem do share..."
mkdir -p /mnt/kion-storage
mkdir -p /var/log/kion

# ── Montar o share de rede ────────────────────────────────────
# Preencha USUARIO e SENHA com as credenciais de domínio
# que têm acesso ao \\kion-ad\KionStorage
USUARIO="PREENCHA_USUARIO_AD"
SENHA="PREENCHA_SENHA_AD"
DOMINIO="kion-ad"

echo "==> Montando \\\\kion-ad\\KionStorage em /mnt/kion-storage..."
mount -t cifs //kion-ad/KionStorage /mnt/kion-storage \
    -o username="${USUARIO}",password="${SENHA}",domain="${DOMINIO}",\
uid=1000,gid=1000,iocharset=utf8,vers=3.0

# ── Montar automaticamente no boot ───────────────────────────
FSTAB_ENTRY="//kion-ad/KionStorage /mnt/kion-storage cifs username=${USUARIO},password=${SENHA},domain=${DOMINIO},uid=1000,gid=1000,iocharset=utf8,vers=3.0,_netdev 0 0"

if ! grep -q "kion-ad/KionStorage" /etc/fstab; then
    echo "${FSTAB_ENTRY}" >> /etc/fstab
    echo "==> Entrada adicionada ao /etc/fstab (monta no boot)"
fi

# ── Verificar conteúdo do share ──────────────────────────────
echo ""
echo "==> Conteúdo de /mnt/kion-storage/ANALYTICS/dados:"
ls -lh /mnt/kion-storage/ANALYTICS/dados/ 2>/dev/null || \
    echo "ATENÇÃO: pasta não encontrada — verifique o caminho no share"

# ── Subir o container ────────────────────────────────────────
echo ""
echo "==> Iniciando container Docker..."
cd /opt/kion-alerta
docker compose up -d --build

echo ""
echo "============================================================"
echo "  Setup concluído!"
echo "  Verifique os logs: docker logs -f kion-alerta-faturamento"
echo "============================================================"
