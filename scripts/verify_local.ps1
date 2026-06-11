$ErrorActionPreference = "Stop"

python -m py_compile app.py
docker compose -f docker-compose.local.yml config --quiet
docker compose -f docker-compose.remote.yml config --quiet

Write-Host "MMFI local verification OK"
