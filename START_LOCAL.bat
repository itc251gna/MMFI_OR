@echo off
docker compose -f docker-compose.local.yml up -d --build
echo MMFI is available at http://localhost:5050
