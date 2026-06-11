# MMFI 251 ΓΝΑ

Εφαρμογή παρακολούθησης οικονομικής τακτοποίησης χειρουργικών επεμβάσεων.

## Λειτουργίες v1.0

- Καρτέλα χειρουργείου με ημερομηνία, περιγραφή, υπεύθυνο ιατρό, προμηθευτή και σημειώσεις.
- Αυτόματη δημιουργία workflow 13 βημάτων σύμφωνα με την απαίτηση των χρηστών.
- Ολοκλήρωση βήματος με αυτόματη καταγραφή ημερομηνίας, χρήστη και σχολίου.
- Δικαιώματα ανά σταθμό εργασίας για τοπικούς fallback χρήστες και SSO groups.
- Αναφορές με φίλτρα και εξαγωγή Excel.
- Audit log.
- Κρυπτογραφημένα backup με manifest, SHA-256, verify και download.
- Κοινό SSO μοντέλο μέσω `X-SSO-*` headers, με local fallback login.

## Local run

```powershell
docker compose -f docker-compose.local.yml up -d --build
```

Η εφαρμογή ανοίγει στο `http://localhost:5050`.

Default local fallback values are read from the local `.env`, which is intentionally not committed.

Αλλάξτε τα production secrets στο `.env` πριν από παραγωγική λειτουργία.

## SSO groups

- `/apps/mmfi/users`: πρόσβαση εφαρμογής.
- `/apps/mmfi/admins`: διαχειριστές εφαρμογής.
- `/apps/global/admins`: κοινοί διαχειριστές.
- `/apps/mmfi/stations/<station_code>`: δικαίωμα ολοκλήρωσης ενεργειών συγκεκριμένου σταθμού.

Station codes:

- `surgery_secretariat`
- `isupply_committee`
- `orders_office`
- `mef`
- `budget_accounting`
- `small_procurement`
- `receiving_committee`
- `finance_office`

## Production notes

Production discipline follows the shared intranet pattern:

- Git remote: `git@github.com:itc251gna/MMFI_OR.git`.
- Deploy only an explicit release tag or commit SHA that is contained in `origin/main`.
- Do not deploy mutable refs such as `main`, `origin/main`, or `HEAD`.
- Keep production `.env`, runtime PostgreSQL data, backups and uploads outside Git.
- Canonical production deployment folder: `/home/kmh251/deployment/mmfi`.
- Canonical production URL through `app-gateway-nginx`: `https://mmfi.251gh.local/`.

Create a local release tag:

```powershell
.\scripts\create_release.ps1 -Tag mmfi-vYYYY-MM-DD-name -CommitMessage "Release MMFI ..." -Push
```

Deploy on the production VM from the production checkout:

```bash
cd /home/kmh251/deployment/mmfi
./scripts/deploy_production.sh mmfi-vYYYY-MM-DD-name
```

Το `docker-compose.remote.yml` εκθέτει μόνο το `mmfi-app` στο εσωτερικό Docker network. Το TLS και το SSO boundary ανήκουν στο κεντρικό app gateway, όπως στις υπάρχουσες εφαρμογές.
