$ErrorActionPreference = "Stop"
$env:DISABLE_SCHEDULER = "1"
$env:MEDICO_SAMPLE_MODE = "1"
$verifyDb = Join-Path (Get-Location) "runtime\verify.sqlite3"
$verifyBackups = Join-Path (Get-Location) "runtime\verify_backups"
if (Test-Path $verifyDb) { Remove-Item -LiteralPath $verifyDb -Force }
if (Test-Path $verifyBackups) { Remove-Item -LiteralPath $verifyBackups -Recurse -Force }
$verifyDbUrlPath = $verifyDb.Replace("\", "/")
$env:DATABASE_URL = "sqlite:///$verifyDbUrlPath"
$env:BACKUP_FOLDER = $verifyBackups

python -m py_compile app.py

@'
from app import app, db, sync_patients_from_medico, PatientRecord, VisitorRecord

app.config["WTF_CSRF_ENABLED"] = False
with app.app_context():
    sync = sync_patients_from_medico(actor="verify")
    assert sync.status == "OK", sync.message
    assert PatientRecord.query.count() >= 1
    if VisitorRecord.query.count() == 0:
        visitor = VisitorRecord(full_name="TEST VISITOR", identity_number="TEST", destination="Gate", status="active")
        db.session.add(visitor)
        db.session.commit()

client = app.test_client()
health = client.get("/health")
assert health.status_code == 200
login = client.post("/login", data={"username": "admin", "password": "admin12345"}, follow_redirects=False)
assert login.status_code == 302, login.status_code
for path in ["/", "/patients", "/visitors", "/scan", "/transit/patients", "/transit/no-admin-discharge", "/audit", "/manage_users", "/manage_backups"]:
    response = client.get(path)
    assert response.status_code == 200, (path, response.status_code)
print("SEC_OFFICE verification OK")
'@ | python -

docker compose -f docker-compose.local.yml config --quiet

if (Test-Path $verifyDb) { Remove-Item -LiteralPath $verifyDb -Force }
if (Test-Path $verifyBackups) { Remove-Item -LiteralPath $verifyBackups -Recurse -Force }
