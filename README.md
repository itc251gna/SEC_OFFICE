# SEC OFFICE 251 GNA

Flask/PostgreSQL intranet application for replacing the legacy `HmerKatNosil.aspx` daily security-office reports.

## Core behavior

- Syncs the daily hospitalized-patient list from MEDICO `HMERKATNOSIL`.
- Stores every patient row internally and preserves manual fields such as the two escorts.
- Updates only MEDICO-owned fields when external data changes.
- Deletes patients from the internal active list when a successful MEDICO sync no longer returns them.
- Provides persistent visitor/person-access records with authorization document metadata and optional file upload.
- Generates QR labels for patients and visitors and records scan audit events.
- Keeps the old movement and no-administrative-discharge tabs as live read-only MEDICO views.
- Uses local fallback login and supports the shared SSO header contract.
- Includes encrypted backup creation, verification and retention.

## Local run

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

Local development can run with `MEDICO_SAMPLE_MODE=1` and listens on `http://localhost:5055`.

Fallback users in the local `.env`:

- `admin` / `admin12345`
- `security_office` / `security12345`
- `transit_office` / `transit12345`
- `security_point` / `point12345`

## Docker local

```powershell
docker compose -f docker-compose.local.yml up --build
```

## Production

Production follows the common intranet application pattern:

- Local development stays in this repository with SQLite/sample mode or direct MEDICO credentials in a private `.env`.
- Remote production runs with `docker-compose.remote.yml`, PostgreSQL, the shared SSO gateway and Keycloak groups.
- Secrets stay in the remote `.env` and are not committed.
- Runtime data stays under `runtime/` and `static/uploads/` and is not committed.

The expected production URL is:

```text
https://sec-office.251gh.local/
```

Copy `.env.example` to `.env`, set strong secrets, set `MEDICO_SAMPLE_MODE=0`, provide MEDICO Oracle credentials through environment variables, and set `POSTGRES_PASSWORD` for `docker-compose.remote.yml`.

The MEDICO database currently requires python-oracledb thick mode because it is Oracle 10g. Place the Linux x86-64 Oracle Instant Client 11.2 Basic package contents under:

```text
runtime/oracle/instantclient/
```

Place `tnsnames.ora` under:

```text
runtime/oracle/network/admin/tnsnames.ora
```

The remote compose file maps those directories to:

```text
ORACLE_CLIENT_LIB_DIR=/opt/oracle/instantclient
ORACLE_CONFIG_DIR=/app/runtime/oracle/network/admin
TNS_ADMIN=/app/runtime/oracle/network/admin
```

SSO groups are configurable with the `SSO_*` variables and match the common gateway header flow. The default groups are:

```text
/apps/sec-office/users
/apps/sec-office/admins
/apps/sec-office/security-office
/apps/sec-office/transit-office
/apps/sec-office/security-point
/apps/global/admins
```
