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

Production deployment status as of 2026-06-11:

```text
Host: linuxsrv01 / 10.4.51.232
Folder: /home/kmh251/deployment/sec_office
Gateway: /home/kmh251/deployment/app_gateway/nginx.conf.sso-phase1p-mmfi
Git commit: see GitHub main
```

The app, PostgreSQL container, Keycloak groups, gateway route, SSO-first app login, Keycloak-first user management, and live MEDICO bridge are deployed. Hostname use requires DNS:

```text
sec-office.251gh.local -> 10.4.51.232
```

Copy `.env.example` to `.env`, set strong secrets, set `MEDICO_SAMPLE_MODE=0`, and set `POSTGRES_PASSWORD` for `docker-compose.remote.yml`.

Production MEDICO access uses the legacy IIS bridge endpoint in `legacy_bridge/SecOfficeMedicoBridge.ashx`, deployed under the existing DailyReports application. Set:

```text
MEDICO_BRIDGE_URL=http://10.4.55.149/portal_services/SecOfficeMedicoBridge.ashx
```

The bridge executes the same legacy MEDICO queries through the existing `ConnectionString2` provider and is restricted to requests from `linuxsrv01` (`10.4.51.232`). Direct Linux Oracle thick mode remains available for future use if a compatible Oracle client is installed; leave `MEDICO_BRIDGE_URL` empty to use `MEDICO_DSN`, `MEDICO_USER`, and `MEDICO_PASSWORD`.

SSO groups are configurable with the `SSO_*` variables and match the common gateway header flow. The default groups are:

```text
/apps/sec-office/users
/apps/sec-office/admins
/apps/sec-office/security-office
/apps/sec-office/transit-office
/apps/sec-office/security-point
/apps/global/admins
```

In production SSO sessions, local fallback user create/update/delete is disabled by default. Keep `ALLOW_LOCAL_USER_ADMIN_FROM_SSO=0` unless deliberately opening an emergency recovery path; normal users and rights are managed in Keycloak groups.
