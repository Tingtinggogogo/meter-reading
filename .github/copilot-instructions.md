# Meter Reading Repository Instructions

## Source and architecture

- GitHub is the only application source of truth. Never reconstruct or redeploy source files extracted from a running container.
- Preserve the root `Dockerfile` as the Zeabur build contract unless the user explicitly approves a build-system change.
- The application consists of `public/index.html`, the FastAPI service in `app/main.py`, PostgreSQL, Excel exports, and QR endpoints.
- Never place real database URLs, submission codes, administrator passwords, or other secrets in the repository, logs, URLs, or exports.

## Business rules

- Site identifier is `STO401`.
- Users enter raw meter-dial readings.
- `boka` and `huaman` high-voltage readings use multiplier `3000`.
- `solar1` through `solar4` have no transformation and use multiplier `1`.
- `charger1` through `charger3` have no transformation and use multiplier `1`.
- Water readings have no transformation.
- Electricity is measured in kWh; water is measured in metric tons.
- The server's `Asia/Shanghai` date and time are authoritative.
- One site record is retained per day; another submission on the same day updates that record.
- Export month choices may include the current server month and past months, never future months.
- Excel detail rows and cached formula results must follow the same multiplier rules.

## Change workflow

- Never commit directly to `main`.
- Create a focused branch and pull request.
- Require the `test-and-build` check before merging.
- Keep unrelated changes out of the pull request.
- Use staging with an isolated PostgreSQL database for risky behavior, schema, authentication, export, or deployment changes.
- Explain impact and rollback before deleting data, services, domains, volumes, or databases.

## Required verification

- Run `python scripts/audit_project.py`.
- Run `python -m pytest -q`.
- Build the production Docker image in CI.
- Add or update a regression test for every behavior change.
- Scan all related surfaces before changing a business rule: backend configuration, UI copy, Excel generation, tests, and README.
- Verify the deployed commit SHA, `/api/health`, critical UI text, API behavior, and export result.
- Treat CLI success as provisional; read provider state back after infrastructure mutations.

## Mobile compatibility

- Design for current iOS Safari, Android Chrome/system browsers, HarmonyOS, and common embedded WebViews.
- Prefer standards-based navigation and signed GET file downloads.
- Test direct links, QR targets, form controls, responsive width, submission, queries, and Excel download.
- Do not claim compatibility with an affected device/browser without evidence from that path.
- Separate application defects from DNS, TLS, CDN, carrier routing, and organization security filtering.

## Data and recovery

- Never point staging or preview deployments at the production database.
- Back up production data before risky database operations.
- Excel exports are business artifacts, not complete PostgreSQL backups.
- Roll back application code by deploying a known-good Git commit with the same Dockerfile.
