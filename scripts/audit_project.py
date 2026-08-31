from __future__ import annotations

from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent


def require(condition: bool, message: str, errors: list[str]) -> None:
    if not condition:
        errors.append(message)


def read_file(relative_path: str, errors: list[str]) -> str:
    path = ROOT / relative_path
    if not path.is_file():
        errors.append(f"missing {relative_path}")
        return ""
    return path.read_text(encoding="utf-8")


def main() -> int:
    errors: list[str] = []

    required_files = (
        ".github/copilot-instructions.md",
        "app/main.py",
        "public/index.html",
        "tests/test_app.py",
    )
    for relative_path in required_files:
        require((ROOT / relative_path).is_file(), f"missing {relative_path}", errors)

    dockerfile = read_file("Dockerfile", errors)
    require("uvicorn app.main:app" in dockerfile, "Dockerfile must run app.main:app", errors)

    env_example = read_file(".env.example", errors)
    for variable in ("DATABASE_URL", "SUBMISSION_CODE", "ADMIN_PASSWORD", "PUBLIC_URL", "PORT"):
        require(f"{variable}=" in env_example, f".env.example missing {variable}", errors)

    workflow = read_file(".github/workflows/ci.yml", errors)
    require("python -m pytest -q" in workflow, "CI must run the test suite", errors)
    require("docker build" in workflow, "CI must build the production image", errors)
    require("python scripts/audit_project.py" in workflow, "CI must run the project audit", errors)

    if errors:
        for error in errors:
            print(f"FAIL: {error}")
        return 1

    print("Project delivery guardrails passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
