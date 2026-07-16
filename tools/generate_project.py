from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

try:
    import yaml
except ImportError:
    print(
        "ERROR: Falta la dependencia PyYAML.\n"
        "Instálala ejecutando:\n"
        "py -m pip install pyyaml"
    )
    sys.exit(1)


ROOT_DIR = Path(__file__).resolve().parent.parent
PROJECT_FILE = ROOT_DIR / "project.yml"
README_FILE = ROOT_DIR / "README.md"
CITATION_FILE = ROOT_DIR / "CITATION.cff"
PROJECT_INFO_FILE = ROOT_DIR / "docs" / "project-information.md"


def load_project_config() -> dict[str, Any]:
    if not PROJECT_FILE.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo de configuración: {PROJECT_FILE}"
        )

    with PROJECT_FILE.open("r", encoding="utf-8") as file:
        data = yaml.safe_load(file)

    if not isinstance(data, dict):
        raise ValueError("El archivo project.yml no contiene una estructura válida.")

    project = data.get("project")

    if not isinstance(project, dict):
        raise ValueError(
            "El archivo project.yml debe contener una sección principal 'project'."
        )

    required_fields = [
        "name",
        "description",
        "version",
        "organization",
        "author",
        "license",
        "github",
    ]

    missing_fields = [
        field
        for field in required_fields
        if not project.get(field)
    ]

    if missing_fields:
        raise ValueError(
            "Faltan campos obligatorios en project.yml: "
            + ", ".join(missing_fields)
        )

    return project


def value_or_pending(value: Any) -> str:
    if value is None:
        return "Pending"

    text = str(value).strip()
    return text if text else "Pending"


def yaml_boolean(value: Any) -> str:
    return "true" if bool(value) else "false"


def build_badges(project: dict[str, Any]) -> str:
    github = str(project["github"]).rstrip("/")
    repository_path = github.removeprefix("https://github.com/")

    version = str(project["version"])
    license_name = str(project["license"])
    python_version = str(project.get("python", ">=3.11")).replace(">", "%3E")

    return "\n".join(
        [
            (
                f"[![Version](https://img.shields.io/badge/"
                f"version-{version}-blue.svg)]({github}/releases)"
            ),
            (
                f"[![Python](https://img.shields.io/badge/"
                f"python-{python_version}-blue.svg)]"
                f"(https://www.python.org/)"
            ),
            (
                f"[![License](https://img.shields.io/badge/"
                f"license-{license_name}-green.svg)]"
                f"({github}/blob/main/LICENSE)"
            ),
            (
                f"[![GitHub issues](https://img.shields.io/github/issues/"
                f"{repository_path})]({github}/issues)"
            ),
        ]
    )


def build_readme(project: dict[str, Any]) -> str:
    name = str(project["name"])
    description = str(project["description"])
    version = str(project["version"])
    status = value_or_pending(project.get("status"))
    organization = str(project["organization"])
    author = str(project["author"])
    license_name = str(project["license"])
    github = str(project["github"])
    documentation = value_or_pending(project.get("documentation"))
    releases = value_or_pending(project.get("releases"))
    issues = value_or_pending(project.get("issues"))
    python_version = value_or_pending(project.get("python"))

    scientific = project.get("scientific", {})
    outputs = project.get("outputs", {})
    keywords = project.get("keywords", [])

    guideline = value_or_pending(scientific.get("guideline"))
    validation = value_or_pending(scientific.get("validation"))
    peer_review = "Yes" if scientific.get("peer_review") else "No"

    enabled_outputs = [
        output.upper()
        for output, enabled in outputs.items()
        if enabled
    ]

    output_text = ", ".join(enabled_outputs) if enabled_outputs else "Pending"

    keyword_lines = "\n".join(
        f"- {keyword}"
        for keyword in keywords
    )

    badges = build_badges(project)

    return f"""# {name}

{badges}

## Overview

{description}

This project is developed by **{organization}** and provides a reproducible
clinical workflow for the analysis, classification and reporting of resting
blood pressure measurements.

## Current version

- **Version:** {version}
- **Status:** {status}
- **Python:** {python_version}
- **License:** {license_name}

## Scientific framework

- **Clinical guideline:** {guideline}
- **Validation status:** {validation}
- **Peer reviewed:** {peer_review}

The software is intended to support structured blood pressure analysis and
report generation. It does not replace clinical assessment, diagnosis or
medical decision-making by qualified healthcare professionals.

## Main outputs

{output_text}

## Installation

Clone the repository:

```powershell
git clone {github}.git
cd {name}