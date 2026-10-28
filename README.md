# Solution Architecture Skills

A collection of reusable Copilot skills and supporting scripts for Solution Architecture reporting, governance, and diagramming.

## Skills

- Application Integration Mapper
- Divestiture Separation Plan
- Solution Architect Project Status Review
- SA Engagement Analysis
- SA KPI Heatmap
- SV-Noidia-Solution Architect KPI Performance Heatmap
- SA Effectiveness Scorecard
- SA Workload Dashboard
- SAD Document Generator
- SARB Due Date Heatmap
- SARB SLA Reasoning Analysis
- SARB SLA Report
- Swimlane Flow Diagram Generator
- TechPass Aging Report
- Weekly SA Reports Scheduler

## Use

Each skill folder contains a `SKILL.md` with its purpose, input requirements, and usage steps. Script-based skills also include a `README.md` and the executable script. Shared path helpers are in `_shared`.

The PowerShell launcher (`Invoke-Skill.ps1`) reads `skills-manifest.json` and presents a menu for script-based skills. Some content-authoring skills are assisted workflows and must be run through a Copilot chat or CLI that supports the skill instructions.

## Data handling

Scripts operate on files supplied by the user and write outputs to local folders. Review the input and output paths before running a skill. Do not commit customer exports, generated reports, credentials, or other sensitive data to source control.

The repository is a reusable toolkit, not a source of live Work Queue or integration data. Internal systems require the user's normal authorized access.

## Requirements

Requirements vary by skill. See the skill's `requirements.txt` where present and its README for details. Python-based workflows use Python 3.10 or later; PowerPoint generation and rendering may require a local Microsoft PowerPoint installation.
