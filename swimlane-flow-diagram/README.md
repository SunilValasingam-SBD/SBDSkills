# Swimlane Flow Diagram Generator

## Executive Summary

Process and workflow diagrams are often needed quickly to explain how a status, approval, or system handoff actually works across multiple teams or systems — but hand-drawn or ad-hoc diagrams frequently look unprofessional, with crooked lines, missing arrows, or overlapping text that undermine credibility in front of leadership. This skill turns a plain description of a business process into a polished, presentation-quality swimlane diagram, with each team or system shown in its own clearly labeled lane and each step connected by clean, visible arrows. Because it is built as an editable file rather than a flattened picture, project teams can keep refining it themselves after delivery. The result is a reliable, C-level-ready visual that development and executive audiences can both understand at a glance.

## Business Process Flow

1. The requester describes the process to be diagrammed — the ordered steps, which actor/system/team performs each step (the "lanes"), and any decision points, branches, or end states (success/error/stop).
2. The tool organizes the process into horizontal lanes (one per actor/system) with steps flowing left-to-right in the order they happen, using distinct shapes for regular process steps, decisions, and terminal outcomes (e.g. green for success, red for error).
3. It generates the diagram as an editable PowerPoint file, not a static picture, so it can be opened and adjusted later like any other presentation slide.
4. Before delivering anything, the tool automatically opens the generated file, straightens every connecting line, exports a picture of each slide, and inspects those pictures for common defects — hidden arrows, text spilling outside a box, or content cut off the edge of the slide — fixing anything it finds and re-checking until the result is clean.
5. If the process has multiple variants or scenarios (e.g. different statuses), a separate, consistently styled slide is produced for each one.
6. The requester receives the final editable PowerPoint (plus quick-look preview images of each slide) ready to drop into a deck or walk through live with stakeholders.

## Technical Flow

This is a model-driven ("assisted") skill built around two bundled helper scripts/toolkit files rather than one fixed pipeline:

1. The model parses the user's process description into an ordered list of steps per lane, including decision points, branch conditions, and terminal outcomes.
2. The model copies `build_swimlane_pptx.py` (the diagram-building toolkit) into the working output folder and customizes it for the specific process: renaming/recoloring `LANE_FILL`/`LANE_TAB` for the actual lanes, and replacing the example slide function with one function per slide/scenario that calls the toolkit's `add_box()` helper (for process/decision/stop/success/error nodes, built on native `python-pptx` shapes with shadowing and dynamic sizing based on wrapped text line count) and `connect()` helper (elbow connectors bound to shape connection sites with arrowheads and branch labels), following the toolkit's documented spacing rules (≥0.3–0.4in gaps between directly connected shapes, re-rendering after any shape resize).
3. The model runs the customized script (`py build_swimlane_pptx.py`) to generate the `.pptx`.
4. **Mandatory self-check**: the model runs `render_and_verify.py "<path to the generated .pptx>"`, which uses PowerPoint COM automation (`pywin32`) to open the file, force every connector shape to reroute cleanly (`RerouteConnections`), save the file, and export each slide as a high-resolution PNG.
5. The model visually inspects each exported PNG for missing/hidden arrows (usually a zero/negative lane gap), text overflowing a shape's border, content clipped by the slide edge, or mis-pointed arrowheads/wrong connection-site mapping (checked against the empirically-verified `NOTES` section at the bottom of `build_swimlane_pptx.py`).
6. The model fixes anything found and repeats steps 3–5 until the rendered output is clean, then delivers the final `.pptx` (and optionally the preview PNGs) to the requested location, cleaning up any temporary working files afterward.

## Inputs

- A description of the business/technical process: ordered steps, the actor/system/lane responsible for each step, and any decision points, branches, or terminal outcomes.
- Optional: desired output location/filename for the final `.pptx`.

## Outputs

- **`<diagram name>.pptx`** — the final editable swimlane process flow diagram, with one slide per distinct process scenario/variant.
- **`slide_<n>.png`** — one high-resolution preview image per slide, exported during the mandatory self-check for visual verification and quick-look sharing.
