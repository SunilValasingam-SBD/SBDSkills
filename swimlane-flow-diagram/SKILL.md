---
name: swimlane-flow-diagram
description: 'Generate a professional, C-level-ready swimlane process flow diagram as an editable PowerPoint (.pptx) using python-pptx native shapes and connectors, with a mandatory visual self-check step. Use when the user describes a multi-step business/technical process across multiple actors or systems (e.g. status/stage x actor tables) and wants a clear flow diagram for both developers and executives.'
argument-hint: '[process steps/table, list of lanes/actors]'
user-invocable: true
disable-model-invocation: false
---

# Swimlane Flow Diagram Generator

Build a polished swimlane flow diagram (one lane per actor/system, steps flowing
left-to-right in chronological order) as a native, editable PowerPoint file — not a static
image and not a matplotlib chart. PowerPoint's own shape/text/connector engine avoids the
text-overflow and crooked-line problems that hand-drawn (matplotlib) diagrams are prone to,
*provided* the known pitfalls below are respected.

## Use When

- The user gives you a process broken into steps, each attributable to an actor/system/lane
  (e.g. a table of "Status x System" steps, a swimlane BPMN-style description, a runbook).
- The user wants something presentable to both developers (accurate detail) and executives
  (clean, uncluttered, professional styling).
- The user wants an editable deliverable they can tweak later, not a flattened image.

## Procedure

1. Parse the user's process description into an ordered list of steps per lane, including
   any decision points (branches), stop conditions, and success/error outcomes.
2. Copy `build_swimlane_pptx.py` into the working output folder and edit it for this
   specific process:
   - Rename/re-color `LANE_FILL` / `LANE_TAB` for the actual lanes.
   - Replace `build_example_slide()` with one function per slide/scenario, using
     `add_box()` for process/decision/stop/success/error nodes and `connect()` for
     arrows between them, following the spacing rules below.
   - One slide per distinct scenario/status if the user's process has variants (e.g.
     "Prelim" vs "WIP"), all using the same lane/style conventions for consistency.
3. Run the script (`py build_swimlane_pptx.py`) to generate the `.pptx`.
4. **Mandatory self-check**: run
   `py render_and_verify.py "<path to the generated .pptx>"`
   (requires `pywin32`: `py -m pip install pywin32`, and a local PowerPoint install). This
   reroutes every connector cleanly and exports each slide as a PNG.
5. View the exported PNG(s). Check specifically for:
   - Arrows that are missing/invisible where two shapes are close together or stacked
     across adjacent lanes (usually means zero/negative gap — see Spacing below).
   - Text spilling past a shape's border (box undersized for its wrapped line count).
   - Any shape or text clipped by the slide edge (widen `prs.slide_width` if so).
   - Arrowheads pointing the wrong way or a line passing straight through a shape's text
     (this means a connection-site mapping regressed — verify against the NOTES in
     `build_swimlane_pptx.py`, don't guess).
6. Fix anything found, re-run steps 3–5 until clean, then deliver the `.pptx` (and,
   optionally, the exported PNGs as quick-look previews) to the user's requested location.
7. Clean up temporary scripts/crops from the working folder; keep only the final `.pptx`
   and preview PNGs as deliverables.

## Spacing Rules (the most common source of visible bugs)

- Leave at least ~0.3–0.4in of clear space between any two shapes that will be directly
  connected, especially vertically adjacent lanes (e.g. a decision diamond in one lane and
  an outcome box in the lane above/below it).
- If you enlarge a shape (e.g. widen/heighten a decision diamond so its text fits), also
  grow that shape's lane height by a matching amount — otherwise the enlarged shape can end
  up touching its neighbor, producing a zero-length, invisible connector. This exact
  regression happened once: enlarging diamonds for text-fit fixed one problem but caused
  boxes to visually touch with no visible arrow. Always re-render after resizing anything.
- Compute box height from the number of wrapped text lines, not a fixed guess.

## Reusable Prompt (paste this, filling in your process, if invoking manually)

> Create a professional, C-level-ready swimlane process flow diagram as an editable
> PowerPoint (.pptx) file (not a static image), using Python's python-pptx library.
> Requirements:
> 1. Horizontal swimlanes, one per actor/system, each with a colored tab label on the left.
>    Steps flow left-to-right chronologically based on which actor performs each step.
> 2. Rounded rectangles for process steps, diamonds for decisions, distinctly colored
>    rounded rectangles for terminal outcomes (green=success, red=error, gray=stop).
> 3. Elbow connectors bound to shape connection sites, with visible arrowheads and Yes/No
>    branch labels.
> 4. Dynamically sized boxes to fit wrapped text; generous gaps (>= 0.3-0.4in) between any
>    two vertically-stacked connected shapes so arrows are never hidden or overlapping.
> 5. After generating, use PowerPoint COM automation to reroute every connector, save, export
>    each slide to PNG, and visually inspect before calling it done — check for overlapping
>    shapes, missing/hidden arrowheads, clipped text, and off-slide content.
> 6. Clean corporate palette, Calibri font, subtle shadows, title + subtitle + separator
>    line, and a legend (place at the bottom if lanes need left-side space).
>
> Here is my process data: [paste steps/table here]

## Notes

- Requires `python-pptx` and `pywin32` (`py -m pip install python-pptx pywin32`) plus a
  local PowerPoint installation for the render/verify step.
- See the inline `NOTES` section at the bottom of `build_swimlane_pptx.py` for the specific
  connection-site index mapping and other empirically-verified fixes — don't rediscover
  these from scratch.
