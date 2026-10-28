---
name: sad-document-generator
description: 'Generate a Solution Architecture Document (SAD) Word file that follows the SBD SAD-Template.pdf structure, from a set of input documents (BRD, kickoff deck, SARB deck, questionnaire, etc.). Rewrites source verbiage into confident, professional, executive-summary-quality prose for every section (never copy-pasted raw/hedgy source text). Produces process/sequence diagrams per BRD requirement (as images plus websequencediagrams.com scripts) and a polished swimlane-style high-level architecture diagram (via the swimlane-flow-diagram skill toolkit) when Ardoq exports are unavailable. Use for creating or drafting a SAD, solution architecture document, or process sequence diagrams for a project.'
argument-hint: 'Provide the input document paths (BRD, kickoff/SARB decks, questionnaires, etc.) and the output .docx path when prompted'
user-invocable: true
disable-model-invocation: false
---

# SAD (Solution Architecture Document) Generator

Produce a complete Word document following the SBD `SAD-Template.pdf` structure (see
[SAD-Template-structure.md](SAD-Template-structure.md) for the full, authoritative section list
and per-section guidance) from a project's input documents.

This skill is content-generation-heavy: the bundled scripts extract and render, but **you (the
model) must read the extracted content and author the actual narrative, tables, and diagram
scripts** for the specific project — there is no generic auto-fill.

## Inputs needed from the user

Ask for (accept sensible defaults / proceed with what's given rather than blocking):
1. Paths to available input documents — BRD, kickoff deck, SARB deck, questionnaire, prior
   architecture notes, or any other project material (.docx/.pptx/.xlsx/.pdf/.txt).
2. The output `.docx` path/filename (default: `<Project Name> - Solution Architecture Document.docx`
   in the same folder as the inputs).
3. Whether Ardoq Level 2 architecture diagram exports are available to paste in directly. If not,
   this skill draws a conceptual substitute instead.

## Writing Standard: Executive-Grade, Confident Narrative

Source documents (questionnaires, BRDs, kickoff notes) are frequently written in a hedgy,
inconsistent, diary-style, or vendor-casual tone — raw dated change-log notes, repeated
qualifiers ("we don't currently... however we intend to...", "Not applicable - see above"),
inconsistent terminology, and informal phrasing. **Never copy that verbiage verbatim into the
SAD.** Every section — not just the Overview — must be rewritten into clear, confident,
professional prose suitable for executive/SARB review, while staying strictly faithful to the
underlying facts (rewrite tone and structure, never invent content).

Apply these rules to every paragraph, bullet, and table cell you author:

- **Active voice, declarative statements.** Replace "We intend to use anonymized data for..."
  with "Anonymized data will be used for...". State the current design decision as a fact, not a
  hope.
- **One clean statement per fact — no stacked hedges.** Collapse "We currently use X, however we
  don't do Y yet, though we may consider Z in future" into: a clear current-state sentence, plus
  (if genuinely relevant) a separate, explicitly labeled "Planned" or "Roadmap" sentence — not a
  single run-on qualifier chain.
- **Strip diary/change-log formatting.** Source "Update" fields with dated journal entries (e.g.
  "July 2: ...; June 12: ...; May 21: ...") must be converted into a single current **Status**
  statement (e.g. "Status: Complete — in production since <date>; retained in PPM for training
  reference.") — never reproduce the raw entry-by-entry log in the SAD body.
- **Normalize repeated "Not applicable" answers.** Where a source document answers many
  questions with "Not applicable — see X above", consolidate them into one clear statement (e.g.
  "AI-specific governance controls are not yet applicable, as the solution uses deterministic
  logic rather than AI/ML models; this will be revisited when AI features are introduced.")
  instead of repeating the deflection per bullet.
- **Consistent terminology.** Pick one term per concept (e.g. "Solution Architect" not SA/Sol
  Arch/solutioning lead interchangeably) and use it throughout the document.
- **No fabrication.** If the source is genuinely silent or ambiguous on a point, say so plainly —
  "Not documented in the source materials; recommend confirming with <owning team>." — rather
  than softening the gap with vague language or guessing.
- **Executive tone, not marketing tone.** Confident and precise, not promotional. Avoid vendor
  marketing phrases ("industry-leading", "seamless", "cutting-edge") unless directly quoting a
  claim that needs attribution.

**Example transformation** (source → SAD):
> Source: *"We currently use commercial LLMs... All LLM outputs undergo human review before use,
> and none of this tooling is customer-facing... We do not currently use AI in our application
> because current models do not meet the necessary standard of correctness... We are building the
> data infrastructure to support AI-powered recommendations in future releases (12-18 month
> timeline)."*

> SAD: *"The production application uses deterministic, rules-based logic exclusively — no AI/ML
> models are used in customer-facing recommendations today. LLMs are used only as internal
> development tools (e.g. code review, documentation drafting), with all output human-reviewed
> and never exposed to customers. AI-powered recommendations are on the roadmap within 12–18
> months, pending model accuracy meeting the required correctness standard; AI governance
> controls (bias audits, explainability, content filtering) will be defined ahead of that
> release and are tracked as an open item until then."*

## Procedure

1. **Extract raw content** from every input document into a working `extract/` folder:
   ```powershell
   py "<skill_dir>\lib\extract_docs.py" --input "<file1>" "<file2>" ... --out-dir "<work_dir>\extract"
   ```
   (or `--input-dir "<folder>"` to sweep an entire folder). Read each resulting `.txt` file.

2. **Identify the high-level business requirements (BRs)** from the BRD content — these drive
   section 2.2 (one sequence diagram per requirement is expected; this is called out in the
   template as "a critical section which needs to be diligently completed").

3. **Draft section 1 content** (Overview, Constraints, Risks/Contingencies) as prose/bullets/table
   based on what the input documents actually say, **rewritten per the Writing Standard above** —
   confident, professional, executive-ready prose, never the source's raw wording. Do not invent
   facts; if information for a sub-section isn't available in the inputs, say so explicitly and
   note it as an open item rather than fabricating detail.

4. **Build the high-level architecture diagram (section 2.1):**
   - If the user has Ardoq exports, note that they should paste those in directly (as images) in
     place of/alongside the conceptual diagram, and skip to step 5.
   - Otherwise, build a polished swimlane-style conceptual diagram using the
     **`swimlane-flow-diagram` skill's toolkit** (same `build_swimlane_pptx.py` /
     `render_and_verify.py` pattern — see that skill for the full helper API and pitfalls to
     avoid). This replaces the older plain matplotlib look with the same professional,
     shadowed-shape, lane-labeled style used for process-flow diagrams:
     - Invoke/reactivate the `swimlane-flow-diagram` skill if not already loaded.
     - Copy `build_swimlane_pptx.py` into `<work_dir>\diagrams\` and customize it for this
       project: lanes = systems/actors/hosting tiers (e.g. Client, Application, Database,
       Integration/Middleware, Infrastructure), boxes = components with their role, connectors =
       data/integration flows labeled with protocol or method (TLS, API, SFTP, etc.).
     - Build one slide for **Current State** and one for **Future State** (or a single slide if
       there is no current-state system to contrast against — common for net-new builds).
     - Run the generator, then the **mandatory self-check**
       (`py render_and_verify.py "<pptx path>"`) to reroute connectors and export PNGs; visually
       inspect each PNG for the same issues called out in that skill (overlapping shapes, hidden
       arrows, clipped text, off-slide content) and fix before proceeding.
     - Clearly caption both diagrams as conceptual substitutes for the Ardoq export, noting they
       should be replaced with the Ardoq Level 2 export prior to SARB/TechPASS submission.
     - Keep the editable `.pptx` as a companion deliverable alongside the `.docx`; embed the
       exported PNG(s) into section 2.1 of the Word document.
   - Only fall back to the simpler `lib\arch_diagram.py` (`new_figure`, `box`, `arrow`, `save`)
     matplotlib helpers if PowerPoint/pywin32 isn't available in the environment to run the
     swimlane self-check — note this limitation to the user if it happens.

5. **Build sequence diagrams (section 2.2):** for each BR identified in step 2, define
   `participants` (systems/actors, left→right) and `messages` (tuples of
   `(from, to, label, style)` where style is `sync`/`return`/`async`), then render with
   `lib\seqrender.py`'s `render_sequence(...)` to a PNG, and also emit the equivalent
   websequencediagrams.com script text (`title ...` + `A->B: label` lines) as a `.wsd.txt` file —
   this gives both a rendered image and an editable fallback per the template's guidance.

6. **Draft section 2.3 details** (2.3.1–2.3.9: Systems, System Connections, Integration
   Approaches, Middleware & Transfer Technologies, Data Objects, Data Strategy, System Hosting,
   System Existence, Deployment & Support) from the extracted content, **rewritten per the Writing
   Standard above** — a systems table (2.3.1) is typically the anchor; the rest can be short
   paragraphs/bullets referencing it, in confident declarative prose rather than restating the
   source Q&A format. Apply the SA standard note in 2.3.6 (integrations should be Delta,
   Asynchronous, Real-Time) unless the inputs state otherwise.

7. **Draft section 2.4** (Interface Inventory) as a table if interface/integration details are
   available in the inputs (table cell text should still follow the Writing Standard — concise,
   consistent terminology); otherwise note that the definitive list should be extracted from the
   Ardoq report and where the master list is maintained.

8. **Assemble the Word document**: write a per-project Python script that imports
   `lib\sad_builder.py`'s `SADDocument` class and calls, in order: `cover_page(...)`, then for
   each template heading from [SAD-Template-structure.md](SAD-Template-structure.md):
   `heading(...)`, `para(...)`/`bullets(...)`/`table(...)`, `add_sequence_diagram(...)` for each BR
   in 2.2, `image(...)` for the architecture diagram(s) in 2.1 (the swimlane-generated PNG(s) from
   step 4), ending with `save("<output path>.docx")`. Run the script with `py`.

9. **Verify**: confirm the output `.docx` opens (e.g. re-open with `python-docx` and print heading
   count) and that every section in [SAD-Template-structure.md](SAD-Template-structure.md) is
   present with real content (no leftover placeholder/lorem text). Spot-check a few paragraphs
   against the Writing Standard — if any read like raw source copy-paste (hedgy, diary-style,
   inconsistent terms), rewrite them before delivering.

10. Report to the user: output `.docx` path, the companion architecture `.pptx` path (from step
    4), list of BRs covered by sequence diagrams, and any sections left as open items due to
    missing source information.

## Notes

- Never fabricate specific names, dates, systems, or figures not present in the input documents —
  flag gaps instead. The Writing Standard governs *tone and structure*, never *content* — rewrite
  freely, invent nothing.
- Keep diagrams readable: wrap long labels, keep to 3–6 participants per sequence diagram where
  possible (split into multiple diagrams for complex flows).
- The architecture diagram (section 2.1) requires the `swimlane-flow-diagram` skill's
  dependencies too: `python-pptx` and `pywin32`, plus a local PowerPoint install for the mandatory
  render/verify self-check. If unavailable, fall back to `lib\arch_diagram.py` and note the
  limitation to the user.
- The bundled generator libraries require Python 3.10+ and the packages in `requirements.txt`.
  Install with `py -m pip install -r requirements.txt`.
