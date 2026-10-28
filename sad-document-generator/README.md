# SAD Document Generator

## Executive Summary

Writing a Solution Architecture Document today means a Solution Architect manually synthesizing BRDs, kickoff decks, SARB decks, and questionnaire answers into a single polished, SARB-ready narrative — a time-consuming, inconsistent process where quality depends heavily on who's writing it. This skill turns that raw project material into a complete, professionally written Solution Architecture Document automatically, following the official company template, with every section rewritten into confident, executive-grade prose rather than copied from informal source notes. It also generates the required process/sequence diagrams and a conceptual architecture diagram so the document is submission-ready. The result is faster SAD production, consistent quality and terminology across every project, and SARB reviewers reading a clear, decision-ready document instead of a raw assembly of source notes.

## Business Process Flow

1. The Solution Architect gathers whatever project material already exists — the BRD, kickoff deck, SARB deck, intake questionnaire, prior architecture notes — in whatever format they're in (Word, PowerPoint, Excel, PDF, plain text).
2. The Solution Architect provides those file paths (and, if available, exported architecture diagrams from the company's architecture repository) along with a desired output filename.
3. The tool reads and extracts the raw content of every document provided.
4. It identifies the business requirements described in the BRD and determines how many process/sequence diagrams are needed (one per requirement).
5. It drafts every required section of the official SAD template — summary, constraints, risks, architecture diagrams, sequence diagrams, solution design details, interface inventory — rewriting any raw, hedgy, or diary-style source language into clear, confident, decision-ready business prose. It never invents facts: anything the source material doesn't cover is explicitly flagged as an open item for the architect to confirm, rather than guessed at.
6. If an official architecture-repository export isn't available, it produces a clean, presentation-quality conceptual architecture diagram (current state and/or future state) as a stand-in, clearly labeled as a placeholder to be swapped out before formal review.
7. It assembles everything into a single, properly formatted Word document matching the company's SAD template structure, with the architecture diagram also delivered as an editable companion file.
8. The Solution Architect reviews the finished document for accuracy, confirms any flagged open items with the project team, and submits it through the normal SARB/TechPASS review process.

## Technical Flow

This is a model-driven ("assisted") skill: no single script produces the final document. Per `SKILL.md`, the model carries out the following steps itself, calling bundled helper scripts for extraction, rendering, and assembly along the way:

1. **Extraction** — the model runs `lib\extract_docs.py --input "<file1>" "<file2>" ... --out-dir "<work_dir>\extract"` (or `--input-dir "<folder>"`) to pull plain text out of every `.docx`/`.pptx`/`.xlsx`/`.pdf`/`.txt` input into a working `extract/` folder, then reads each resulting `.txt` file itself.
2. **Requirement identification and drafting** — the model reads the extracted BRD content to identify business requirements, then authors section 1 (Overview, Constraints, Risks/Contingencies) prose/bullets/tables itself, applying the skill's documented Writing Standard (active voice, one clean statement per fact, no diary-style update logs, consolidated "Not applicable" answers, consistent terminology, no fabrication).
3. **Architecture diagram (section 2.1)** — if Ardoq exports exist, the model notes they should be pasted in directly. Otherwise, it invokes the sibling `swimlane-flow-diagram` skill's toolkit: it copies `build_swimlane_pptx.py` into a working `diagrams\` folder, customizes lanes (systems/actors/hosting tiers) and boxes/connectors for the project, builds a Current State and/or Future State slide, runs the generator, then performs the **mandatory self-check** via that toolkit's `render_and_verify.py` (PowerPoint COM automation reroutes connectors and exports PNGs for visual inspection) before embedding the PNG(s) into the Word document and keeping the editable `.pptx` as a companion deliverable. If PowerPoint/pywin32 isn't available, the model falls back to the simpler `lib\arch_diagram.py` matplotlib helpers (`new_figure`, `box`, `arrow`, `save`) and notes the limitation.
4. **Sequence diagrams (section 2.2)** — for each identified business requirement, the model defines participants and labeled messages, then renders each to a PNG using `lib\seqrender.py`'s `render_sequence(...)`, also emitting an equivalent websequencediagrams.com script (`.wsd.txt`) as an editable fallback.
5. **Remaining sections (2.3–2.4)** — the model drafts Systems/System Connections/Integration Approaches/Middleware/Data Objects/Data Strategy/Hosting/Existence/Deployment details and the Interface Inventory table from the extracted content, again rewritten per the Writing Standard, applying the standard SA note that integrations should be Delta/Asynchronous/Real-Time unless inputs state otherwise.
6. **Assembly** — the model writes a per-project Python script that imports `lib\sad_builder.py`'s `SADDocument` class (a `python-docx` wrapper) and calls `cover_page(...)`, then `heading(...)`/`para(...)`/`bullets(...)`/`table(...)`/`add_sequence_diagram(...)`/`image(...)` for each template section in order, ending with `save("<output path>.docx")`, then runs the script.
7. **Verification** — the model re-opens the output `.docx` with `python-docx` to confirm heading count/content presence, and spot-checks sections against the Writing Standard before delivering.
8. The model reports back the `.docx` path, the companion architecture `.pptx` path, the list of requirements covered by sequence diagrams, and any sections left as open items.

## Inputs

- Project input documents: BRD, kickoff deck, SARB deck, questionnaire, prior architecture notes, or any other project material (`.docx`/`.pptx`/`.xlsx`/`.pdf`/`.txt`).
- Desired output `.docx` path/filename (defaults to `<Project Name> - Solution Architecture Document.docx` alongside the inputs if not specified).
- Optional: Ardoq Level 2 architecture diagram exports, if available, to be pasted in directly in place of the generated conceptual diagram.

## Outputs

- **`<Project Name> - Solution Architecture Document.docx`** — the complete SAD following the SBD template structure, with rewritten executive-grade prose for every section.
- **Architecture diagram `.pptx`** — editable swimlane-style conceptual Current State/Future State diagram (or matplotlib fallback), embedded as PNG(s) in the document and also delivered as a standalone editable file.
- **Sequence diagram PNGs** — one rendered image per identified business requirement, embedded in section 2.2.
- **`.wsd.txt` files** — one editable websequencediagrams.com script per sequence diagram, as an editable fallback/companion to each PNG.
- **`extract/` working folder** — the raw extracted text pulled from each input document, used to author the SAD content.
