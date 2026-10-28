let skills = [];
let selectedSkill = null;

const skillListEl = document.getElementById("skill-list");
const emptyStateEl = document.getElementById("empty-state");
const detailEl = document.getElementById("skill-detail");
const nameEl = document.getElementById("skill-name");
const modeEl = document.getElementById("skill-mode");
const descEl = document.getElementById("skill-description");
const noteEl = document.getElementById("skill-note");
const resultsEl = document.getElementById("skill-results");
const formEl = document.getElementById("params-form");
const runBtn = document.getElementById("run-btn");
const copyBtn = document.getElementById("copy-btn");
const spinnerEl = document.getElementById("spinner");
const outputSection = document.getElementById("output-section");
const outputLog = document.getElementById("output-log");
const resultFilesEl = document.getElementById("result-files");

async function init() {
  const res = await fetch("/api/skills");
  const manifest = await res.json();
  skills = manifest.skills;
  renderList();
}

function renderList() {
  skillListEl.innerHTML = "";
  skills.forEach((s, idx) => {
    const card = document.createElement("div");
    card.className = "skill-card";
    card.innerHTML = `<div class="name">${escapeHtml(s.name)}</div><div class="tag">${s.mode === "assisted" ? "Assisted" : "Script"}</div>`;
    card.addEventListener("click", () => selectSkill(idx, card));
    skillListEl.appendChild(card);
  });
}

function selectSkill(idx, cardEl) {
  document.querySelectorAll(".skill-card").forEach((c) => c.classList.remove("selected"));
  cardEl.classList.add("selected");
  selectedSkill = skills[idx];
  renderDetail(selectedSkill);
}

function renderDetail(s) {
  emptyStateEl.style.display = "none";
  detailEl.style.display = "block";

  nameEl.textContent = s.name;
  modeEl.textContent = s.mode === "assisted" ? "Assisted" : `Script (${s.runner})`;
  modeEl.className = "badge" + (s.mode === "assisted" ? " assisted" : "");
  descEl.textContent = s.description;

  if (s.note) {
    noteEl.style.display = "block";
    noteEl.textContent = "Note: " + s.note;
  } else {
    noteEl.style.display = "none";
  }

  resultsEl.textContent = s.outputs || "(not documented)";

  formEl.innerHTML = "";
  (s.params || []).forEach((p) => {
    const row = document.createElement("div");
    row.className = "param-row";
    row.dataset.flag = p.flag;

    const label = document.createElement("label");
    label.innerHTML = `${escapeHtml(p.flag)} <span class="hint">(${p.type}${p.required ? ", required" : ", optional"})</span>${p.required ? '<span class="required-mark">*</span>' : ""} — ${escapeHtml(p.prompt)}`;
    row.appendChild(label);

    const group = document.createElement("div");
    group.className = "param-input-group";

    if (p.type === "switch") {
      const chk = document.createElement("input");
      chk.type = "checkbox";
      chk.dataset.input = "true";
      group.appendChild(chk);
    } else {
      const input = document.createElement("input");
      input.type = "text";
      input.dataset.input = "true";
      group.appendChild(input);

      if (p.type === "path" || p.type === "path-list") {
        const browseBtn = document.createElement("button");
        browseBtn.type = "button";
        browseBtn.className = "browse-btn";
        browseBtn.textContent = "Browse…";
        browseBtn.addEventListener("click", () => browseFor(p, input));
        group.appendChild(browseBtn);
      }
    }

    row.appendChild(group);
    formEl.appendChild(row);
  });

  runBtn.style.display = s.mode === "assisted" ? "none" : "inline-block";
  copyBtn.style.display = s.mode === "assisted" ? "inline-block" : "none";
  outputSection.style.display = "none";
  outputLog.textContent = "";
  resultFilesEl.innerHTML = "";
}

function guessDialogKind(p) {
  const prompt = (p.prompt || "").toLowerCase();
  const flag = (p.flag || "").toLowerCase();
  if (flag.includes("directory") || prompt.includes("directory")) return "folder";
  if (prompt.startsWith("output") || prompt.includes("output") && prompt.includes("path")) return "save";
  return "open";
}

async function browseFor(p, input) {
  const kind = guessDialogKind(p);
  const multiple = p.type === "path-list";
  const res = await fetch("/api/browse", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ kind, multiple }),
  });
  const data = await res.json();
  if (data.error) {
    alert("Could not open dialog: " + data.error);
    return;
  }
  if (multiple && data.paths) {
    input.value = data.paths.join(";");
  } else if (data.path) {
    input.value = data.path;
  }
}

function collectParams() {
  const params = {};
  document.querySelectorAll("#params-form .param-row").forEach((row) => {
    const flag = row.dataset.flag;
    const el = row.querySelector("[data-input]");
    if (el.type === "checkbox") {
      params[flag] = el.checked;
    } else {
      params[flag] = el.value;
    }
  });
  return params;
}

function validateRequired() {
  for (const p of selectedSkill.params || []) {
    if (p.required && p.type !== "switch") {
      const row = document.querySelector(`#params-form .param-row[data-flag="${cssEscape(p.flag)}"]`);
      const el = row.querySelector("[data-input]");
      if (!el.value || !el.value.trim()) {
        alert("Please provide a value for: " + p.prompt);
        el.focus();
        return false;
      }
    }
  }
  return true;
}

function cssEscape(s) {
  return s.replace(/["\\]/g, "\\$&");
}

function escapeHtml(s) {
  const div = document.createElement("div");
  div.textContent = s == null ? "" : String(s);
  return div.innerHTML;
}

runBtn.addEventListener("click", async () => {
  if (!selectedSkill) return;
  if (!validateRequired()) return;

  const params = collectParams();
  runBtn.disabled = true;
  spinnerEl.style.display = "inline";
  outputSection.style.display = "block";
  outputLog.textContent = "Running " + selectedSkill.name + " ...\n";
  resultFilesEl.innerHTML = "";

  try {
    const res = await fetch("/api/run", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ skillId: selectedSkill.id, params }),
    });
    const data = await res.json();
    outputLog.textContent += data.output || "(no output)";
    outputLog.textContent += "\n\n" + (data.success ? "✔ Completed successfully." : "✘ Finished with errors.");

    (data.resultFiles || []).forEach((rel) => {
      const link = document.createElement("a");
      link.textContent = "Open: " + rel;
      link.addEventListener("click", async () => {
        await fetch(`/api/open-file?skill=${encodeURIComponent(selectedSkill.folder)}&path=${encodeURIComponent(rel)}`);
      });
      resultFilesEl.appendChild(link);
    });
  } catch (e) {
    outputLog.textContent += "\nRequest failed: " + e;
  } finally {
    runBtn.disabled = false;
    spinnerEl.style.display = "none";
  }
});

copyBtn.addEventListener("click", () => {
  if (!selectedSkill) return;
  const params = collectParams();
  const lines = [`Use the ${selectedSkill.id} skill.`];
  for (const p of selectedSkill.params || []) {
    const v = params[p.flag];
    if (v !== undefined && v !== "" && v !== false) {
      lines.push(`- ${p.flag}: ${v}`);
    }
  }
  const text = lines.join("\n");
  navigator.clipboard.writeText(text).then(() => {
    outputSection.style.display = "block";
    outputLog.textContent = "Copied to clipboard — paste into a Copilot CLI chat:\n\n" + text;
  });
});

init();
