(function () {
  "use strict";

  const DB_NAME = "earlyfit_food_reviewer";
  const DB_VERSION = 2;
  const UPLOAD_KEY = "current";
  const MACRO_KEYS = ["cal", "carbs", "fat", "prot", "fib"];
  const MACRO_LABELS = {
    cal: "Cal (kcal)",
    carbs: "Carbs (g)",
    fat: "Fat (g)",
    prot: "Protein (g)",
    fib: "Fiber (g)",
  };
  const CORRECTION_LABELS = {
    cal: "cal",
    carbs: "carbs",
    fat: "fat",
    prot: "protein",
    fib: "fiber",
  };
  const CORRECTION_UNITS = {
    cal: "kcal",
    carbs: "g",
    fat: "g",
    prot: "g",
    fib: "g",
  };

  let foods = [];
  let meta = {};
  let datasetId = "";
  let reviews = {};
  let cursor = 0;
  let filter = "all";
  let bothWrongOpen = false;
  let activeMacro = null;
  let draftCorrections = {};
  let outputHandle = null;
  let outputFileName = "";
  let excelWriteTimer = null;
  let loadedJsonName = "";
  let usingUploadedJson = false;

  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => document.querySelectorAll(sel);

  // ── Serving ──────────────────────────────────────────────────────────────

  function parseBaseServing(servingTypes, baseVolume) {
    const line =
      (servingTypes || []).find((s) => /\[BASE\]/i.test(s)) ||
      servingTypes?.[0] ||
      "";
    const cleaned = line.replace(/\[BASE\]\s*/i, "").trim();
    const bracket = cleaned.indexOf("[[");
    const main = bracket >= 0 ? cleaned.slice(0, bracket).trim() : cleaned;
    const parts = main.split(/[—–-]/).map((s) => s.trim());
    return {
      label: parts[0] || "—",
      volume: parts[1] || baseVolume || "",
    };
  }

  function formatServingVolume(volume) {
    if (!volume) return "";
    const v = String(volume).trim();
    const num = parseFloat(v.replace(/[^\d.]/g, ""));
    if (isNaN(num)) return v;
    const rounded = Number.isInteger(num) ? String(Math.round(num)) : String(num);
    return `${rounded}gm`;
  }

  // ── Macros helpers ─────────────────────────────────────────────────────

  function formatMacros(m) {
    if (!m) return "";
    return MACRO_KEYS.map((k) => {
      const v = m[k];
      return v === null || v === undefined || v === "" ? "0.0" : Number(v).toFixed(1);
    }).join(" | ");
  }

  function diffPct(a, b) {
    if (!a || !b || !a.cal || !b.cal || a.cal <= 0) return null;
    return ((b.cal - a.cal) / a.cal * 100).toFixed(1);
  }

  function diffColor(pct) {
    if (pct === null) return { cls: "gray", label: "—" };
    const abs = Math.abs(parseFloat(pct));
    const sign = parseFloat(pct) > 0 ? "+" : "";
    const label = `${sign}${pct}%`;
    if (abs <= 3) return { cls: "green", label };
    if (abs <= 10) return { cls: "yellow", label };
    if (abs <= 30) return { cls: "orange", label };
    return { cls: "red", label };
  }

  function formatBothWrongCorrection(corrections) {
    if (!corrections) return "";
    return MACRO_KEYS.filter((k) => corrections[k] != null && corrections[k] !== "")
      .map((k) => `${CORRECTION_LABELS[k]}=${corrections[k]} ${CORRECTION_UNITS[k]}`)
      .join(", ");
  }

  function computeDatasetId(m) {
    const src = m?.source || "unknown";
    const count = m?.count || 0;
    const gen = m?.generated_at || "";
    return `${src}|${count}|${gen}`;
  }

  // ── IndexedDB ──────────────────────────────────────────────────────────

  function openDB() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = (event) => {
        const db = req.result;
        if (!db.objectStoreNames.contains("datasets")) {
          db.createObjectStore("datasets");
        }
        if (!db.objectStoreNames.contains("fileHandles")) {
          db.createObjectStore("fileHandles");
        }
        if (!db.objectStoreNames.contains("uploadedFoods")) {
          db.createObjectStore("uploadedFoods");
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function idbGet(store, key) {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(store, "readonly");
      const req = tx.objectStore(store).get(key);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function idbSet(store, key, value) {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(store, "readwrite");
      tx.objectStore(store).put(value, key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  async function idbDelete(store, key) {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(store, "readwrite");
      tx.objectStore(store).delete(key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  async function saveUploadedFoods(data, filename) {
    await idbSet("uploadedFoods", UPLOAD_KEY, {
      data,
      filename,
      savedAt: new Date().toISOString(),
    });
    sessionStorage.setItem("food_upload_active", "1");
  }

  async function clearUploadedFoods() {
    await idbDelete("uploadedFoods", UPLOAD_KEY);
    sessionStorage.removeItem("food_upload_active");
    loadedJsonName = "";
    usingUploadedJson = false;
  }

  async function getUploadedFoods() {
    return idbGet("uploadedFoods", UPLOAD_KEY);
  }

  function updateJsonStatus() {
    const status = $("#json-status");
    const removeBtn = $("#btn-remove-json");
    if (usingUploadedJson && loadedJsonName) {
      status.textContent = `Loaded: ${loadedJsonName}`;
      status.classList.remove("hidden");
      removeBtn.classList.remove("hidden");
    } else {
      status.textContent = "";
      status.classList.add("hidden");
      removeBtn.classList.add("hidden");
    }
  }

  async function loadDatasetState() {
    const data = await idbGet("datasets", datasetId);
    if (data) {
      reviews = data.reviews || {};
      cursor = data.cursor || 0;
      filter = data.filter || "all";
      outputFileName = data.outputFileName || "";
    } else {
      reviews = {};
      cursor = 0;
      filter = "all";
      outputFileName = "";
    }
    outputHandle = await idbGet("fileHandles", datasetId);
    if (outputHandle) {
      try {
        const perm = await outputHandle.queryPermission({ mode: "readwrite" });
        if (perm !== "granted") {
          const req = await outputHandle.requestPermission({ mode: "readwrite" });
          if (req !== "granted") outputHandle = null;
        }
      } catch (_) {
        outputHandle = null;
      }
    }
  }

  async function saveDatasetState() {
    await idbSet("datasets", datasetId, {
      reviews,
      cursor,
      filter,
      outputFileName,
    });
  }

  async function saveFileHandle(handle, name) {
    outputHandle = handle;
    outputFileName = name;
    await idbSet("fileHandles", datasetId, handle);
    await saveDatasetState();
  }

  // ── Export ─────────────────────────────────────────────────────────────

  function buildVerifiedRows() {
    const headers = [
      "food_name",
      "serving_type",
      "volume",
      "option1_macros",
      "option2_macros",
      "selected_macros",
      "selected_option",
      "both_wrong_correction",
      "comment",
    ];
    const macroSub = [
      "",
      "",
      "",
      "cal|carbs|fat|prot|fib",
      "cal|carbs|fat|prot|fib",
      "cal|carbs|fat|prot|fib",
      "",
      "",
      "",
    ];
    const rows = [headers, macroSub];

    foods.forEach((f) => {
      const review = reviews[f.id];
      if (!review) return;

      const serving = parseBaseServing(f.serving_types, f.base_volume);
      let selectedMacros = "";
      if (review.choice === "option1") {
        selectedMacros = f.input_macros_raw || formatMacros(f.input_macros);
      } else if (review.choice === "option2") {
        selectedMacros = f.final_macros_raw || formatMacros(f.final_macros);
      }

      rows.push([
        f.original_food_name,
        serving.label,
        serving.volume,
        f.input_macros_raw || formatMacros(f.input_macros),
        f.final_macros_raw || formatMacros(f.final_macros),
        selectedMacros,
        review.choice,
        review.choice === "both_wrong"
          ? formatBothWrongCorrection(review.corrections)
          : "",
        review.comment || "",
      ]);
    });

    return rows;
  }

  function buildWorkbook() {
    const rows = buildVerifiedRows();
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Verified");
    return wb;
  }

  async function writeExcelToHandle() {
    if (!outputHandle || typeof XLSX === "undefined") return false;
    try {
      const wb = buildWorkbook();
      const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const writable = await outputHandle.createWritable();
      await writable.write(buf);
      await writable.close();
      return true;
    } catch (err) {
      console.error("Excel write failed:", err);
      showToast("Could not update linked file");
      return false;
    }
  }

  function scheduleExcelWrite() {
    clearTimeout(excelWriteTimer);
    excelWriteTimer = setTimeout(async () => {
      if (outputHandle) {
        const ok = await writeExcelToHandle();
        if (ok) showToast("Saved to linked file");
      }
    }, 300);
  }

  function downloadExcel() {
    if (typeof XLSX === "undefined") {
      showToast("Excel library not loaded");
      return;
    }
    const verified = Object.keys(reviews).length;
    if (!verified) {
      showToast("No verified foods yet");
      return;
    }
    const wb = buildWorkbook();
    const date = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `food_macro_reviews_${date}.xlsx`);
    showToast("Excel downloaded");
  }

  // ── Reviews ────────────────────────────────────────────────────────────

  async function saveReview(review) {
    const food = getFiltered()[cursor];
    if (!food) return;

    reviews[food.id] = { ...review, ts: new Date().toISOString() };
    await saveDatasetState();
    scheduleExcelWrite();
    render();
  }

  async function removeReview(foodId) {
    delete reviews[foodId];
    await saveDatasetState();
    scheduleExcelWrite();
    render();
  }

  async function linkOutputFile() {
    if (!window.showSaveFilePicker) {
      showToast("Use Chrome/Edge for linked file auto-save");
      downloadExcel();
      return;
    }
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: outputFileName || "food_macro_reviews.xlsx",
        types: [
          {
            description: "Excel",
            accept: {
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
            },
          },
        ],
      });
      await saveFileHandle(handle, handle.name);
      await writeExcelToHandle();
      updateFileStatus();
      $("#link-banner").classList.add("hidden");
      showToast("Output file linked");
    } catch (err) {
      if (err.name !== "AbortError") showToast("Could not link file");
    }
  }

  // ── UI helpers ─────────────────────────────────────────────────────────

  function escapeHtml(s) {
    const d = document.createElement("div");
    d.textContent = s;
    return d.innerHTML;
  }

  function showToast(msg) {
    const el = $("#toast");
    el.textContent = msg;
    el.classList.remove("hidden");
    setTimeout(() => el.classList.add("hidden"), 1800);
  }

  function getFiltered() {
    return foods.filter((f) => {
      if (filter === "pending") return !reviews[f.id];
      if (filter === "done") return !!reviews[f.id];
      return true;
    });
  }

  function updateFileStatus() {
    const el = $("#file-status");
    const btn = $("#btn-link");
    if (outputHandle && outputFileName) {
      el.textContent = `Linked: ${outputFileName}`;
      el.classList.add("linked");
      btn.textContent = "Linked ✓";
      btn.classList.add("linked");
      $("#link-banner").classList.add("hidden");
    } else {
      el.textContent = "No output file linked";
      el.classList.remove("linked");
      btn.textContent = "Link Output";
      btn.classList.remove("linked");
      if (Object.keys(reviews).length > 0) {
        $("#link-banner").classList.remove("hidden");
      }
    }
  }

  function renderMacroCard(option, label, macros, chosen, onChoose) {
    const card = document.createElement("button");
    card.type = "button";
    card.className = `macro-card ${option}${chosen ? " chosen" : ""}`;

    const header = document.createElement("div");
    header.className = "macro-card-header";
    header.innerHTML = `${label}${chosen ? '<span class="selected-mark">✓ selected</span>' : ""}`;
    card.appendChild(header);

    const body = document.createElement("div");
    body.className = "macro-card-body";

    if (macros) {
      MACRO_KEYS.forEach((k) => {
        const row = document.createElement("div");
        row.className = "macro-row";
        const val = macros[k];
        const valCls = val === null ? "empty" : val === 0 && k !== "cal" ? "zero-warn" : "";
        row.innerHTML = `<span class="macro-label">${MACRO_LABELS[k]}</span><span class="macro-value ${valCls}">${val === null ? "—" : val}</span>`;
        body.appendChild(row);
      });
    } else {
      const empty = document.createElement("p");
      empty.className = "macro-empty";
      empty.textContent = "No data available";
      body.appendChild(empty);
    }

    card.appendChild(body);
    card.addEventListener("click", onChoose);
    return card;
  }

  function renderFoodInfo(food) {
    const el = $("#food-info");
    el.innerHTML = "";

    const review = reviews[food.id];
    const serving = parseBaseServing(food.serving_types, food.base_volume);
    const volumeText = formatServingVolume(serving.volume);

    el.innerHTML = `
      <h2 class="food-title">${escapeHtml(food.original_food_name)}</h2>
      <div class="serving-card">
        <span class="serving-card-title">serving :</span>
        <span class="serving-card-value">${escapeHtml(serving.label)} | ${escapeHtml(volumeText || "—")}</span>
      </div>
    `;

    if (review) {
      const choiceLabel =
        review.choice === "option1"
          ? "Option 1 chosen"
          : review.choice === "option2"
            ? "Option 2 chosen"
            : "Both Wrong saved";
      const badge = document.createElement("div");
      badge.className = `review-badge ${review.choice}`;
      badge.innerHTML = `✓ ${choiceLabel} <button type="button" aria-label="Undo review">✕</button>`;
      badge.querySelector("button").addEventListener("click", () => removeReview(food.id));
      el.appendChild(badge);
    }
  }

  function syncBothWrongUI(food) {
    const review = reviews[food.id];
    const panel = $("#both-wrong-panel");
    const btn = $("#btn-both-wrong");

    if (review?.choice === "both_wrong") {
      bothWrongOpen = true;
      draftCorrections = { ...(review.corrections || {}) };
      $("#comment-input").value = review.comment || "";
    } else if (!bothWrongOpen) {
      draftCorrections = {};
      activeMacro = null;
      $("#comment-input").value = "";
    }

    panel.classList.toggle("hidden", !bothWrongOpen);
    panel.classList.toggle("active", review?.choice === "both_wrong");
    btn.classList.toggle("active", bothWrongOpen || review?.choice === "both_wrong");

    $$(".macro-chip").forEach((chip) => {
      const key = chip.dataset.macro;
      chip.classList.toggle("active", activeMacro === key);
      chip.classList.toggle("has-value", draftCorrections[key] != null && draftCorrections[key] !== "");
    });

    const input = $("#correction-input");
    const label = $("#active-macro-label");
    if (activeMacro) {
      label.textContent = `Correct ${MACRO_LABELS[activeMacro]}`;
      input.disabled = false;
      input.value = draftCorrections[activeMacro] ?? "";
    } else {
      label.textContent = "Select a macro above";
      input.disabled = true;
      input.value = "";
    }
  }

  function render() {
    const filtered = getFiltered();
    const total = foods.length;
    const done = Object.keys(reviews).length;
    const pct = total ? Math.round((done / total) * 100) : 0;

    $("#progress-text").textContent = `${done}/${total} reviewed · ${pct}% done`;
    $("#progress-fill").style.width = `${pct}%`;
    $("#footer-text").textContent = `Auto-saved · ${done} verified`;
    updateFileStatus();
    updateJsonStatus();

    $$(".filter-btn").forEach((b) => {
      b.classList.toggle("active", b.dataset.filter === filter);
    });

    const empty = filtered.length === 0;
    $("#empty-state").classList.toggle("hidden", !empty);
    $("#review-panel").classList.toggle("hidden", empty);

    if (empty) return;

    if (cursor >= filtered.length) cursor = filtered.length - 1;
    if (cursor < 0) cursor = 0;

    const food = filtered[cursor];
    $("#cursor-text").textContent = `${cursor + 1}/${filtered.length}`;

    renderFoodInfo(food);

    const cards = $("#macro-cards");
    cards.innerHTML = "";
    const chosen = reviews[food.id]?.choice ?? null;
    const optionChosen = chosen === "option1" || chosen === "option2";

    cards.appendChild(
      renderMacroCard("option1", "Option 1", food.input_macros, chosen === "option1", () =>
        handleChooseOption("option1")
      )
    );
    cards.appendChild(
      renderMacroCard("option2", "Option 2", food.final_macros, chosen === "option2", () =>
        handleChooseOption("option2")
      )
    );

    if (optionChosen) bothWrongOpen = false;
    syncBothWrongUI(food);

    $("#btn-prev").disabled = cursor === 0;
    $("#btn-next").disabled = cursor >= filtered.length - 1;
  }

  async function moveCursor(dir) {
    const filtered = getFiltered();
    cursor = Math.max(0, Math.min(filtered.length - 1, cursor + dir));
    bothWrongOpen = false;
    activeMacro = null;
    draftCorrections = {};
    await saveDatasetState();
    render();
  }

  async function handleChooseOption(choice) {
    bothWrongOpen = false;
    await saveReview({ choice });
    showToast(choice === "option1" ? "✓ Option 1 saved" : "✓ Option 2 saved");
    setTimeout(() => moveCursor(1), 300);
  }

  async function handleBothWrongSave() {
    const food = getFiltered()[cursor];
    if (!food) return;

    const comment = $("#comment-input").value.trim();
    const hasCorrections = MACRO_KEYS.some(
      (k) => draftCorrections[k] != null && draftCorrections[k] !== ""
    );

    if (!hasCorrections && !comment) {
      showToast("Add a correction or comment");
      return;
    }

    const corrections = {};
    MACRO_KEYS.forEach((k) => {
      if (draftCorrections[k] != null && draftCorrections[k] !== "") {
        corrections[k] = String(draftCorrections[k]);
      }
    });

    await saveReview({
      choice: "both_wrong",
      corrections,
      comment,
    });
    showToast("✓ Both Wrong saved");
    setTimeout(() => {
      bothWrongOpen = false;
      activeMacro = null;
      moveCursor(1);
    }, 300);
  }

  async function loadFoodData(data, filename, options = {}) {
    if (!data || !Array.isArray(data.foods)) {
      throw new Error("Invalid JSON: expected { foods: [...] }");
    }
    foods = data.foods;
    meta = data.meta || { source: filename || "upload", count: foods.length };
    if (!meta.count) meta.count = foods.length;
    datasetId = computeDatasetId(meta);
    loadedJsonName = filename || meta.source || "uploaded.json";
    usingUploadedJson = !!options.persisted || !!options.fromUpload;

    if (options.fromUpload) {
      await saveUploadedFoods(data, loadedJsonName);
      usingUploadedJson = true;
    }

    await loadDatasetState();
    bothWrongOpen = false;
    activeMacro = null;
    draftCorrections = {};
    $("#loading").classList.add("hidden");
    $("#main").classList.remove("hidden");
    render();
  }

  async function loadDefaultFoods() {
    const res = await fetch("./foods.json");
    if (!res.ok) throw new Error(`Failed to load foods.json (${res.status})`);
    const data = await res.json();
    loadedJsonName = "";
    usingUploadedJson = false;
    await loadFoodData(data, "foods.json");
  }

  async function removeUploadedJson() {
    await clearUploadedFoods();
    showToast("Uploaded JSON removed");
    $("#loading").classList.remove("hidden");
    $("#main").classList.add("hidden");
    try {
      await loadDefaultFoods();
    } catch (err) {
      $("#loading").classList.add("hidden");
      $("#error").textContent = err.message;
      $("#error").classList.remove("hidden");
    }
  }

  function bindEvents() {
    $("#btn-export").addEventListener("click", downloadExcel);
    $("#btn-link").addEventListener("click", linkOutputFile);
    $("#btn-link-banner").addEventListener("click", linkOutputFile);

    $("#input-json").addEventListener("change", async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        const text = await file.text();
        const data = JSON.parse(text);
        await loadFoodData(data, file.name, { fromUpload: true });
        showToast(`Loaded ${foods.length} foods`);
      } catch (err) {
        showToast(err.message || "Failed to load JSON");
      }
      e.target.value = "";
    });

    $("#btn-remove-json").addEventListener("click", async () => {
      if (!confirm("Remove uploaded JSON and load default foods.json?")) return;
      await removeUploadedJson();
    });

    $("#btn-both-wrong").addEventListener("click", () => {
      bothWrongOpen = !bothWrongOpen;
      if (!bothWrongOpen) {
        activeMacro = null;
      }
      render();
    });

    $("#btn-save-both-wrong").addEventListener("click", handleBothWrongSave);

    $$(".macro-chip").forEach((chip) => {
      chip.addEventListener("click", () => {
        if (activeMacro && $("#correction-input").value !== "") {
          draftCorrections[activeMacro] = $("#correction-input").value;
        }
        activeMacro = chip.dataset.macro;
        bothWrongOpen = true;
        render();
        $("#correction-input").focus();
      });
    });

    $("#correction-input").addEventListener("input", (e) => {
      if (activeMacro) {
        draftCorrections[activeMacro] = e.target.value;
        const chip = document.querySelector(`.macro-chip[data-macro="${activeMacro}"]`);
        if (chip) chip.classList.toggle("has-value", e.target.value !== "");
      }
    });

    $("#btn-prev").addEventListener("click", () => moveCursor(-1));
    $("#btn-next").addEventListener("click", () => moveCursor(1));

    $("#btn-view-all").addEventListener("click", () => {
      filter = "all";
      cursor = 0;
      render();
    });

    $$(".filter-btn").forEach((btn) => {
      btn.addEventListener("click", async () => {
        filter = btn.dataset.filter;
        cursor = 0;
        await saveDatasetState();
        render();
      });
    });

    window.addEventListener("keydown", (e) => {
      if (e.target.tagName === "INPUT" || e.target.tagName === "TEXTAREA") return;
      if (e.key === "1") handleChooseOption("option1");
      if (e.key === "2") handleChooseOption("option2");
      if (e.key === "ArrowUp") moveCursor(-1);
      if (e.key === "ArrowDown") moveCursor(1);
    });

    window.addEventListener("beforeunload", () => {
      saveDatasetState();
    });
  }

  async function init() {
    bindEvents();

    try {
      const uploaded = await getUploadedFoods();
      if (uploaded?.data) {
        await loadFoodData(uploaded.data, uploaded.filename, { persisted: true });
        return;
      }
      await loadDefaultFoods();
    } catch (err) {
      $("#loading").classList.add("hidden");
      const errEl = $("#error");
      errEl.textContent = `${err.message} — use Upload JSON to load a file.`;
      errEl.classList.remove("hidden");
    }
  }

  init();
})();
