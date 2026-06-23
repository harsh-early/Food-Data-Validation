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
  let referOption = null;
  let outputHandle = null;
  let outputFileName = "";
  let excelWriteTimer = null;
  let loadedJsonName = "";
  let usingUploadedJson = false;
  let pendingExcelSync = false;

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

  function formatReferOptionExport(food, review) {
    if (review.choice !== "both_wrong" || !review.refer_option) return "";
    const macros =
      review.refer_option === "option1"
        ? food.input_macros_raw || formatMacros(food.input_macros)
        : food.final_macros_raw || formatMacros(food.final_macros);
    return `${review.refer_option} | ${macros}`;
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
      const allowed = await ensureOutputPermission(outputHandle);
      if (!allowed) outputHandle = null;
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
      "refer_option",
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
      "option|cal|carbs|fat|prot|fib",
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
        formatReferOptionExport(f, review),
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

  async function writeExcelToHandle(options = {}) {
    const { quiet = false } = options;
    if (!outputHandle || typeof XLSX === "undefined") return false;

    const rowCount = Object.keys(reviews).length;
    const maxAttempts = 4;

    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      try {
        const wb = buildWorkbook();
        const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
        const writable = await outputHandle.createWritable();
        if (typeof writable.truncate === "function") {
          await writable.truncate(0);
        }
        await writable.write(buf);
        await writable.close();
        pendingExcelSync = false;
        if (!quiet) {
          showToast(`Sheet updated · ${rowCount} reviewed row(s)`);
        }
        updateFileStatus(rowCount);
        return true;
      } catch (err) {
        console.error(`Excel write attempt ${attempt} failed:`, err);
        if (attempt < maxAttempts) {
          await new Promise((r) => setTimeout(r, 350 * attempt));
        }
      }
    }

    pendingExcelSync = true;
    if (!quiet) {
      showToast("Close Excel, then click Link Output to sync");
    }
    updateFileStatus(rowCount, true);
    return false;
  }

  function scheduleExcelWrite() {
    clearTimeout(excelWriteTimer);
    excelWriteTimer = setTimeout(async () => {
      if (!outputHandle) return;
      await writeExcelToHandle({ quiet: true });
    }, 300);
  }

  function downloadExcel() {
    if (typeof XLSX === "undefined") {
      showToast("Excel library not loaded — refresh the page");
      return;
    }

    const verifiedCount = Object.keys(reviews).length;

    try {
      const wb = buildWorkbook();
      const date = new Date().toISOString().slice(0, 10);
      const filename = `food_macro_reviews_${date}.xlsx`;
      const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([buf], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      showToast(
        verifiedCount
          ? `Downloaded ${verifiedCount} reviewed item(s)`
          : "Downloaded — no reviews yet (headers only)"
      );
    } catch (err) {
      console.error("Excel download failed:", err);
      showToast("Download failed — try Chrome or Edge");
    }
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

  async function ensureOutputPermission(handle) {
    if (!handle) return false;
    try {
      const perm = await handle.queryPermission({ mode: "readwrite" });
      if (perm === "granted") return true;
      const req = await handle.requestPermission({ mode: "readwrite" });
      return req === "granted";
    } catch (_) {
      return false;
    }
  }

  async function connectOutputFile(handle) {
    await saveFileHandle(handle, handle.name);
    const ok = await writeExcelToHandle();
    updateFileStatus();
    $("#link-banner").classList.add("hidden");
    const n = Object.keys(reviews).length;
    showToast(ok ? `Linked · ${n} row(s) synced` : "Linked but could not write");
    return ok;
  }

  async function syncLinkedOutput() {
    if (!outputHandle) return false;
    const allowed = await ensureOutputPermission(outputHandle);
    if (!allowed) {
      outputHandle = null;
      return false;
    }
    return writeExcelToHandle();
  }

  const EXCEL_PICKER_TYPES = [
    {
      description: "Excel",
      accept: {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
      },
    },
  ];

  function hasFileSystemAccess() {
    return (
      typeof window.showSaveFilePicker === "function" ||
      typeof window.showOpenFilePicker === "function"
    );
  }

  function showLinkOutputModal() {
    return new Promise((resolve) => {
      const modal = $("#link-output-modal");
      const onChoice = (choice) => {
        modal.classList.add("hidden");
        $("#link-modal-new").removeEventListener("click", onNew);
        $("#link-modal-existing").removeEventListener("click", onExisting);
        $("#link-modal-cancel").removeEventListener("click", onCancel);
        modal.removeEventListener("click", onBackdrop);
        document.removeEventListener("keydown", onKey);
        resolve(choice);
      };
      const onNew = () => onChoice("new");
      const onExisting = () => onChoice("existing");
      const onCancel = () => onChoice(null);
      const onBackdrop = (e) => {
        if (e.target === modal) onChoice(null);
      };
      const onKey = (e) => {
        if (e.key === "Escape") onChoice(null);
      };

      $("#link-modal-new").addEventListener("click", onNew);
      $("#link-modal-existing").addEventListener("click", onExisting);
      $("#link-modal-cancel").addEventListener("click", onCancel);
      modal.addEventListener("click", onBackdrop);
      document.addEventListener("keydown", onKey);
      modal.classList.remove("hidden");
    });
  }

  async function pickNewOutputFile() {
    if (!window.showSaveFilePicker) {
      showToast("Use Chrome/Edge for linked file auto-save");
      downloadExcel();
      return;
    }

    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: outputFileName || "food_macro_reviews.xlsx",
        types: EXCEL_PICKER_TYPES,
      });
      await connectOutputFile(handle);
    } catch (err) {
      if (err.name !== "AbortError") showToast("Could not create file");
    }
  }

  async function pickExistingOutputFile() {
    if (!window.showOpenFilePicker) {
      showToast("Use Chrome/Edge for linked file auto-save");
      downloadExcel();
      return;
    }

    try {
      const [handle] = await window.showOpenFilePicker({
        types: EXCEL_PICKER_TYPES,
        multiple: false,
      });
      await connectOutputFile(handle);
    } catch (err) {
      if (err.name !== "AbortError") showToast("Could not link file");
    }
  }

  async function pickOutputFile(preferExisting = false) {
    if (!hasFileSystemAccess()) {
      showToast("Use Chrome/Edge for linked file auto-save");
      downloadExcel();
      return;
    }

    if (preferExisting) {
      await pickExistingOutputFile();
      return;
    }

    const choice = await showLinkOutputModal();
    if (choice === "new") await pickNewOutputFile();
    else if (choice === "existing") await pickExistingOutputFile();
  }

  async function linkOutputFile() {
    if (!hasFileSystemAccess()) {
      showToast("Use Chrome/Edge for linked file auto-save");
      downloadExcel();
      return;
    }

    if (outputHandle) {
      const ok = await syncLinkedOutput();
      if (ok) {
        const n = Object.keys(reviews).length;
        updateFileStatus();
        $("#link-banner").classList.add("hidden");
        showToast(`Updated linked file · ${n} row(s)`);
      } else {
        showToast("Lost access — select your output file again");
        await pickOutputFile(true);
      }
      return;
    }

    await pickOutputFile();
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

  function updateFileStatus(rowCount, syncFailed) {
    const el = $("#file-status");
    const btn = $("#btn-link");
    const count = rowCount ?? Object.keys(reviews).length;
    const linked = !!(outputHandle && outputFileName);

    if (linked) {
      let text = `Linked: ${outputFileName} · ${count} row(s)`;
      if (pendingExcelSync || syncFailed) {
        text += " · sync pending (close Excel)";
      }
      el.textContent = text;
      el.classList.add("linked");
      if (pendingExcelSync || syncFailed) {
        el.classList.add("sync-pending");
      } else {
        el.classList.remove("sync-pending");
      }
      btn.textContent = "Linked ✓";
      btn.classList.add("linked");
      $("#link-banner").classList.add("hidden");
    } else {
      el.textContent = "No output file linked";
      el.classList.remove("linked", "sync-pending");
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
      <div class="food-header-card">
        <h2 class="food-title">${escapeHtml(food.original_food_name)}</h2>
        <div class="serving-line">
          <span class="serving-card-title">serving :</span>
          <span class="serving-card-value">${escapeHtml(serving.label)} | ${escapeHtml(volumeText || "—")}</span>
        </div>
      </div>
    `;

    if (review) {
      const choiceLabel =
        review.choice === "option1"
          ? "Option 1 chosen"
          : review.choice === "option2"
            ? "Option 2 chosen"
            : `Both Wrong saved${review.refer_option ? ` · refer ${review.refer_option === "option1" ? "Option 1" : "Option 2"}` : ""}`;
      const badge = document.createElement("div");
      badge.className = `review-badge ${review.choice}`;
      badge.innerHTML = `✓ ${choiceLabel} <button type="button" aria-label="Undo review">✕</button>`;
      badge.querySelector("button").addEventListener("click", () => removeReview(food.id));
      el.appendChild(badge);
    }
  }

  function clearCorrectionInputs() {
    MACRO_KEYS.forEach((k) => {
      const input = document.querySelector(`[data-correct="${k}"]`);
      if (input) input.value = "";
    });
    $("#comment-input").value = "";
    referOption = null;
  }

  function readCorrectionsFromForm() {
    const corrections = {};
    MACRO_KEYS.forEach((k) => {
      const input = document.querySelector(`[data-correct="${k}"]`);
      const val = input?.value?.trim();
      if (val) corrections[k] = val;
    });
    return corrections;
  }

  function fillCorrectionInputs(corrections) {
    MACRO_KEYS.forEach((k) => {
      const input = document.querySelector(`[data-correct="${k}"]`);
      if (input) input.value = corrections?.[k] ?? "";
    });
  }

  function syncBothWrongUI(food) {
    const review = reviews[food.id];
    const panel = $("#both-wrong-panel");
    const btn = $("#btn-both-wrong");

    if (review?.choice === "both_wrong") {
      bothWrongOpen = true;
      referOption = review.refer_option || null;
      fillCorrectionInputs(review.corrections);
      $("#comment-input").value = review.comment || "";
    } else if (!bothWrongOpen) {
      clearCorrectionInputs();
    }

    panel.classList.toggle("hidden", !bothWrongOpen);
    panel.classList.toggle("active", review?.choice === "both_wrong");
    btn.classList.toggle("active", bothWrongOpen || review?.choice === "both_wrong");

    $("#btn-refer-option1")?.classList.toggle("active", referOption === "option1");
    $("#btn-refer-option2")?.classList.toggle("active", referOption === "option2");
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
    clearCorrectionInputs();
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
    const corrections = readCorrectionsFromForm();
    const hasCorrections = Object.keys(corrections).length > 0;

    if (!hasCorrections && !comment) {
      showToast("Add a correction or comment");
      return;
    }

    const review = {
      choice: "both_wrong",
      corrections,
      comment,
    };
    if (referOption) review.refer_option = referOption;

    await saveReview(review);
    showToast("✓ Both Wrong saved");
    setTimeout(() => {
      bothWrongOpen = false;
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
    clearCorrectionInputs();
    $("#loading").classList.add("hidden");
    $("#main").classList.remove("hidden");
    if (outputHandle) scheduleExcelWrite();
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
      render();
    });

    $("#btn-save-both-wrong").addEventListener("click", handleBothWrongSave);

    $("#btn-refer-option1").addEventListener("click", () => {
      referOption = "option1";
      bothWrongOpen = true;
      render();
    });

    $("#btn-refer-option2").addEventListener("click", () => {
      referOption = "option2";
      bothWrongOpen = true;
      render();
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

    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible" && pendingExcelSync && outputHandle) {
        scheduleExcelWrite();
      }
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
