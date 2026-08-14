/* ===== Bradesco Modal/Card Console — app.js (100% client-side) ===== */
/* global XLSX, JSZip */
(function () {
  "use strict";

  // ============================================================
  //  STATE
  // ============================================================
  const HISTORY_KEY = "bradesco_modal_card_history_v2";
  const state = {
    xlsxFile: null,
    imageFile: null,
    folderFiles: [],
    searchMode: "single", // "single" | "folder"
    running: false,
    aborted: false,
    logs: [],
    matches: [],
    failures: [],
    imageMeta: null,
    runId: null,
    lastRun: null,
    // ---- edição / exportação do xlsx (preenchidos após uma execução) ----
    currentZip: null,          // JSZip do arquivo carregado (mutável)
    currentWorkbook: null,     // workbook do SheetJS (usado p/ resolver merges)
    currentParams: null,       // params da última execução
    currentSheetXmlPath: null, // caminho do xml da aba dentro do zip
    mediaIndex: {},            // celula -> { mediaPath, row, col }
    pendingEdits: {},          // celula -> { cellEdits: {col: {value, isDate}}, imageFile: File|null, imageMediaPath }
  };

  // ============================================================
  //  UTIL
  // ============================================================
  const $ = (id) => document.getElementById(id);
  const on = (el, ev, fn) => el && el.addEventListener(ev, fn);
  const nowTs = () => new Date().toTimeString().slice(0, 8);
  const uuid = () => (crypto.randomUUID?.() || (Date.now() + "-" + Math.random().toString(16).slice(2)));
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const yieldTick = () => new Promise((r) => setTimeout(r, 0));

  function toast(msg, kind = "info") {
    const c = $("toast-container");
    const t = document.createElement("div");
    t.className = "toast" + (kind ? " toast-" + kind : "");
    t.textContent = msg;
    c.appendChild(t);
    setTimeout(() => t.remove(), 3500);
  }

  function escapeHtml(str) {
    if (str == null) return "";
    return String(str)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#039;");
  }

  function formatDate(iso) {
    if (!iso) return "-";
    try { return new Date(iso).toLocaleString("pt-BR"); } catch { return iso; }
  }

  function colLetterToIndex(letters) {
    // 'A' -> 1, 'B' -> 2, ..., 'AA' -> 27
    let n = 0;
    const s = String(letters).toUpperCase();
    for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n;
  }
  function colIndexToLetter(n) {
    let s = "";
    while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  // ============================================================
  //  UI: Tabs
  // ============================================================
  document.querySelectorAll(".tab").forEach((tab) => {
    on(tab, "click", () => {
      document.querySelectorAll(".tab").forEach((t) => t.classList.toggle("active", t === tab));
      const target = tab.dataset.tab;
      $("panel-run").classList.toggle("hidden", target !== "run");
      $("panel-sheet").classList.toggle("hidden", target !== "sheet");
      $("panel-history").classList.toggle("hidden", target !== "history");
      if (target === "history") renderHistory(loadHistory());
      if (target === "sheet") renderSheetTable();
    });
  });

  // ============================================================
  //  UI: Dropzones
  // ============================================================
  function setupDropzone(dzId, inputId, iconId, nameId, clearId, onFile) {
    const dz = $(dzId), input = $(inputId), icon = $(iconId),
          nameEl = $(nameId), clearBtn = $(clearId);

    const setFile = (f) => {
      onFile(f);
      if (f) {
        nameEl.textContent = f.name;
        nameEl.classList.add("filled");
        icon.classList.add("filled");
        clearBtn.classList.remove("hidden");
      } else {
        nameEl.textContent = "Arraste ou clique para selecionar";
        nameEl.classList.remove("filled");
        icon.classList.remove("filled");
        clearBtn.classList.add("hidden");
        input.value = "";
      }
      updateRunButton();
    };

    on(dz, "click", (e) => { if (e.target !== clearBtn) input.click(); });
    on(input, "change", (e) => { if (e.target.files?.[0]) setFile(e.target.files[0]); });
    on(clearBtn, "click", (e) => { e.stopPropagation(); setFile(null); });
    on(dz, "dragover", (e) => { e.preventDefault(); dz.classList.add("dragging"); });
    on(dz, "dragleave", () => dz.classList.remove("dragging"));
    on(dz, "drop", (e) => {
      e.preventDefault(); dz.classList.remove("dragging");
      if (e.dataTransfer.files?.[0]) setFile(e.dataTransfer.files[0]);
    });
  }
  setupDropzone("dz-xlsx", "input-xlsx", "dz-xlsx-icon", "dz-xlsx-name", "dz-xlsx-clear",
    (f) => (state.xlsxFile = f));
  setupDropzone("dz-image", "input-image", "dz-image-icon", "dz-image-name", "dz-image-clear",
    (f) => (state.imageFile = f));

  // ---- Pasta de imagens (webkitdirectory) ----
  function setupFolderDropzone() {
    const dz = $("dz-folder"), input = $("input-folder"), icon = $("dz-folder-icon"),
          nameEl = $("dz-folder-name"), clearBtn = $("dz-folder-clear");

    const setFiles = (files) => {
      const imgs = Array.from(files || []).filter((f) => f.type.startsWith("image/"));
      state.folderFiles = imgs;
      if (imgs.length) {
        const folderName = imgs[0].webkitRelativePath ? imgs[0].webkitRelativePath.split("/")[0] : "pasta selecionada";
        nameEl.textContent = `${folderName} · ${imgs.length} imagem(ns)`;
        nameEl.classList.add("filled");
        icon.classList.add("filled");
        clearBtn.classList.remove("hidden");
      } else {
        nameEl.textContent = "Clique para selecionar uma pasta";
        nameEl.classList.remove("filled");
        icon.classList.remove("filled");
        clearBtn.classList.add("hidden");
        input.value = "";
      }
      updateRunButton();
    };

    on(dz, "click", (e) => { if (e.target !== clearBtn) input.click(); });
    on(input, "change", (e) => setFiles(e.target.files));
    on(clearBtn, "click", (e) => { e.stopPropagation(); setFiles(null); });
  }
  setupFolderDropzone();

  function setSearchMode(mode) {
    state.searchMode = mode;
    $("mode-single").classList.toggle("active", mode === "single");
    $("mode-folder").classList.toggle("active", mode === "folder");
    $("dz-image").classList.toggle("hidden", mode !== "single");
    $("dz-folder").classList.toggle("hidden", mode !== "folder");
    updateRunButton();
  }
  on($("mode-single"), "click", () => setSearchMode("single"));
  on($("mode-folder"), "click", () => setSearchMode("folder"));

  function updateRunButton() {
    const hasSearchTarget = state.searchMode === "folder"
      ? state.folderFiles.length > 0
      : !!state.imageFile;
    $("btn-run").disabled = !(state.xlsxFile && hasSearchTarget) || state.running;
  }

  // ---- Guia "Como usar" ----
  on($("btn-guide-toggle"), "click", () => {
    $("guide-panel").classList.toggle("hidden");
    $("guide-chevron").classList.toggle("rotated");
  });

  // ============================================================
  //  UI: Advanced collapsible + modals
  // ============================================================
  on($("btn-advanced-toggle"), "click", () => {
    $("advanced-panel").classList.toggle("hidden");
    $("advanced-chevron").classList.toggle("rotated");
  });
  document.querySelectorAll("[data-close]").forEach((el) =>
    on(el, "click", (e) => {
      if (e.target === el || el.classList.contains("modal-close")) {
        el.closest(".modal").classList.add("hidden");
      }
    })
  );

  // ============================================================
  //  TERMINAL (log stream)
  // ============================================================
  const termBody = $("terminal-output");
  const termLinesCount = $("term-lines-count");
  const termStatus = $("term-status");
  const termStatusText = $("term-status-text");

  function appendLog(log) {
    const idle = $("idle-cursor"); if (idle) idle.remove();
    const div = document.createElement("div");
    div.className = "log-line log-" + (log.level || "info");
    if (log.ts) {
      const ts = document.createElement("span");
      ts.className = "log-ts"; ts.textContent = `[${log.ts}]`;
      div.appendChild(ts);
    }
    div.appendChild(document.createTextNode(log.message ?? ""));
    termBody.appendChild(div);
    termBody.scrollTop = termBody.scrollHeight;
    state.logs.push(log);
    termLinesCount.textContent = state.logs.length;
  }
  const log = (level, message) => appendLog({ level, message, ts: nowTs() });

  function clearTerminal() {
    state.logs = [];
    state.matches = [];
    state.failures = [];
    state.imageMeta = null;
    state.runId = null;
    state.lastRun = null;
    state.currentZip = null;
    state.currentWorkbook = null;
    state.currentParams = null;
    state.currentSheetXmlPath = null;
    state.mediaIndex = {};
    state.pendingEdits = {};
    termLinesCount.textContent = "0";
    termBody.innerHTML = `
      <div class="log-line log-dim">$ Bradesco Modal/Card Console v2.0 (100% offline)</div>
      <div class="log-line log-dim">$ Aguardando arquivos... selecione a planilha e a imagem, depois clique em EXECUTAR.</div>
      <div class="log-line" id="idle-cursor"><span class="log-success">➜</span> <span class="log-info">console</span><span class="cursor-blink"></span></div>
    `;
    $("summary-card").classList.add("hidden");
    $("matches-preview").classList.add("hidden");
    $("mp-grid").innerHTML = "";
    $("mp-count-badge").textContent = "0";
    updateEditControls();
    renderSheetTable();
  }
  on($("btn-clear"), "click", () => { if (!state.running) clearTerminal(); });

  function setRunning(running, statusText) {
    state.running = running;
    $("btn-run").classList.toggle("hidden", running);
    $("btn-stop").classList.toggle("hidden", !running);
    termStatus.classList.toggle("hidden", !running);
    if (statusText) termStatusText.textContent = statusText;
    updateRunButton();
  }

  // ============================================================
  //  IMAGE PROCESSING (pHash + Sharpness) — pure JS
  // ============================================================

  // Load a Blob/File/URL into an HTMLImageElement
  function loadImage(src) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.crossOrigin = "anonymous";
      img.onload = () => resolve(img);
      img.onerror = (e) => reject(new Error("Falha ao carregar imagem"));
      img.src = typeof src === "string" ? src : URL.createObjectURL(src);
    });
  }

  // Precompute cosine table for 32-point DCT
  const DCT_N = 32;
  const COS_TABLE = (() => {
    const t = new Float64Array(DCT_N * DCT_N);
    for (let k = 0; k < DCT_N; k++)
      for (let n = 0; n < DCT_N; n++)
        t[k * DCT_N + n] = Math.cos((Math.PI / DCT_N) * (n + 0.5) * k);
    return t;
  })();

  function dct1d(vec, out) {
    const N = DCT_N;
    for (let k = 0; k < N; k++) {
      let s = 0;
      const off = k * N;
      for (let n = 0; n < N; n++) s += vec[n] * COS_TABLE[off + n];
      out[k] = s;
    }
  }

  function computePHashFromCanvas(imgLike) {
    const N = DCT_N;
    const c = document.createElement("canvas");
    c.width = N; c.height = N;
    const ctx = c.getContext("2d");
    ctx.drawImage(imgLike, 0, 0, N, N);
    const data = ctx.getImageData(0, 0, N, N).data;
    // Grayscale
    const gray = new Float64Array(N * N);
    for (let i = 0; i < N * N; i++)
      gray[i] = 0.299 * data[i * 4] + 0.587 * data[i * 4 + 1] + 0.114 * data[i * 4 + 2];

    // DCT rows
    const rows = new Float64Array(N * N);
    const rowIn = new Float64Array(N);
    const rowOut = new Float64Array(N);
    for (let y = 0; y < N; y++) {
      for (let x = 0; x < N; x++) rowIn[x] = gray[y * N + x];
      dct1d(rowIn, rowOut);
      for (let x = 0; x < N; x++) rows[y * N + x] = rowOut[x];
    }
    // DCT cols
    const dct = new Float64Array(N * N);
    const colIn = new Float64Array(N);
    const colOut = new Float64Array(N);
    for (let x = 0; x < N; x++) {
      for (let y = 0; y < N; y++) colIn[y] = rows[y * N + x];
      dct1d(colIn, colOut);
      for (let y = 0; y < N; y++) dct[y * N + x] = colOut[y];
    }
    // Top-left 8x8
    const low = new Float64Array(64);
    for (let y = 0; y < 8; y++)
      for (let x = 0; x < 8; x++) low[y * 8 + x] = dct[y * N + x];
    // Median excluding DC (index 0)
    const rest = Array.from(low).slice(1).sort((a, b) => a - b);
    const median = rest[Math.floor(rest.length / 2)];
    // 64-bit binary hash
    const hash = new Uint8Array(64);
    for (let i = 0; i < 64; i++) hash[i] = low[i] > median ? 1 : 0;
    return hash;
  }

  function hamming(h1, h2) {
    let d = 0;
    for (let i = 0; i < 64; i++) if (h1[i] !== h2[i]) d++;
    return d;
  }

  async function pHashFromBlob(blob) {
    const img = await loadImage(blob);
    return computePHashFromCanvas(img);
  }

  async function computeSharpness(blob) {
    const img = await loadImage(blob);
    // Downscale big images to keep it fast
    const maxDim = 512;
    let w = img.naturalWidth, h = img.naturalHeight;
    if (w > maxDim || h > maxDim) {
      const scale = maxDim / Math.max(w, h);
      w = Math.max(1, Math.round(w * scale));
      h = Math.max(1, Math.round(h * scale));
    }
    const c = document.createElement("canvas");
    c.width = w; c.height = h;
    const ctx = c.getContext("2d");
    ctx.drawImage(img, 0, 0, w, h);
    const data = ctx.getImageData(0, 0, w, h).data;
    const gray = new Float64Array(w * h);
    for (let i = 0; i < w * h; i++)
      gray[i] = 0.299 * data[i * 4] + 0.587 * data[i * 4 + 1] + 0.114 * data[i * 4 + 2];

    // PIL FIND_EDGES kernel: [-1,-1,-1;-1,8,-1;-1,-1,-1]
    let sum = 0, sumSq = 0, count = 0;
    for (let y = 1; y < h - 1; y++) {
      for (let x = 1; x < w - 1; x++) {
        const c8 = 8 * gray[y * w + x];
        const around =
          gray[(y - 1) * w + (x - 1)] + gray[(y - 1) * w + x] + gray[(y - 1) * w + (x + 1)] +
          gray[y * w + (x - 1)] + gray[y * w + (x + 1)] +
          gray[(y + 1) * w + (x - 1)] + gray[(y + 1) * w + x] + gray[(y + 1) * w + (x + 1)];
        const v = c8 - around;
        sum += v; sumSq += v * v; count++;
      }
    }
    if (count === 0) return 0;
    const mean = sum / count;
    return sumSq / count - mean * mean; // variance
  }

  async function getImageDimensions(blob) {
    const img = await loadImage(blob);
    return { width: img.naturalWidth, height: img.naturalHeight };
  }

  // ============================================================
  //  XLSX PARSING (SheetJS + JSZip)
  // ============================================================

  async function loadXlsxData(file) {
    const buf = await file.arrayBuffer();
    const workbook = window.XLSX.read(buf, { type: "array", cellDates: true });
    const zip = await window.JSZip.loadAsync(buf);
    return { workbook, zip };
  }

  function findSheetXmlPath(zip, workbook, sheetName) {
    // Read workbook.xml and rels to map sheet name -> xml file
    return Promise.all([
      zip.file("xl/workbook.xml").async("string"),
      zip.file("xl/_rels/workbook.xml.rels").async("string"),
    ]).then(([wbXml, relsXml]) => {
      const parser = new DOMParser();
      const wb = parser.parseFromString(wbXml, "text/xml");
      const rels = parser.parseFromString(relsXml, "text/xml");
      let rId = null;
      const sheets = wb.getElementsByTagName("sheet");
      for (let i = 0; i < sheets.length; i++) {
        if (sheets[i].getAttribute("name") === sheetName) {
          rId = sheets[i].getAttribute("r:id") || sheets[i].getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
          break;
        }
      }
      if (!rId) return null;
      const relsList = rels.getElementsByTagName("Relationship");
      for (let i = 0; i < relsList.length; i++) {
        if (relsList[i].getAttribute("Id") === rId) {
          let target = relsList[i].getAttribute("Target");
          if (target.startsWith("/")) target = target.substring(1);
          else target = "xl/" + target;
          return target;
        }
      }
      return null;
    });
  }

  async function findDrawingsForSheet(zip, sheetXmlPath) {
    const dir = sheetXmlPath.substring(0, sheetXmlPath.lastIndexOf("/"));
    const base = sheetXmlPath.substring(sheetXmlPath.lastIndexOf("/") + 1);
    const relsPath = `${dir}/_rels/${base}.rels`;
    const relsFile = zip.file(relsPath);
    if (!relsFile) return [];
    const relsXml = await relsFile.async("string");
    const parser = new DOMParser();
    const rels = parser.parseFromString(relsXml, "text/xml");
    const list = [];
    const relsNodes = rels.getElementsByTagName("Relationship");
    for (let i = 0; i < relsNodes.length; i++) {
      const type = relsNodes[i].getAttribute("Type") || "";
      if (type.indexOf("/drawing") !== -1) {
        let target = relsNodes[i].getAttribute("Target");
        // Resolve relative to sheet dir
        if (target.startsWith("/")) {
          target = target.substring(1);
        } else if (target.startsWith("../")) {
          const parent = dir.substring(0, dir.lastIndexOf("/"));
          target = parent + "/" + target.substring(3);
        } else {
          target = dir + "/" + target;
        }
        list.push(target);
      }
    }
    return list;
  }

  async function extractDrawingImages(zip, drawingPath) {
    const drawingXml = await zip.file(drawingPath).async("string");
    const dir = drawingPath.substring(0, drawingPath.lastIndexOf("/"));
    const base = drawingPath.substring(drawingPath.lastIndexOf("/") + 1);
    const relsPath = `${dir}/_rels/${base}.rels`;
    const relsFile = zip.file(relsPath);
    const embedMap = {};
    if (relsFile) {
      const relsXml = await relsFile.async("string");
      const parser = new DOMParser();
      const relsDoc = parser.parseFromString(relsXml, "text/xml");
      const nodes = relsDoc.getElementsByTagName("Relationship");
      for (let i = 0; i < nodes.length; i++) {
        const id = nodes[i].getAttribute("Id");
        let target = nodes[i].getAttribute("Target");
        if (target.startsWith("/")) target = target.substring(1);
        else if (target.startsWith("../")) {
          const parent = dir.substring(0, dir.lastIndexOf("/"));
          target = parent + "/" + target.substring(3);
        } else target = dir + "/" + target;
        embedMap[id] = target;
      }
    }

    const parser2 = new DOMParser();
    const doc = parser2.parseFromString(drawingXml, "text/xml");
    // Both namespaced and non-namespaced
    function collect(tag) {
      return Array.from(doc.getElementsByTagName(tag));
    }
    const anchors = [
      ...collect("xdr:oneCellAnchor"),
      ...collect("xdr:twoCellAnchor"),
      ...collect("xdr:absoluteAnchor"),
      ...collect("oneCellAnchor"),
      ...collect("twoCellAnchor"),
      ...collect("absoluteAnchor"),
    ];

    function firstChildText(node, ...tagAliases) {
      for (const t of tagAliases) {
        const els = node.getElementsByTagName(t);
        if (els.length) return els[0].textContent;
      }
      return null;
    }

    const results = [];
    for (const a of anchors) {
      // From cell
      let row = 0, col = 0;
      const from = a.getElementsByTagName("xdr:from")[0] || a.getElementsByTagName("from")[0];
      if (from) {
        const c = firstChildText(from, "xdr:col", "col");
        const r = firstChildText(from, "xdr:row", "row");
        col = parseInt(c || "0", 10);
        row = parseInt(r || "0", 10);
      }
      // pic > blipFill > blip r:embed
      const pic = a.getElementsByTagName("xdr:pic")[0] || a.getElementsByTagName("pic")[0];
      if (!pic) continue;
      const blipEls = pic.getElementsByTagName("a:blip");
      const blip = blipEls[0] || pic.getElementsByTagName("blip")[0];
      if (!blip) continue;
      const embed = blip.getAttribute("r:embed")
        || blip.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");
      if (!embed) continue;
      const mediaPath = embedMap[embed];
      if (!mediaPath) continue;
      results.push({ row: row + 1, col: col + 1, mediaPath });
    }
    return results;
  }

  function getCellValueWithMerges(sheet, cellRef) {
    const cell = sheet[cellRef];
    if (cell !== undefined) return cell;
    if (!sheet["!merges"]) return null;
    const target = window.XLSX.utils.decode_cell(cellRef);
    for (const range of sheet["!merges"]) {
      if (target.r >= range.s.r && target.r <= range.e.r &&
          target.c >= range.s.c && target.c <= range.e.c) {
        const masterRef = window.XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c });
        return sheet[masterRef] || null;
      }
    }
    return null;
  }

  function cellToDisplay(cell) {
    if (!cell) return "-";
    if (cell.t === "d" || cell.v instanceof Date) {
      const d = cell.v instanceof Date ? cell.v : new Date(cell.v);
      return d;
    }
    return cell.v;
  }

  function classifyExpiry(dateVal, alertDays) {
    if (!(dateVal instanceof Date) || isNaN(dateVal.getTime())) return "SEM DATA / OK";
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const target = new Date(dateVal);
    target.setHours(0, 0, 0, 0);
    const diffDays = Math.round((target.getTime() - today.getTime()) / 86400000);
    if (diffDays < 0) return "VENCIDA";
    if (diffDays === 0) return "VENCE HOJE";
    if (alertDays.includes(diffDays)) return `VENCE EM ${diffDays} DIA(S)`;
    return "OK (fora da janela de alerta)";
  }

  // Menor número = mais urgente. Usado para ordenar os resultados.
  function urgencyRank(status) {
    if (!status) return 999;
    if (status.includes("VENCIDA")) return 0;
    if (status.includes("VENCE HOJE")) return 1;
    const m = /VENCE EM (\d+) DIA/.exec(status);
    if (m) return 2 + parseInt(m[1], 10);
    return 999; // OK / SEM DATA
  }
  function sortByUrgency(matches) {
    return [...matches].sort((a, b) => urgencyRank(a.status_vencimento) - urgencyRank(b.status_vencimento));
  }

  function formatDateBR(d) {
    if (!(d instanceof Date) || isNaN(d.getTime())) return d != null ? String(d) : "-";
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const yy = d.getFullYear();
    return `${dd}/${mm}/${yy}`;
  }

  // ============================================================
  //  MAIN PROCESS
  // ============================================================
  async function runProcess() {
    if (!state.xlsxFile) return toast("Selecione o arquivo Excel", "error");
    const isFolder = state.searchMode === "folder";
    if (isFolder && state.folderFiles.length === 0) return toast("Selecione uma pasta com imagens", "error");
    if (!isFolder && !state.imageFile) return toast("Selecione a imagem de pesquisa", "error");

    state.aborted = false;
    state.logs = [];
    state.matches = [];
    state.failures = [];
    state.imageMeta = null;
    state.runId = uuid();
    state.pendingEdits = {};
    state.mediaIndex = {};
    termBody.innerHTML = "";
    termLinesCount.textContent = "0";
    $("summary-card").classList.add("hidden");
    updateEditControls();
    setRunning(true, "iniciando");

    const searchFiles = isFolder ? state.folderFiles : [state.imageFile];
    const params = {
      sheet_name: $("p-sheet").value.trim(),
      description_col: $("p-desc-col").value.trim().toUpperCase(),
      expiry_col: $("p-exp-col").value.trim().toUpperCase(),
      min_similarity: parseFloat($("p-min-sim").value) || 90,
      expiry_alert_days: $("p-alert-days").value.split(",").map((s) => parseInt(s.trim(), 10)).filter((n) => !isNaN(n)),
      sharpness_threshold: parseFloat($("p-sharp").value) || 500,
      xlsx_filename: state.xlsxFile.name,
      image_filename: isFolder ? `${searchFiles.length} imagem(ns) da pasta` : state.imageFile.name,
      search_mode: state.searchMode,
    };

    // Reset preview area
    $("mp-grid").innerHTML = "";
    $("matches-preview").classList.add("hidden");
    $("mp-count-badge").textContent = "0";

    try {
      // ---- Reference image(s) info ----
      log("header", "=".repeat(60));
      log("header", isFolder ? `PESQUISA POR PASTA - ${searchFiles.length} imagem(ns)` : `INFORMACOES DA IMAGEM - ${params.image_filename}`);
      log("header", "=".repeat(60));

      if (!isFolder) {
        const refDim = await getImageDimensions(state.imageFile);
        const refKb = Math.round((state.imageFile.size / 1024) * 100) / 100;
        const refSharp = await computeSharpness(state.imageFile);
        const sharpOk = refSharp >= params.sharpness_threshold;
        const sharpLbl = sharpOk ? "Boa nitidez" : "Nitidez baixa (imagem pode estar borrada ou em baixa resolucao)";
        log("info", `Dimensoes : ${refDim.width}x${refDim.height} px`);
        log("info", `Peso      : ${refKb} KB`);
        log(sharpOk ? "success" : "warn", `Nitidez   : ${sharpLbl} (score: ${Math.round(refSharp * 10) / 10})`);
        state.imageMeta = {
          filename: params.image_filename, width: refDim.width, height: refDim.height,
          weight_kb: refKb, sharpness_score: Math.round(refSharp * 10) / 10, sharpness_ok: sharpOk,
        };
      } else {
        log("info", `Pasta selecionada com ${searchFiles.length} imagem(ns) para pesquisar.`);
        state.imageMeta = { filename: params.image_filename, folder: true, count: searchFiles.length };
      }
      log("header", "=".repeat(60));
      await yieldTick();
      if (state.aborted) return finishAborted();

      // ---- Open workbook ----
      log("info", `Abrindo planilha: ${params.xlsx_filename}`);
      const { workbook, zip } = await loadXlsxData(state.xlsxFile);
      if (!workbook.SheetNames.includes(params.sheet_name)) {
        log("error", `Aba '${params.sheet_name}' nao encontrada. Abas disponiveis: ${workbook.SheetNames.join(", ")}`);
        return finalize(false, `Aba ${params.sheet_name} nao encontrada`);
      }
      const sheet = workbook.Sheets[params.sheet_name];

      // ---- Locate drawings/images ----
      setRunning(true, "extraindo imagens");
      const sheetXmlPath = await findSheetXmlPath(zip, workbook, params.sheet_name);
      if (!sheetXmlPath) {
        log("error", "Nao foi possivel localizar o XML da aba.");
        return finalize(false, "Sheet XML nao encontrado");
      }
      const drawings = await findDrawingsForSheet(zip, sheetXmlPath);
      if (drawings.length === 0) {
        log("error", "Nenhuma imagem encontrada na aba.");
        return finalize(true, null);
      }
      const allImgs = [];
      for (const d of drawings) {
        try {
          const list = await extractDrawingImages(zip, d);
          allImgs.push(...list);
        } catch (err) {
          log("warn", `Falha ao ler drawing ${d}: ${err.message}`);
        }
      }
      if (allImgs.length === 0) {
        log("error", "Nenhuma imagem embutida foi encontrada na aba.");
        return finalize(true, null);
      }

      log("info", `Imagens encontradas na planilha: ${allImgs.length}`);
      log("info", "Calculando hash de cada imagem da planilha (uma unica vez)...");
      await yieldTick();

      // ---- Pré-computa hash + metadados de cada imagem embutida (uma única vez) ----
      setRunning(true, "extraindo imagens");
      const catalog = []; // { row, col, celula, mediaPath, hash }
      for (let i = 0; i < allImgs.length; i++) {
        if (state.aborted) return finishAborted();
        const { row, col, mediaPath } = allImgs[i];
        const celula = `${colIndexToLetter(col)}${row}`;
        state.mediaIndex[celula] = { mediaPath, row, col };
        const mediaFile = zip.file(mediaPath);
        if (!mediaFile) {
          state.failures.push(`Imagem #${i + 1} (linha ${row}): media '${mediaPath}' nao encontrada`);
          log("warn", `Imagem #${i + 1} (linha ${row}): media nao encontrada (${mediaPath})`);
          continue;
        }
        try {
          const blob = await mediaFile.async("blob");
          const hash = await pHashFromBlob(blob);
          catalog.push({ row, col, celula, mediaPath, hash });
        } catch (err) {
          state.failures.push(`Imagem #${i + 1} (linha ${row}): falha ao decodificar (${err.message})`);
          log("warn", `Imagem #${i + 1} (linha ${row}): falha ao decodificar (${err.message})`);
        }
        if ((i + 1) % 40 === 0) await yieldTick();
      }

      log("info", "Comparando com a(s) imagem(ns) de pesquisa...");
      log("dim", "");
      await yieldTick();

      // ---- Compare each search image against the catalog ----
      setRunning(true, "comparando");
      const foundCelulas = new Set();
      for (let si = 0; si < searchFiles.length; si++) {
        if (state.aborted) return finishAborted();
        const searchFile = searchFiles[si];
        let refHash;
        try {
          refHash = await pHashFromBlob(searchFile);
        } catch (err) {
          state.failures.push(`Imagem de pesquisa '${searchFile.name}': falha ao decodificar (${err.message})`);
          log("warn", `Imagem de pesquisa '${searchFile.name}': falha ao decodificar (${err.message})`);
          continue;
        }

        for (let i = 0; i < catalog.length; i++) {
          if (state.aborted) return finishAborted();
          const { row, col, celula, mediaPath, hash } = catalog[i];
          const distancia = hamming(refHash, hash);
          const similaridade = Math.round(((64 - distancia) / 64) * 10000) / 100;
          if (distancia !== 0 && similaridade < params.min_similarity) {
            if ((i + 1) % 60 === 0) await yieldTick();
            continue;
          }
          // Numa pasta, evita reportar a mesma célula duas vezes (fica com o melhor match)
          if (isFolder && foundCelulas.has(celula)) continue;

          const isIdentical = distancia === 0 || similaridade >= 95;
          const tipo = isIdentical ? "identica" : "semelhante";

          const descRef = params.description_col + row;
          const expRef = params.expiry_col + row;
          const descCell = getCellValueWithMerges(sheet, descRef);
          const expCell = getCellValueWithMerges(sheet, expRef);
          const descricao = descCell?.v != null ? String(descCell.v) : "-";
          const expValue = expCell ? cellToDisplay(expCell) : null;
          const status = classifyExpiry(expValue instanceof Date ? expValue : null, params.expiry_alert_days);
          const vencimento = expValue instanceof Date ? formatDateBR(expValue) : (expValue != null ? String(expValue) : "-");

          let thumbnailUrl = null, thumbnailData = null;
          try {
            const mediaFile = zip.file(mediaPath);
            const mediaBlob = await mediaFile.async("blob");
            thumbnailUrl = URL.createObjectURL(mediaBlob);
            thumbnailData = await makeThumbnailDataUrl(mediaBlob, 200);
          } catch (_) { /* ignora falha no thumb */ }

          const match = {
            celula, linha: row, coluna: col, distancia, similaridade, descricao,
            vencimento, status_vencimento: status, tipo,
            thumbnailUrl, thumbnailData,
            pesquisaImagem: isFolder ? searchFile.name : null,
          };
          state.matches.push(match);
          foundCelulas.add(celula);

          const tag = distancia === 0 ? "IDENTICA" : isIdentical ? `IDENTICA (${similaridade}%)` : `SEMELHANTE (${similaridade}%)`;
          const level = isIdentical ? "success" : "warn";
          log(level, `>> ${tag} encontrada${isFolder ? ` (pesquisa: ${searchFile.name})` : ""}`);
          log("info", `   Celula da imagem : ${celula}`);
          log("info", `   Descricao        : ${descricao}`);
          log("info", `   Vencimento       : ${vencimento}`);
          log("info", `   Status           : ${status}`);
          log("dim", "-".repeat(60));
          await yieldTick();
        }
      }

      // ---- Ordena por urgência de vencimento e renderiza tudo de uma vez ----
      state.matches = sortByUrgency(state.matches);
      $("mp-grid").innerHTML = "";
      state.matches.forEach(renderMatchCard);

      // ---- Guarda referências para permitir edição/exportação depois ----
      state.currentZip = zip;
      state.currentWorkbook = workbook;
      state.currentParams = params;
      state.currentSheetXmlPath = sheetXmlPath;

      // ---- Final report ----
      log("header", "=".repeat(60));
      log("header", "RESULTADO DA BUSCA");
      log("header", "=".repeat(60));
      if (state.matches.length === 0) log("warn", "Nenhuma ocorrencia da imagem foi encontrada na planilha.");
      else log("success", `${state.matches.length} ocorrencia(s) encontrada(s), ordenadas por urgencia de vencimento.`);
      if (state.failures.length > 0) {
        log("warn", `${state.failures.length} imagem(ns) nao puderam ser processadas:`);
        state.failures.forEach((m) => log("warn", `   - ${m}`));
      }
      log("header", "=".repeat(60));

      finalize(true, null, params);
      renderSheetTable();
    } catch (err) {
      console.error(err);
      log("error", `Erro fatal: ${err.message}`);
      finalize(false, err.message);
    } finally {
      setRunning(false);
    }
  }

  function finishAborted() {
    log("warn", "Execucao interrompida pelo usuario.");
    finalize(false, "interrompida");
    setRunning(false);
  }

  function finalize(success, error, params) {
    const run = {
      id: state.runId,
      started_at: state._startedAt || new Date().toISOString(),
      finished_at: new Date().toISOString(),
      params: params || collectParams(),
      matches: state.matches,
      failures: state.failures,
      image_meta: state.imageMeta,
      success,
      error,
    };
    state.lastRun = run;
    if (success && (state.matches.length > 0 || state.failures.length > 0 || state.imageMeta)) {
      saveRun(run);
      toast("Execução salva no histórico", "success");
    }
    if (success && state.matches.length > 0) renderSummary(state.matches);
  }

  function collectParams() {
    return {
      sheet_name: $("p-sheet").value.trim(),
      description_col: $("p-desc-col").value.trim().toUpperCase(),
      expiry_col: $("p-exp-col").value.trim().toUpperCase(),
      min_similarity: parseFloat($("p-min-sim").value) || 90,
      xlsx_filename: state.xlsxFile?.name,
      image_filename: state.imageFile?.name,
    };
  }

  function renderSummary(matches) {
    $("stat-total").textContent = matches.length;
    $("stat-identical").textContent = matches.filter((m) => m.tipo === "identica").length;
    $("stat-similar").textContent = matches.filter((m) => m.tipo === "semelhante").length;
    $("summary-card").classList.remove("hidden");
    $("stat-identical").classList.add("stat-green");
  }

  // ============================================================
  //  THUMBNAILS + MATCH CARDS + LIGHTBOX
  // ============================================================
  function statusVariant(status) {
    if (!status) return "st-green";
    if (status.includes("VENCIDA") || status.includes("VENCE HOJE")) return "st-red";
    if (status.startsWith("VENCE EM")) return "st-amber";
    return "st-green";
  }

  async function makeThumbnailDataUrl(blob, maxDim) {
    if (!maxDim) maxDim = 200;
    try {
      const img = await loadImage(blob);
      const c = document.createElement("canvas");
      const scale = Math.min(1, maxDim / Math.max(img.naturalWidth, img.naturalHeight));
      c.width = Math.max(1, Math.round(img.naturalWidth * scale));
      c.height = Math.max(1, Math.round(img.naturalHeight * scale));
      c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
      return c.toDataURL("image/jpeg", 0.7);
    } catch { return null; }
  }

  function renderMatchCard(m) {
    const grid = $("mp-grid");
    $("matches-preview").classList.remove("hidden");

    const isIdentical = m.tipo === "identica";
    const tagCls = isIdentical ? "tag-identical" : "tag-similar";
    const tagText = isIdentical ? "Idêntica" : "Semelhante";
    const simCls = m.similaridade >= 95 ? "high" : (m.similaridade >= 85 ? "mid" : "");
    const stCls = statusVariant(m.status_vencimento);
    const src = m.thumbnailUrl || m.thumbnailData || "";

    const edited = !!state.pendingEdits[m.celula];
    const searchTag = m.pesquisaImagem ? `<div class="mp-cell" style="margin-top:-2px">de: ${escapeHtml(m.pesquisaImagem)}</div>` : "";

    const card = document.createElement("div");
    card.className = "mp-card";
    card.dataset.testid = `match-card-${m.celula}`;
    card.innerHTML = `
      <div class="mp-thumb" data-lightbox="${escapeHtml(src)}" data-caption="Célula ${escapeHtml(m.celula)} · ${escapeHtml(m.descricao)}">
        ${src
          ? `<img src="${src}" alt="preview ${escapeHtml(m.celula)}" loading="lazy" />`
          : `<span style="color:#999;font-size:12px">sem preview</span>`}
        <span class="mp-thumb-tag ${tagCls}">${tagText}</span>
        <span class="mp-thumb-sim ${simCls}">${m.similaridade}%</span>
        <button type="button" class="mp-edit-btn" data-edit="${escapeHtml(m.celula)}" title="Editar linha">✏️</button>
        ${edited ? `<span class="mp-edited-badge">Editado</span>` : ""}
      </div>
      <div class="mp-info">
        <div class="mp-cell">CÉLULA ${escapeHtml(m.celula)} · LINHA ${m.linha}</div>
        ${searchTag}
        <div class="mp-desc">${escapeHtml(m.descricao)}</div>
        <div class="mp-exp">
          <span class="mp-exp-date">${escapeHtml(m.vencimento)}</span>
          <span class="mp-exp-status ${stCls}">${escapeHtml(m.status_vencimento)}</span>
        </div>
      </div>
    `;
    grid.appendChild(card);
    $("mp-count-badge").textContent = state.matches.length;

    const thumb = card.querySelector(".mp-thumb");
    on(thumb, "click", (e) => {
      if (e.target.closest(".mp-edit-btn")) return;
      openLightbox(thumb.dataset.lightbox, thumb.dataset.caption);
    });
    const editBtn = card.querySelector(".mp-edit-btn");
    on(editBtn, "click", (e) => { e.stopPropagation(); openEditModal(m.celula); });
  }

  function openLightbox(src, caption) {
    if (!src) return;
    $("lightbox-img").src = src;
    $("lightbox-caption").textContent = caption || "";
    $("modal-lightbox").classList.remove("hidden");
  }

  on($("btn-run"), "click", () => {
    state._startedAt = new Date().toISOString();
    runProcess();
  });
  on($("btn-stop"), "click", () => { state.aborted = true; });

  // ============================================================
  //  CSV EXPORT (client-side)
  // ============================================================
  function downloadCsv(matches, runId) {
    if (!matches?.length) return toast("Sem matches para exportar", "warn");
    const header = ["Tipo", "Celula", "Linha", "Similaridade (%)", "Descricao", "Vencimento", "Status Vencimento"];
    const rows = matches.map((m) => [m.tipo, m.celula, m.linha, m.similaridade, m.descricao, m.vencimento, m.status_vencimento]);
    const csv = "\ufeff" + [header, ...rows]
      .map((r) => r.map((c) => `"${String(c ?? "").replace(/"/g, '""')}"`).join(";"))
      .join("\r\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `modal-card-${String(runId).slice(0, 8)}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  }

  on($("btn-export-csv"), "click", () => {
    if (state.lastRun) downloadCsv(state.lastRun.matches, state.lastRun.id);
    else if (state.matches.length) downloadCsv(state.matches, state.runId);
  });

  // ============================================================
  //  HISTORY (localStorage)
  // ============================================================
  function loadHistory() {
    try { return JSON.parse(localStorage.getItem(HISTORY_KEY) || "[]"); }
    catch { return []; }
  }
  function persistHistory(list) {
    try { localStorage.setItem(HISTORY_KEY, JSON.stringify(list.slice(0, 100))); }
    catch (e) { toast("Não foi possível salvar no localStorage (talvez cheio)", "warn"); }
  }
  function saveRun(run) {
    // Persistimos só o thumbnailData (base64 pequeno); descartamos blob URLs voláteis
    const clone = {
      ...run,
      matches: (run.matches || []).map((m) => ({
        celula: m.celula, linha: m.linha, distancia: m.distancia,
        similaridade: m.similaridade, descricao: m.descricao,
        vencimento: m.vencimento, status_vencimento: m.status_vencimento,
        tipo: m.tipo, thumbnailData: m.thumbnailData || null,
      })),
    };
    const list = loadHistory();
    list.unshift(clone);
    persistHistory(list);
  }
  function deleteRun(id) {
    persistHistory(loadHistory().filter((r) => r.id !== id));
    renderHistory(loadHistory());
    toast("Execução excluída", "success");
  }
  function clearAllHistory() {
    if (!confirm("Excluir todas as execuções?")) return;
    localStorage.removeItem(HISTORY_KEY);
    renderHistory([]);
    toast("Histórico limpo", "success");
  }

  function renderHistory(runs) {
    const list = $("history-list");
    const empty = $("history-empty");
    list.innerHTML = "";
    $("history-count").textContent = `${runs.length} registro(s)`;
    empty.classList.toggle("hidden", runs.length > 0);

    runs.forEach((r) => {
      const card = document.createElement("div");
      card.className = "history-item";
      card.dataset.testid = `history-item-${r.id}`;
      const matchCount = r.matches?.length || 0;
      const failCount = r.failures?.length || 0;
      const badge = r.success
        ? '<span class="badge badge-success">Concluído</span>'
        : '<span class="badge badge-fail">Falhou</span>';
      card.innerHTML = `
        <div class="history-item-top">
          <div style="flex:1; min-width:0">
            <div class="hi-meta">
              ${badge}
              <span class="hi-time">${formatDate(r.started_at)}</span>
            </div>
            <div class="hi-files">
              <div class="hi-file">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
                <span class="name">${escapeHtml(r.params?.xlsx_filename || "-")}</span>
                <span class="dim">· ${escapeHtml(r.params?.sheet_name || "")}</span>
              </div>
              <div class="hi-file">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg>
                <span class="name">${escapeHtml(r.params?.image_filename || "-")}</span>
              </div>
            </div>
            <div class="hi-stats">
              <span><b>${matchCount}</b> match(es)</span>
              ${failCount > 0 ? `<span style="color:#b45309"><b>${failCount}</b> falha(s)</span>` : ""}
            </div>
          </div>
          <div class="hi-actions">
            <button class="hi-btn" data-view="${r.id}" title="Visualizar">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
            </button>
            <button class="hi-btn" data-csv="${r.id}" title="Exportar CSV">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            </button>
            <button class="hi-btn danger" data-del="${r.id}" title="Excluir">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-2 14a2 2 0 0 1-2 2H9a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6M14 11v6"/></svg>
            </button>
          </div>
        </div>
      `;
      list.appendChild(card);
    });

    list.querySelectorAll("[data-view]").forEach((b) =>
      on(b, "click", () => openDetail(b.dataset.view)));
    list.querySelectorAll("[data-csv]").forEach((b) =>
      on(b, "click", () => {
        const r = loadHistory().find((x) => x.id === b.dataset.csv);
        if (r) downloadCsv(r.matches, r.id);
      }));
    list.querySelectorAll("[data-del]").forEach((b) =>
      on(b, "click", () => deleteRun(b.dataset.del)));
  }

  function openDetail(id) {
    const r = loadHistory().find((x) => x.id === id);
    if (!r) return;
    const p = r.params || {};
    const matches = r.matches || [];
    let matchesHtml = "";
    if (matches.length) {
      matchesHtml = `
        <h4 style="margin:16px 0 8px;font-size:13px;font-weight:700">Matches encontrados</h4>
        <div style="overflow-x:auto">
          <table class="detail-table">
            <thead><tr><th></th><th>Tipo</th><th>Célula</th><th>Sim.</th><th>Descrição</th><th>Vencimento</th><th>Status</th></tr></thead>
            <tbody>
              ${matches.map((m) => `
                <tr>
                  <td>${m.thumbnailData ? `<img src="${m.thumbnailData}" alt="thumb" style="width:52px;height:40px;object-fit:contain;border-radius:4px;background:#f5f5f5;cursor:zoom-in" data-thumb="${escapeHtml(m.thumbnailData)}" data-caption="Célula ${escapeHtml(m.celula)} · ${escapeHtml(m.descricao)}" />` : ""}</td>
                  <td><span class="badge ${m.tipo === "identica" ? "badge-success" : "badge-warn"}">${escapeHtml(m.tipo)}</span></td>
                  <td class="mono">${escapeHtml(m.celula)}</td>
                  <td>${m.similaridade}%</td>
                  <td>${escapeHtml(m.descricao)}</td>
                  <td>${escapeHtml(m.vencimento)}</td>
                  <td style="color:var(--bradesco-red);font-weight:600">${escapeHtml(m.status_vencimento)}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
      `;
    }
    $("detail-body").innerHTML = `
      <div class="detail-meta">
        <div><b>Início:</b> ${formatDate(r.started_at)}</div>
        <div><b>Fim:</b> ${formatDate(r.finished_at)}</div>
        <div><b>Planilha:</b> ${escapeHtml(p.xlsx_filename || "-")}</div>
        <div><b>Aba:</b> ${escapeHtml(p.sheet_name || "-")}</div>
        <div><b>Imagem:</b> ${escapeHtml(p.image_filename || "-")}</div>
        <div><b>Similaridade mín:</b> ${p.min_similarity ?? "-"}%</div>
      </div>
      ${matchesHtml}
      ${matches.length ? `<button class="btn-primary" style="margin-top:16px" id="detail-csv">Exportar CSV</button>` : ""}
    `;
    $("modal-detail").classList.remove("hidden");
    const btn = document.getElementById("detail-csv");
    if (btn) btn.onclick = () => downloadCsv(matches, r.id);
    document.querySelectorAll("#detail-body [data-thumb]").forEach((el) => {
      on(el, "click", () => openLightbox(el.dataset.thumb, el.dataset.caption));
    });
  }

  on($("btn-refresh-history"), "click", () => renderHistory(loadHistory()));
  on($("btn-clear-history"), "click", clearAllHistory);

  // ============================================================
  //  EDIÇÃO DE LINHAS (fica em memória até exportar)
  // ============================================================
  function getMatchByCelula(celula) {
    return state.matches.find((m) => m.celula === celula);
  }

  function updateEditControls() {
    const count = Object.keys(state.pendingEdits).length;
    const lbl = $("edit-count-label");
    if (lbl) lbl.textContent = `${count} edição(ões) pendente(s)`;
    const disabled = count === 0;
    ["btn-save-xlsx", "btn-save-xlsx-2"].forEach((id) => { const b = $(id); if (b) b.disabled = disabled; });
  }

  function refreshMatchCardsAfterEdit() {
    $("mp-grid").innerHTML = "";
    state.matches.forEach(renderMatchCard);
    const total = $("stat-total"); if (total && !$("summary-card").classList.contains("hidden")) renderSummary(state.matches);
  }

  let editModalCelula = null;

  function addEditFieldRow(col, value, isDate) {
    const row = document.createElement("div");
    row.className = "edit-field-row";
    row.innerHTML = `
      <input type="text" class="ef-col" maxlength="3" value="${escapeHtml(col || "")}" placeholder="Col." />
      <input type="text" class="ef-val" value="${escapeHtml(value ?? "")}" placeholder="${isDate ? "dd/mm/aaaa" : "Valor"}" data-isdate="${isDate ? "1" : "0"}" />
      <button type="button" class="edit-field-remove" title="Remover">×</button>
    `;
    row.querySelector(".edit-field-remove").addEventListener("click", () => row.remove());
    $("edit-fields-list").appendChild(row);
  }
  on($("btn-add-field"), "click", () => addEditFieldRow("", "", false));

  function openEditModal(celula) {
    const m = getMatchByCelula(celula);
    if (!m) return;
    editModalCelula = celula;
    $("edit-cell-label").textContent = `${celula} (linha ${m.linha})`;
    $("edit-image-current").src = m.thumbnailUrl || m.thumbnailData || "";
    $("edit-image-input").value = "";
    $("edit-image-newname").classList.add("hidden");
    $("edit-image-newname").textContent = "";

    const existing = state.pendingEdits[celula];
    const list = $("edit-fields-list");
    list.innerHTML = "";

    const params = state.currentParams || {};
    const descCol = params.description_col || "B";
    const expCol = params.expiry_col || "F";

    const fieldsToShow = [
      { col: descCol, value: existing?.cellEdits?.[descCol]?.value ?? m.descricao, isDate: false },
      { col: expCol, value: existing?.cellEdits?.[expCol]?.value ?? m.vencimento, isDate: true },
    ];
    if (existing) {
      Object.keys(existing.cellEdits || {}).forEach((c) => {
        if (c !== descCol && c !== expCol) {
          fieldsToShow.push({ col: c, value: existing.cellEdits[c].value, isDate: existing.cellEdits[c].isDate });
        }
      });
    }
    fieldsToShow.forEach((f) => addEditFieldRow(f.col, f.value, f.isDate));

    $("modal-edit").classList.remove("hidden");
  }

  function parseBRDate(str) {
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(str).trim());
    if (!m) return null;
    const d = new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
    return isNaN(d.getTime()) ? null : d;
  }

  on($("edit-image-input"), "change", (e) => {
    const f = e.target.files && e.target.files[0];
    const nameEl = $("edit-image-newname");
    if (f) {
      $("edit-image-current").src = URL.createObjectURL(f);
      nameEl.textContent = `Nova imagem selecionada: ${f.name}`;
      nameEl.classList.remove("hidden");
    } else {
      nameEl.classList.add("hidden");
    }
  });

  on($("btn-save-edit"), "click", () => {
    if (!editModalCelula) return;
    const celula = editModalCelula;
    const rows = Array.from(document.querySelectorAll("#edit-fields-list .edit-field-row"));
    const cellEdits = {};
    for (const row of rows) {
      const col = row.querySelector(".ef-col").value.trim().toUpperCase();
      const valInput = row.querySelector(".ef-val");
      const val = valInput.value;
      const isDate = valInput.dataset.isdate === "1";
      if (!col || !/^[A-Z]{1,3}$/.test(col)) continue;
      if (isDate) {
        if (val.trim() === "") continue;
        const d = parseBRDate(val);
        if (!d) { toast(`Data inválida na coluna ${col} (use dd/mm/aaaa)`, "error"); return; }
        cellEdits[col] = { value: val, isDate: true, dateObj: d };
      } else {
        cellEdits[col] = { value: val, isDate: false };
      }
    }
    const imageInput = $("edit-image-input");
    const newImageFile = imageInput.files && imageInput.files[0] ? imageInput.files[0] : null;
    const imageFile = newImageFile || state.pendingEdits[celula]?.imageFile || null;

    if (Object.keys(cellEdits).length === 0 && !imageFile) {
      delete state.pendingEdits[celula];
    } else {
      state.pendingEdits[celula] = {
        cellEdits,
        imageFile,
        imageMediaPath: state.mediaIndex[celula]?.mediaPath || null,
      };
    }

    // Atualiza a exibição (card + tabela) imediatamente, mesmo sem exportar ainda
    const m = getMatchByCelula(celula);
    const params = state.currentParams || {};
    if (m) {
      if (cellEdits[params.description_col]) m.descricao = cellEdits[params.description_col].value;
      if (cellEdits[params.expiry_col]) {
        m.vencimento = cellEdits[params.expiry_col].value;
        const d = cellEdits[params.expiry_col].dateObj;
        if (d) m.status_vencimento = classifyExpiry(d, params.expiry_alert_days || [15, 7, 3, 0]);
      }
      if (newImageFile) m.thumbnailUrl = URL.createObjectURL(newImageFile);
    }
    state.matches = sortByUrgency(state.matches);

    updateEditControls();
    refreshMatchCardsAfterEdit();
    renderSheetTable();
    $("modal-edit").classList.add("hidden");
    toast("Edição salva na memória. Clique em 'Salvar alterações no Excel' para gerar o arquivo.", "success");
  });

  // ============================================================
  //  ABA "PLANILHA" — visualizar/editar sem reabrir o Excel
  // ============================================================
  function renderSheetTable() {
    const matches = state.matches || [];
    const wrap = $("sheet-table-wrap");
    const empty = $("sheet-empty");
    const body = $("sheet-table-body");
    const countLbl = $("sheet-count");

    if (!matches.length) {
      wrap.classList.add("hidden");
      empty.classList.remove("hidden");
      countLbl.textContent = "Nenhum resultado carregado ainda — rode uma pesquisa na aba Execução.";
      return;
    }
    wrap.classList.remove("hidden");
    empty.classList.add("hidden");

    const filterVal = ($("sheet-filter")?.value || "").toLowerCase().trim();
    const filtered = filterVal
      ? matches.filter((m) => [m.celula, m.descricao, m.status_vencimento, m.vencimento].join(" ").toLowerCase().includes(filterVal))
      : matches;

    countLbl.textContent = `${filtered.length} de ${matches.length} linha(s) · ordenado por urgência de vencimento`;

    body.innerHTML = "";
    filtered.forEach((m) => {
      const edited = !!state.pendingEdits[m.celula];
      const rank = urgencyRank(m.status_vencimento);
      const dotColor = rank <= 1 ? "#cc092f" : rank < 999 ? "#f2a825" : "#22c55e";
      const src = m.thumbnailUrl || m.thumbnailData || "";
      const tr = document.createElement("tr");
      if (edited) tr.classList.add("row-edited");
      tr.innerHTML = `
        <td>${src ? `<img class="sheet-thumb" src="${src}" data-lightbox="${escapeHtml(src)}" data-caption="Célula ${escapeHtml(m.celula)}" />` : ""}</td>
        <td><span class="sheet-urgency-dot" style="background:${dotColor}"></span></td>
        <td class="mono">${escapeHtml(m.celula)}</td>
        <td class="sheet-cell-desc" title="${escapeHtml(m.descricao)}">${escapeHtml(m.descricao)}${edited ? ' <b style="color:#15803d">●</b>' : ""}</td>
        <td>${escapeHtml(m.vencimento)}</td>
        <td><span class="mp-exp-status ${statusVariant(m.status_vencimento)}">${escapeHtml(m.status_vencimento)}</span></td>
        <td><button class="sheet-row-edit-btn" data-edit="${escapeHtml(m.celula)}" title="Editar">✏️</button></td>
      `;
      body.appendChild(tr);
    });

    body.querySelectorAll("[data-edit]").forEach((b) => on(b, "click", () => openEditModal(b.dataset.edit)));
    body.querySelectorAll("[data-lightbox]").forEach((img) =>
      on(img, "click", () => openLightbox(img.dataset.lightbox, img.dataset.caption)));
  }
  on($("sheet-filter"), "input", () => renderSheetTable());

  // ============================================================
  //  EXPORTAÇÃO — aplica edições + backup automático + baixa .xlsx
  // ============================================================
  const SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const RELS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  function excelSerialFromDate(d) {
    const epoch = Date.UTC(1899, 11, 30);
    const utcDate = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
    return Math.round((utcDate - epoch) / 86400000);
  }

  function resolveMasterCellRef(sheet, ref) {
    if (!sheet || !sheet["!merges"]) return ref;
    const target = window.XLSX.utils.decode_cell(ref);
    for (const range of sheet["!merges"]) {
      if (target.r >= range.s.r && target.r <= range.e.r && target.c >= range.s.c && target.c <= range.e.c) {
        return window.XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c });
      }
    }
    return ref;
  }

  function applyCellEditsToXmlDoc(doc, edits) {
    const sheetData = doc.getElementsByTagName("sheetData")[0];
    edits.forEach(({ ref, value, isDate, dateObj }) => {
      const rowMatch = /^([A-Z]+)(\d+)$/.exec(ref);
      if (!rowMatch) return;
      const rowNum = parseInt(rowMatch[2], 10);
      let cellEl = null;
      const cells = doc.getElementsByTagName("c");
      for (let i = 0; i < cells.length; i++) {
        if (cells[i].getAttribute("r") === ref) { cellEl = cells[i]; break; }
      }
      if (!cellEl && sheetData) {
        let rowEl = null;
        const rows = doc.getElementsByTagName("row");
        for (let i = 0; i < rows.length; i++) {
          if (parseInt(rows[i].getAttribute("r"), 10) === rowNum) { rowEl = rows[i]; break; }
        }
        if (!rowEl) {
          rowEl = doc.createElementNS(SPREADSHEET_NS, "row");
          rowEl.setAttribute("r", String(rowNum));
          let inserted = false;
          for (let i = 0; i < rows.length; i++) {
            if (parseInt(rows[i].getAttribute("r"), 10) > rowNum) { sheetData.insertBefore(rowEl, rows[i]); inserted = true; break; }
          }
          if (!inserted) sheetData.appendChild(rowEl);
        }
        cellEl = doc.createElementNS(SPREADSHEET_NS, "c");
        cellEl.setAttribute("r", ref);
        const colIdx = colLetterToIndex(rowMatch[1]);
        let inserted = false;
        const rowCells = Array.from(rowEl.getElementsByTagName("c"));
        for (const existingC of rowCells) {
          const existingCol = colLetterToIndex((existingC.getAttribute("r") || "A0").replace(/\d+$/, ""));
          if (existingCol > colIdx) { rowEl.insertBefore(cellEl, existingC); inserted = true; break; }
        }
        if (!inserted) rowEl.appendChild(cellEl);
      }
      if (!cellEl) return;

      while (cellEl.firstChild) cellEl.removeChild(cellEl.firstChild);
      cellEl.removeAttribute("t");

      if (isDate && dateObj) {
        const v = doc.createElementNS(SPREADSHEET_NS, "v");
        v.textContent = String(excelSerialFromDate(dateObj));
        cellEl.appendChild(v);
      } else {
        cellEl.setAttribute("t", "inlineStr");
        const is = doc.createElementNS(SPREADSHEET_NS, "is");
        const t = doc.createElementNS(SPREADSHEET_NS, "t");
        t.textContent = String(value ?? "");
        is.appendChild(t);
        cellEl.appendChild(is);
      }
    });
  }

  async function addBackupSheet(zip, sheetXmlPath, originalSheetName, sheetXmlString) {
    const existingPaths = Object.keys(zip.files).filter((p) => /^xl\/worksheets\/sheet\d+\.xml$/.test(p));
    let maxN = 0;
    existingPaths.forEach((p) => { const m = /sheet(\d+)\.xml$/.exec(p); if (m) maxN = Math.max(maxN, parseInt(m[1], 10)); });
    const newN = maxN + 1;
    const newSheetPath = `xl/worksheets/sheet${newN}.xml`;

    const stamp = nowTs().replace(/:/g, "");
    let backupName = `Backup_${originalSheetName}_${stamp}`.replace(/[\[\]\*\?\/\\:]/g, "").slice(0, 31) || `Backup_${stamp}`;

    zip.file(newSheetPath, sheetXmlString);

    const dir = sheetXmlPath.substring(0, sheetXmlPath.lastIndexOf("/"));
    const base = sheetXmlPath.substring(sheetXmlPath.lastIndexOf("/") + 1);
    const origRelsPath = `${dir}/_rels/${base}.rels`;
    const origRelsFile = zip.file(origRelsPath);
    if (origRelsFile) {
      const relsXml = await origRelsFile.async("string");
      zip.file(`xl/worksheets/_rels/sheet${newN}.xml.rels`, relsXml);
    }

    const parser = new DOMParser();

    const wbXml = await zip.file("xl/workbook.xml").async("string");
    const wbDoc = parser.parseFromString(wbXml, "text/xml");
    const sheetsEl = wbDoc.getElementsByTagName("sheets")[0];
    const sheetNodes = wbDoc.getElementsByTagName("sheet");
    let maxSheetId = 0;
    for (let i = 0; i < sheetNodes.length; i++) {
      maxSheetId = Math.max(maxSheetId, parseInt(sheetNodes[i].getAttribute("sheetId") || "0", 10));
    }

    const relsPath = "xl/_rels/workbook.xml.rels";
    const relsXml = await zip.file(relsPath).async("string");
    const relsDoc = parser.parseFromString(relsXml, "text/xml");
    const relNodes = relsDoc.getElementsByTagName("Relationship");
    let maxRid = 0;
    for (let i = 0; i < relNodes.length; i++) {
      const m = /^rId(\d+)$/.exec(relNodes[i].getAttribute("Id") || "");
      if (m) maxRid = Math.max(maxRid, parseInt(m[1], 10));
    }
    const newRid = `rId${maxRid + 1}`;
    const relsRoot = relsDoc.documentElement;
    const newRel = relsDoc.createElementNS(relsRoot.namespaceURI, "Relationship");
    newRel.setAttribute("Id", newRid);
    newRel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
    newRel.setAttribute("Target", `worksheets/sheet${newN}.xml`);
    relsRoot.appendChild(newRel);
    zip.file(relsPath, new XMLSerializer().serializeToString(relsDoc));

    const newSheetEl = wbDoc.createElementNS(sheetsEl.namespaceURI, "sheet");
    newSheetEl.setAttribute("name", backupName);
    newSheetEl.setAttribute("sheetId", String(maxSheetId + 1));
    newSheetEl.setAttributeNS(RELS_NS, "r:id", newRid);
    sheetsEl.appendChild(newSheetEl);
    zip.file("xl/workbook.xml", new XMLSerializer().serializeToString(wbDoc));

    const ctPath = "[Content_Types].xml";
    const ctXml = await zip.file(ctPath).async("string");
    const ctDoc = parser.parseFromString(ctXml, "text/xml");
    const ctRoot = ctDoc.documentElement;
    const override = ctDoc.createElementNS(ctRoot.namespaceURI, "Override");
    override.setAttribute("PartName", `/${newSheetPath}`);
    override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
    ctRoot.appendChild(override);
    zip.file(ctPath, new XMLSerializer().serializeToString(ctDoc));

    return backupName;
  }

  async function reencodeImage(file, mime) {
    const img = await loadImage(file);
    const c = document.createElement("canvas");
    c.width = img.naturalWidth || 1;
    c.height = img.naturalHeight || 1;
    c.getContext("2d").drawImage(img, 0, 0);
    return new Promise((resolve) => c.toBlob((b) => resolve(b), mime, 0.92));
  }

  async function applyEditsAndExport() {
    if (!state.currentZip || !state.currentSheetXmlPath || !state.currentParams) {
      return toast("Rode uma pesquisa primeiro.", "error");
    }
    const pending = state.pendingEdits;
    const editCount = Object.keys(pending).length;
    if (editCount === 0) return toast("Nenhuma edição pendente.", "warn");

    const ok = confirm(
      `Confirmar exportação com ${editCount} linha(s) editada(s)?\n\nUm backup automático dos dados originais será criado numa aba separada dentro do próprio arquivo.`
    );
    if (!ok) return;

    try {
      log("header", "=".repeat(60));
      log("header", "EXPORTANDO PLANILHA EDITADA");
      log("header", "=".repeat(60));

      const zip = state.currentZip;
      const sheetXmlPath = state.currentSheetXmlPath;
      const sheetName = state.currentParams.sheet_name;
      const sheet = state.currentWorkbook.Sheets[sheetName];

      const originalXmlString = await zip.file(sheetXmlPath).async("string");
      const backupName = await addBackupSheet(zip, sheetXmlPath, sheetName, originalXmlString);
      log("success", `Backup automático criado na aba "${backupName}" (dados originais preservados).`);

      const parser = new DOMParser();
      const doc = parser.parseFromString(originalXmlString, "text/xml");
      const editsList = [];
      Object.entries(pending).forEach(([celula, edit]) => {
        Object.entries(edit.cellEdits || {}).forEach(([col, info]) => {
          const rowNum = parseInt(/\d+$/.exec(celula)[0], 10);
          const rawRef = `${col}${rowNum}`;
          const masterRef = resolveMasterCellRef(sheet, rawRef);
          editsList.push({ ref: masterRef, value: info.value, isDate: info.isDate, dateObj: info.dateObj });
        });
      });
      applyCellEditsToXmlDoc(doc, editsList);
      const newXmlString = new XMLSerializer().serializeToString(doc);
      zip.file(sheetXmlPath, newXmlString);
      log("success", `${editsList.length} célula(s) de texto/data atualizada(s) em "${sheetName}".`);

      let imgCount = 0;
      for (const [celula, edit] of Object.entries(pending)) {
        if (!edit.imageFile || !edit.imageMediaPath) continue;
        const ext = (edit.imageMediaPath.split(".").pop() || "png").toLowerCase();
        const mime = ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "gif" ? "image/gif" : "image/png";
        const blob = await reencodeImage(edit.imageFile, mime);
        zip.file(edit.imageMediaPath, blob);
        imgCount++;
        log("success", `Imagem da célula ${celula} substituída (${edit.imageFile.name}).`);
      }
      if (imgCount === 0) log("dim", "Nenhuma imagem foi substituída nesta exportação.");

      log("info", "Gerando arquivo .xlsx final...");
      await yieldTick();
      const outBlob = await zip.generateAsync({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        compression: "DEFLATE",
      });

      const baseName = (state.xlsxFile?.name || "planilha.xlsx").replace(/\.xlsx?$/i, "");
      const stamp = new Date().toISOString().replace(/[:T]/g, "-").slice(0, 19);
      const outName = `${baseName}-editado-${stamp}.xlsx`;
      const a = document.createElement("a");
      a.href = URL.createObjectURL(outBlob);
      a.download = outName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(a.href), 8000);

      log("success", `Arquivo exportado: ${outName}`);
      log("header", "=".repeat(60));
      toast(`Excel exportado com ${editCount} linha(s) editada(s) · backup em "${backupName}"`, "success");

      state.pendingEdits = {};
      updateEditControls();
      refreshMatchCardsAfterEdit();
      renderSheetTable();
    } catch (err) {
      console.error(err);
      log("error", `Erro ao exportar: ${err.message}`);
      toast("Falha ao exportar. Veja detalhes no console (terminal).", "error");
    }
  }

  on($("btn-save-xlsx"), "click", applyEditsAndExport);
  on($("btn-save-xlsx-2"), "click", applyEditsAndExport);

  // Init
  updateRunButton();
  updateEditControls();
  renderSheetTable();
  if (typeof window.XLSX === "undefined" || typeof window.JSZip === "undefined") {
    toast("Falha ao carregar bibliotecas — verifique sua conexão", "error");
  }
})();
