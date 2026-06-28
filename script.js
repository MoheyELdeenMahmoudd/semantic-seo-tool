/**
 * Semantic SEO Tool — Frontend Logic
 * ====================================
 * - Professional loading states with step progress
 * - Typed error handling with clear user messages
 * - Pydantic-aligned data rendering
 * - Excel download via base64 blob
 */

"use strict";

// ─────────────────────────────────────────────────────────────────────────────
// State
// ─────────────────────────────────────────────────────────────────────────────

let _excelB64  = "";
let _briefText = "";

// ─────────────────────────────────────────────────────────────────────────────
// File input UI feedback
// ─────────────────────────────────────────────────────────────────────────────

document.getElementById("myFile").addEventListener("change", function () {
  const statusEl = document.getElementById("myFileStatus");
  const boxEl    = document.getElementById("myBox");
  if (this.files.length > 0) {
    statusEl.textContent = "✅ " + this.files[0].name;
    statusEl.style.color = "#059669";
    boxEl.classList.add("has-file");
  }
});

document.getElementById("compFiles").addEventListener("change", function () {
  const statusEl = document.getElementById("compFileStatus");
  const boxEl    = document.getElementById("compBox");
  if (this.files.length > 0) {
    statusEl.textContent = `✅ تم تحديد ${this.files.length} ملف`;
    statusEl.style.color = "#F43F5E";
    boxEl.classList.add("has-file");
  }
});

document.getElementById("myUrl").addEventListener("input", function () {
  if (this.value.trim()) {
    document.getElementById("myFileStatus").textContent = "🔗 سيتم تحليل الرابط";
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// Loading progress
// ─────────────────────────────────────────────────────────────────────────────

const STEPS = [
  { pct: 8,  msg: "استخراج النصوص من الملفات..." },
  { pct: 20, msg: "إرسال المحتوى لـ Nara AI للتحليل العميق..." },
  { pct: 40, msg: "بناء العناقيد الدلالية (Topic Clusters)..." },
  { pct: 60, msg: "تحليل ملفات المنافسين..." },
  { pct: 75, msg: "حساب الفجوات الدلالية وترتيب الأولويات..." },
  { pct: 88, msg: "توليد خطة المحتوى (Content Brief)..." },
  { pct: 95, msg: "تجهيز التقرير النهائي..." },
];

let _progressTimer  = null;
let _currentStep    = 0;

function startProgress() {
  _currentStep = 0;
  setProgress(STEPS[0].pct, STEPS[0].msg);
  _progressTimer = setInterval(() => {
    _currentStep = Math.min(_currentStep + 1, STEPS.length - 1);
    const step = STEPS[_currentStep];
    setProgress(step.pct, step.msg);
  }, 3500);
}

function stopProgress() {
  clearInterval(_progressTimer);
}

function setProgress(pct, msg) {
  document.getElementById("progressBar").style.width = pct + "%";
  if (msg) document.getElementById("loadingStep").textContent = msg;
}

function setLoadingTitle(text) {
  document.getElementById("loadingTitle").textContent = text;
}

// ─────────────────────────────────────────────────────────────────────────────
// Error handling
// ─────────────────────────────────────────────────────────────────────────────

function showError(msg) {
  const banner = document.getElementById("errorBanner");
  document.getElementById("errorText").textContent = msg;
  banner.classList.add("visible");
  banner.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function hideError() {
  document.getElementById("errorBanner").classList.remove("visible");
}

function uiError(msg) {
  // Show error, restore button, hide loading
  showError(msg);
  setLoading(false);
}

// ─────────────────────────────────────────────────────────────────────────────
// UI state toggle
// ─────────────────────────────────────────────────────────────────────────────

function setLoading(on) {
  const btn     = document.getElementById("analyzeBtn");
  const loading = document.getElementById("loadingCard");
  btn.disabled  = on;
  loading.style.display = on ? "block" : "none";
  if (!on) { stopProgress(); setProgress(0, ""); }
}

// ─────────────────────────────────────────────────────────────────────────────
// Main analysis runner
// ─────────────────────────────────────────────────────────────────────────────

async function runAnalysis() {
  hideError();

  const myFile     = document.getElementById("myFile").files[0];
  const myUrl      = document.getElementById("myUrl").value.trim();
  const compFiles  = document.getElementById("compFiles").files;
  const compUrl1   = document.getElementById("compUrl1").value.trim();
  const userApiKey = document.getElementById("userApiKey").value.trim(); // السطر الجديد

  // التأكد من إدخال الـ API Key
  if (!userApiKey) {
    showError("يرجى إدخال مفتاح Nara API الخاص بك للبدء.");
    return;
  }

  if (!myFile && !myUrl && compFiles.length === 0 && !compUrl1) {
    showError("يرجى رفع ملف أو إدخال رابط على الأقل قبل بدء التحليل.");
    return;
  }

  setLoading(true);
  document.getElementById("resultsArea").style.display = "none";
  _excelB64  = "";
  _briefText = "";

  startProgress();
  setLoadingTitle("جارٍ تحليل المحتوى باستخدام Nara AI...");

  const formData = new FormData();
  formData.append("api_key", userApiKey); // إرسال المفتاح في الـ FormData
  if (myFile)   formData.append("my_article", myFile);
  if (myUrl)    formData.append("my_url",     myUrl);
  for (const f of compFiles) formData.append("competitors", f);
  if (compUrl1) formData.append("comp_urls",  compUrl1);

  try {
    const resp = await fetch("/analyze", {
      method: "POST",
      body:   formData,
    });

    const contentType = resp.headers.get("content-type") || "";
    if (!contentType.includes("application/json")) {
      throw new Error(`الخادم أعاد استجابة غير متوقعة (HTTP ${resp.status}).`);
    }

    const data = await resp.json();

    if (!resp.ok) {
      throw new Error(data.error || `خطأ من الخادم (HTTP ${resp.status}).`);
    }

    stopProgress();
    setProgress(100, "اكتمل التحليل بنجاح ✓");
    setLoadingTitle("تم التحليل بنجاح!");

    setTimeout(() => {
      setLoading(false);
      renderResults(data);
    }, 600);

  } catch (err) {
    console.error("Analysis error:", err);

    let msg = err.message || "حدث خطأ غير متوقع.";

    // Network error (offline / server down)
    if (err instanceof TypeError && err.message.includes("fetch")) {
      msg = "تعذّر الاتصال بالخادم. تأكد أن سيرفر Flask يعمل على المنفذ 5000.";
    }

    uiError(msg);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Badge helpers
// ─────────────────────────────────────────────────────────────────────────────

function typeBadge(entityType) {
  const cls = "badge badge-" + (entityType || "concept").toLowerCase();
  return `<span class="${cls}">${entityType || "Concept"}</span>`;
}

function importanceBadge(imp) {
  const labels = { high: "عالية", medium: "متوسطة", low: "منخفضة" };
  const cls    = "badge badge-" + (imp || "medium");
  return `<span class="${cls}">${labels[imp] || "متوسطة"}</span>`;
}

function priorityBadge(prio) {
  const labels = { critical: "حرجة", high: "عالية", medium: "متوسطة" };
  const cls    = prio === "critical" ? "badge badge-critical"
               : prio === "high"     ? "badge badge-high"
               :                       "badge badge-medium";
  return `<span class="${cls}">${labels[prio] || prio}</span>`;
}

function salienceBar(score) {
  const pct = Math.round((score || 0) * 100);
  return `<span style="font-size:.78rem;font-weight:700;color:#1D4ED8">${pct}%</span>`;
}

// ─────────────────────────────────────────────────────────────────────────────
// Render stats row
// ─────────────────────────────────────────────────────────────────────────────

function renderStats(data) {
  const master = data.master_entities || [];
  const gaps   = data.semantic_gaps   || [];
  const comps  = data.competitor_breakdowns || [];
  const mySite = data.my_site_entities || [];

  const criticalGaps = gaps.filter(g => g.priority === "critical").length;

  const statsEl = document.getElementById("statsRow");
  statsEl.innerHTML = `
    <div class="stat-card">
      <div class="stat-val">${master.length}</div>
      <div class="stat-label">كيان مكتشف</div>
    </div>
    <div class="stat-card">
      <div class="stat-val">${mySite.length}</div>
      <div class="stat-label">كيانات موقعك</div>
    </div>
    <div class="stat-card">
      <div class="stat-val" style="color:var(--rose)">${gaps.length}</div>
      <div class="stat-label">فجوة دلالية</div>
    </div>
    <div class="stat-card">
      <div class="stat-val" style="color:#B91C1C">${criticalGaps}</div>
      <div class="stat-label">فجوة حرجة</div>
    </div>
    <div class="stat-card">
      <div class="stat-val">${comps.length}</div>
      <div class="stat-label">منافس حُلِّل</div>
    </div>
  `;
}

// ─────────────────────────────────────────────────────────────────────────────
// Render gaps table
// ─────────────────────────────────────────────────────────────────────────────

function renderGaps(gaps) {
  const card = document.getElementById("gapsCard");
  const body = document.getElementById("gapsBody");

  if (!gaps || gaps.length === 0) {
    card.style.display = "none";
    return;
  }
  card.style.display = "block";

  body.innerHTML = gaps.slice(0, 80).map(g => `
    <tr>
      <td style="font-weight:700">${g.name || ""}</td>
      <td>${typeBadge(g.type)}</td>
      <td>${priorityBadge(g.priority)}</td>
      <td><span class="mentions-pill rose-pill">${g.competitor_count || 0}</span></td>
      <td><span class="mentions-pill rose-pill">${g.total_mentions || 0}</span></td>
    </tr>
  `).join("");
}

// ─────────────────────────────────────────────────────────────────────────────
// Render master entities table
// ─────────────────────────────────────────────────────────────────────────────

function renderMaster(master, myNames, compNameCounts, hasMyArticle, hasCompetitors) {
  document.getElementById("thMySite").style.display = hasMyArticle  ? "" : "none";
  document.getElementById("thComp").style.display   = hasCompetitors ? "" : "none";

  const body = document.getElementById("masterBody");
  if (!master || master.length === 0) {
    body.innerHTML = `<tr><td colspan="6" style="text-align:center;padding:2rem;color:#94A3B8;font-weight:600">لا توجد بيانات</td></tr>`;
    return;
  }

  body.innerHTML = master.slice(0, 100).map(ent => {
    const key      = (ent.name || "").toLowerCase();
    const inMySite = hasMyArticle  ? (myNames.has(key)
                       ? `<span class="check">✓</span>`
                       : `<span class="dash">—</span>`)
                     : "";
    const inComp   = hasCompetitors
                       ? `<span class="mentions-pill">${compNameCounts[key] || 0}</span>`
                       : "";

    return `
      <tr>
        <td style="font-weight:700">${ent.name || ""}</td>
        <td>${typeBadge(ent.entity_type || ent.type)}</td>
        <td>${importanceBadge(ent.importance)}</td>
        <td>${salienceBar(ent.salience)}</td>
        ${hasMyArticle  ? `<td style="text-align:center">${inMySite}</td>` : ""}
        ${hasCompetitors ? `<td style="text-align:center">${inComp}</td>`   : ""}
      </tr>
    `;
  }).join("");
}

// ─────────────────────────────────────────────────────────────────────────────
// Render content brief
// ─────────────────────────────────────────────────────────────────────────────

function renderBrief(briefText) {
  const card = document.getElementById("briefCard");
  const body = document.getElementById("briefBody");
  if (!briefText || !briefText.trim()) {
    card.style.display = "none";
    return;
  }
  _briefText        = briefText;
  card.style.display = "block";
  body.innerHTML    = marked.parse(briefText);
}

// ─────────────────────────────────────────────────────────────────────────────
// Main render orchestrator
// ─────────────────────────────────────────────────────────────────────────────

function renderResults(data) {
  // Build lookup sets
  const mySiteEntities  = data.my_site_entities || [];
  const myNames         = new Set(mySiteEntities.map(e => (e.name || "").toLowerCase()));

  const compNameCounts  = {};
  for (const comp of (data.competitor_breakdowns || [])) {
    for (const ent of (comp.raw_entities || [])) {
      const key            = (ent.name || "").toLowerCase();
      compNameCounts[key]  = (compNameCounts[key] || 0) + (ent.mentions || 1);
    }
  }

  // Render sections
  renderStats(data);
  renderGaps(data.semantic_gaps || []);
  renderMaster(
    data.master_entities || [],
    myNames,
    compNameCounts,
    data.has_my_article,
    data.has_competitors,
  );
  renderBrief(data.content_brief || "");

  // Store Excel
  if (data.excel_file) {
    _excelB64 = data.excel_file;
    document.getElementById("exportBtn").disabled = false;
  }

  // Show results
  const resultsEl = document.getElementById("resultsArea");
  resultsEl.style.display = "block";
  resultsEl.scrollIntoView({ behavior: "smooth", block: "start" });
}

// ─────────────────────────────────────────────────────────────────────────────
// Excel download
// ─────────────────────────────────────────────────────────────────────────────

function downloadExcel() {
  if (!_excelB64) return;
  try {
    const binary = atob(_excelB64);
    const bytes  = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    const blob   = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const link   = document.createElement("a");
    link.href     = URL.createObjectURL(blob);
    link.download  = `Semantic_SEO_Report_${Date.now()}.xlsx`;
    link.click();
    URL.revokeObjectURL(link.href);
  } catch (err) {
    showError("تعذّر تحميل ملف Excel. حاول مرة أخرى.");
    console.error("Excel download error:", err);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Copy brief to clipboard
// ─────────────────────────────────────────────────────────────────────────────

async function copyBrief() {
  if (!_briefText) return;
  try {
    await navigator.clipboard.writeText(_briefText);
    const btn = document.querySelector(".btn-copy");
    const orig = btn.innerHTML;
    btn.innerHTML = "<i class='fa-solid fa-check'></i> تم النسخ!";
    setTimeout(() => { btn.innerHTML = orig; }, 2000);
  } catch {
    showError("تعذّر نسخ النص. يرجى النسخ يدوياً.");
  }
}
