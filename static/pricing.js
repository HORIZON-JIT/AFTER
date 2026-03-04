/**
 * 価格計算システム フロントエンド
 */
(function () {
  "use strict";

  // DOM
  const inputEl = document.getElementById("hinban-input");
  const inputCount = document.getElementById("input-count");
  const btnCalculate = document.getElementById("btn-calculate");
  const btnClear = document.getElementById("btn-clear");
  const btnSample = document.getElementById("btn-sample");
  const btnExport = document.getElementById("btn-export-csv");
  const calcAdditional = document.getElementById("calc-additional");
  const loadingEl = document.getElementById("loading");
  const summaryEl = document.getElementById("summary-section");
  const filterEl = document.getElementById("filter-section");
  const resultEl = document.getElementById("result-section");
  const resultBody = document.getElementById("result-body");
  const filterInput = document.getElementById("filter-input");
  const filterCategory = document.getElementById("filter-category");
  const filterNull = document.getElementById("filter-null");
  const rateGrid = document.getElementById("rate-grid");

  // State
  let allResults = [];
  let sortCol = -1;
  let sortAsc = true;

  // 掛率ラベル
  const RATE_LABELS = {
    up: "UP", charge: "チャージ",
    m_rate_1: "M掛率1", m_band_1: "M価格帯1",
    m_rate_2: "M掛率2", m_band_2: "M価格帯2",
    m_rate_3: "M掛率3", m_band_3: "M価格帯3",
    rate_4: "4番掛率", rate_e: "E番掛率",
    rate_a: "A番掛率", rate_l: "L番掛率",
    rate_cv: "CV番掛率", rate_um: "UM番掛率", rate_p: "P番掛率",
    hi_var1: "HI変数1", hi_var2: "HI変数2", hi_var3: "HI変数3",
    ka_var1: "仮上代1", ka_var2: "仮上代2", ka_var3: "仮上代3", ka_var4: "仮上代4",
    de_var1: "D仕切変数", jo_rate1: "上代率1", jo_band1: "上代帯1",
    jo_rate2: "上代率2", jo_band2: "上代帯2", jo_rate3: "上代率3",
    h_ratio: "H比較率",
  };

  // Init
  loadDefaultRates();
  updateInputCount();

  // Events
  inputEl.addEventListener("input", updateInputCount);
  btnCalculate.addEventListener("click", runCalculation);
  btnClear.addEventListener("click", () => {
    inputEl.value = "";
    updateInputCount();
    hideResults();
  });
  btnSample.addEventListener("click", () => {
    inputEl.value =
      "M167001-18\n4012273-00\nE505712-06\nA910246-01\nL100001-00\nCV200001-00\nP300001-00";
    updateInputCount();
  });
  btnExport.addEventListener("click", exportCSV);
  filterInput.addEventListener("input", applyFilters);
  filterCategory.addEventListener("change", applyFilters);
  filterNull.addEventListener("change", applyFilters);

  // Sort
  document.querySelectorAll("th.sortable").forEach((th) => {
    th.addEventListener("click", () => {
      const col = parseInt(th.dataset.col);
      if (sortCol === col) {
        sortAsc = !sortAsc;
      } else {
        sortCol = col;
        sortAsc = true;
      }
      renderResults(getFilteredResults());
    });
  });

  // Modal
  document.getElementById("modal-close").addEventListener("click", closeModal);
  document.getElementById("modal-overlay").addEventListener("click", (e) => {
    if (e.target === e.currentTarget) closeModal();
  });

  function updateInputCount() {
    const lines = parseHinbanList();
    inputCount.textContent = lines.length + " 件";
  }

  function parseHinbanList() {
    const raw = inputEl.value.trim();
    if (!raw) return [];
    return raw
      .split(/[\n,;]+/)
      .map((s) => s.trim())
      .filter(Boolean);
  }

  async function loadDefaultRates() {
    try {
      const res = await fetch("/api/pricing/rates");
      const data = await res.json();
      renderRateGrid(data);
    } catch {
      renderRateGrid({});
    }
  }

  function renderRateGrid(defaults) {
    rateGrid.innerHTML = "";
    for (const [key, label] of Object.entries(RATE_LABELS)) {
      const val = defaults[key] ?? "";
      const div = document.createElement("div");
      div.innerHTML =
        `<label>${label}</label>` +
        `<input type="number" step="any" data-rate="${key}" value="${val}">`;
      rateGrid.appendChild(div);
    }
  }

  function collectRates() {
    const rates = {};
    rateGrid.querySelectorAll("input[data-rate]").forEach((inp) => {
      const v = parseFloat(inp.value);
      if (!isNaN(v)) rates[inp.dataset.rate] = v;
    });
    return rates;
  }

  async function runCalculation() {
    const list = parseHinbanList();
    if (list.length === 0) {
      alert("品番を入力してください");
      return;
    }

    showLoading();
    hideResults();

    const body = {
      hinban_list: list,
      calc_additional: calcAdditional.checked,
      rates: collectRates(),
    };

    try {
      const res = await fetch("/api/pricing/calculate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      const data = await res.json();
      if (data.error) {
        alert("エラー: " + data.error);
        hideLoading();
        return;
      }

      allResults = data.results || [];
      showSummary(data);
      renderResults(allResults);
      showResults();
    } catch (e) {
      alert("通信エラー: " + e.message);
    } finally {
      hideLoading();
    }
  }

  function showSummary(data) {
    document.getElementById("sum-count").textContent = data.count;
    document.getElementById("sum-elapsed").textContent = data.elapsed_sec + "s";
    document.getElementById("sum-genka").textContent = fmt(data.total_genka);
    document.getElementById("sum-t").textContent = fmt(data.total_t_sikiri);
    document.getElementById("sum-null").textContent = data.null_count;
    summaryEl.classList.remove("hidden");
  }

  function renderResults(results) {
    // Sort
    if (sortCol >= 0) {
      const keys = [
        null, "hinban", null, "genka", "t_sikiri", "h_sikiri",
        "hi_sikiri", "kari_joudai", "dealer_sikiri", "joudai",
      ];
      const key = keys[sortCol];
      if (key) {
        results = [...results].sort((a, b) => {
          const va = typeof a[key] === "number" ? a[key] : String(a[key] || "");
          const vb = typeof b[key] === "number" ? b[key] : String(b[key] || "");
          if (va < vb) return sortAsc ? -1 : 1;
          if (va > vb) return sortAsc ? 1 : -1;
          return 0;
        });
      } else if (sortCol === 0) {
        // index sort (default order or reverse)
        if (!sortAsc) results = [...results].reverse();
      }
    }

    resultBody.innerHTML = "";
    results.forEach((r, i) => {
      const tr = document.createElement("tr");

      // H仕切り比較ハイライト
      let tClass = "";
      if (r.h_sikiri > 0 && r.t_sikiri > 0) {
        const hRatio = collectRates().h_ratio || 1;
        if (r.t_sikiri >= r.h_sikiri * hRatio) {
          tClass = "price-high";
        } else if (r.t_sikiri < r.h_sikiri) {
          tClass = "price-low";
        }
      }

      const hinbanCell = r.category === "A"
        ? `<td><span class="clickable" data-idx="${i}">${esc(r.hinban)}</span></td>`
        : `<td>${esc(r.hinban)}</td>`;

      tr.innerHTML =
        `<td>${i + 1}</td>` +
        hinbanCell +
        `<td><span class="cat-badge cat-${esc(r.category)}">${esc(r.category)}</span></td>` +
        `<td class="num">${fmt2(r.genka)}</td>` +
        `<td class="num ${tClass}">${fmt(r.t_sikiri)}</td>` +
        `<td class="num">${fmt(r.h_sikiri)}</td>` +
        `<td class="num">${fmt(r.hi_sikiri)}</td>` +
        `<td class="num">${fmt(r.kari_joudai)}</td>` +
        `<td class="num">${fmt(r.dealer_sikiri)}</td>` +
        `<td class="num">${fmt(r.joudai)}</td>` +
        `<td>${r.rate_used || ""}</td>` +
        `<td>${esc(r.buhin_kubun || "")}</td>` +
        `<td>${r.null_check ? '<span class="null-flag">有</span>' : ""}</td>`;

      resultBody.appendChild(tr);
    });

    // A番クリックイベント
    resultBody.querySelectorAll(".clickable").forEach((el) => {
      el.addEventListener("click", () => {
        const idx = parseInt(el.dataset.idx);
        showADetail(results[idx]);
      });
    });
  }

  function getFilteredResults() {
    let filtered = allResults;
    const txt = filterInput.value.trim().toUpperCase();
    const cat = filterCategory.value;
    const nullOnly = filterNull.checked;

    if (txt) {
      filtered = filtered.filter((r) => r.hinban.toUpperCase().includes(txt));
    }
    if (cat) {
      filtered = filtered.filter((r) => r.category === cat);
    }
    if (nullOnly) {
      filtered = filtered.filter((r) => r.null_check === "有");
    }
    return filtered;
  }

  function applyFilters() {
    renderResults(getFilteredResults());
  }

  // A番詳細モーダル
  function showADetail(r) {
    document.getElementById("modal-title").textContent =
      `A番 詳細: ${r.hinban}`;
    const body = document.getElementById("modal-body");

    let html = `<div class="modal-summary">
      <div class="item"><div class="item-label">BOM部品原価合計</div><div class="item-value">${fmt2(r.bom_total)}</div></div>
      <div class="item"><div class="item-label">組立工数コスト</div><div class="item-value">${fmt2(r.assembly_cost)}</div></div>
      <div class="item"><div class="item-label">社外組立費</div><div class="item-value">${fmt(r.outsource_cost)}</div></div>
      <div class="item"><div class="item-label">組立場所</div><div class="item-value">${esc(r.assembly_place || "-")}</div></div>
      <div class="item"><div class="item-label">原価合計</div><div class="item-value">${fmt2(r.genka)}</div></div>
      <div class="item"><div class="item-label">T仕切り</div><div class="item-value">${fmt(r.t_sikiri)}</div></div>
    </div>`;

    if (r.children && r.children.length > 0) {
      html += `<h4>構成部品 (${r.children.length}件)</h4>
      <table>
        <thead><tr><th>#</th><th>品番</th><th>部品名</th><th class="num">員数</th><th class="num">単価</th><th class="num">T仕切り</th></tr></thead>
        <tbody>`;
      r.children.forEach((c, i) => {
        html += `<tr>
          <td>${i + 1}</td>
          <td>${esc(c.hinban)}</td>
          <td>${esc(c.name || "")}</td>
          <td class="num">${c.inzu}</td>
          <td class="num">${fmt2(c.unit_price)}</td>
          <td class="num">${fmt(c.t_sikiri)}</td>
        </tr>`;
      });
      html += `</tbody></table>`;
    }

    body.innerHTML = html;
    document.getElementById("modal-overlay").classList.remove("hidden");
  }

  function closeModal() {
    document.getElementById("modal-overlay").classList.add("hidden");
  }

  // CSV出力
  function exportCSV() {
    if (allResults.length === 0) {
      alert("先に計算を実行してください");
      return;
    }

    const header = [
      "品番", "カテゴリ", "原価", "T仕切り", "H仕切り",
      "HI仕切り", "仮上代", "D仕切り", "上代", "掛率", "部品区分", "NULL",
    ];
    const rows = allResults.map((r) => [
      r.hinban, r.category, r.genka, r.t_sikiri, r.h_sikiri,
      r.hi_sikiri, r.kari_joudai, r.dealer_sikiri, r.joudai,
      r.rate_used, r.buhin_kubun, r.null_check,
    ]);

    const bom = "\uFEFF";
    const csv =
      bom +
      header.join(",") +
      "\n" +
      rows.map((r) => r.map((v) => `"${v}"`).join(",")).join("\n");

    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `price_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  // Helpers
  function showLoading() { loadingEl.classList.remove("hidden"); }
  function hideLoading() { loadingEl.classList.add("hidden"); }
  function hideResults() {
    summaryEl.classList.add("hidden");
    filterEl.classList.add("hidden");
    resultEl.classList.add("hidden");
  }
  function showResults() {
    filterEl.classList.remove("hidden");
    resultEl.classList.remove("hidden");
  }

  function fmt(n) {
    if (n == null || n === 0) return "-";
    return Math.round(n).toLocaleString("ja-JP");
  }
  function fmt2(n) {
    if (n == null || n === 0) return "-";
    return Number(n).toLocaleString("ja-JP", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
  }
  function esc(s) {
    if (!s) return "";
    const d = document.createElement("div");
    d.textContent = s;
    return d.innerHTML;
  }
})();
