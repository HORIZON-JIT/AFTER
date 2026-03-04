/**
 * アフター部門 出荷数ランキングアプリ - フロントエンド
 */
(function () {
    "use strict";

    // --- State ---
    let currentMode = "yearly";
    let rankingData = [];
    let sortCol = "rank_no";
    let sortAsc = true;

    // --- DOM ---
    const $ = (sel) => document.querySelector(sel);
    const $$ = (sel) => document.querySelectorAll(sel);

    const tabs = $$(".tab");
    const pickers = {
        yearly: $("#picker-yearly"),
        monthly: $("#picker-monthly"),
        daily: $("#picker-daily"),
    };
    const inputs = {
        yearly: $("#input-year"),
        monthly: $("#input-month"),
        daily: $("#input-date"),
    };
    const btnSearch = $("#btn-search");
    const loading = $("#loading");
    const tbody = $("#ranking-body");
    const filterInput = $("#filter-input");
    const btnExport = $("#btn-export");
    const summaryPeriod = $("#summary-period");
    const summaryCount = $("#summary-count");
    const summaryTotal = $("#summary-total");

    // --- Init ---
    document.addEventListener("DOMContentLoaded", () => {
        $("#footer-year").textContent = new Date().getFullYear();
        bindEvents();
        // 初回検索を自動実行
        fetchRanking();
    });

    // --- Events ---
    function bindEvents() {
        tabs.forEach((tab) => {
            tab.addEventListener("click", () => switchMode(tab.dataset.mode));
        });

        btnSearch.addEventListener("click", fetchRanking);

        // キーボード対応
        Object.values(inputs).forEach((input) => {
            input.addEventListener("keydown", (e) => {
                if (e.key === "Enter") fetchRanking();
            });
        });

        // 前後ボタン
        $("#year-prev").addEventListener("click", () => stepYear(-1));
        $("#year-next").addEventListener("click", () => stepYear(1));
        $("#month-prev").addEventListener("click", () => stepMonth(-1));
        $("#month-next").addEventListener("click", () => stepMonth(1));
        $("#day-prev").addEventListener("click", () => stepDay(-1));
        $("#day-next").addEventListener("click", () => stepDay(1));

        // フィルタ
        filterInput.addEventListener("input", renderTable);

        // CSV出力
        btnExport.addEventListener("click", exportCSV);

        // ソート
        $$(".sortable").forEach((th) => {
            th.addEventListener("click", () => handleSort(th.dataset.col));
        });
    }

    // --- Mode switching ---
    function switchMode(mode) {
        currentMode = mode;
        tabs.forEach((t) => t.classList.toggle("active", t.dataset.mode === mode));
        Object.keys(pickers).forEach((k) => {
            pickers[k].style.display = k === mode ? "flex" : "none";
        });
        fetchRanking();
    }

    // --- Date navigation ---
    function stepYear(delta) {
        inputs.yearly.value = parseInt(inputs.yearly.value) + delta;
        fetchRanking();
    }

    function stepMonth(delta) {
        const val = inputs.monthly.value;
        const d = new Date(val + "-01");
        d.setMonth(d.getMonth() + delta);
        inputs.monthly.value =
            d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
        fetchRanking();
    }

    function stepDay(delta) {
        const val = inputs.daily.value;
        const d = new Date(val);
        d.setDate(d.getDate() + delta);
        inputs.daily.value = formatDateISO(d);
        fetchRanking();
    }

    function formatDateISO(d) {
        return (
            d.getFullYear() +
            "-" +
            String(d.getMonth() + 1).padStart(2, "0") +
            "-" +
            String(d.getDate()).padStart(2, "0")
        );
    }

    // --- Fetch ---
    function getTarget() {
        switch (currentMode) {
            case "yearly":
                return inputs.yearly.value;
            case "monthly":
                return inputs.monthly.value;
            case "daily":
                return inputs.daily.value;
        }
    }

    async function fetchRanking() {
        const target = getTarget();
        if (!target) return;

        loading.style.display = "flex";
        tbody.innerHTML = "";

        try {
            const url = `/api/ranking?mode=${currentMode}&target=${encodeURIComponent(target)}`;
            const res = await fetch(url);
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            const data = await res.json();

            rankingData = data.ranking || [];
            sortCol = "rank_no";
            sortAsc = true;
            updateSortHeaders();

            // Summary
            summaryPeriod.textContent = data.period_label || "-";
            summaryCount.textContent = data.total_count
                ? data.total_count.toLocaleString() + " 品番"
                : "-";
            const totalQty = rankingData.reduce(
                (sum, r) => sum + (r.total_shukka_suu || 0),
                0
            );
            summaryTotal.textContent = totalQty
                ? totalQty.toLocaleString() + " 個"
                : "-";

            renderTable();
        } catch (err) {
            console.error("Fetch error:", err);
            tbody.innerHTML =
                '<tr><td colspan="5" class="empty-message">データの取得に失敗しました</td></tr>';
        } finally {
            loading.style.display = "none";
        }
    }

    // --- Sort ---
    function handleSort(col) {
        if (sortCol === col) {
            sortAsc = !sortAsc;
        } else {
            sortCol = col;
            sortAsc = col === "hinban" || col === "hm_nm"; // テキストはA→Z、数値は大→小
        }
        updateSortHeaders();
        renderTable();
    }

    function updateSortHeaders() {
        $$(".sortable").forEach((th) => {
            const col = th.dataset.col;
            let label = {
                rank_no: "順位",
                hinban: "品番",
                hm_nm: "品名",
                total_shukka_suu: "出荷数合計",
                shukka_count: "出荷回数",
            }[col];

            if (col === sortCol) {
                label += sortAsc ? " \u25B2" : " \u25BC";
            }
            th.textContent = label;
        });
    }

    // --- Render ---
    function renderTable() {
        const filter = filterInput.value.trim().toLowerCase();

        let filtered = rankingData;
        if (filter) {
            filtered = rankingData.filter(
                (r) =>
                    (r.hinban || "").toLowerCase().includes(filter) ||
                    (r.hm_nm || "").toLowerCase().includes(filter)
            );
        }

        // Sort
        const sorted = [...filtered].sort((a, b) => {
            let va = a[sortCol];
            let vb = b[sortCol];
            if (typeof va === "string") va = va.toLowerCase();
            if (typeof vb === "string") vb = vb.toLowerCase();
            if (va < vb) return sortAsc ? -1 : 1;
            if (va > vb) return sortAsc ? 1 : -1;
            return 0;
        });

        // Max value for bar
        const maxQty = Math.max(
            ...sorted.map((r) => r.total_shukka_suu || 0),
            1
        );

        if (sorted.length === 0) {
            tbody.innerHTML =
                '<tr><td colspan="5" class="empty-message">該当データがありません</td></tr>';
            return;
        }

        const fragment = document.createDocumentFragment();
        sorted.forEach((row, idx) => {
            const tr = document.createElement("tr");

            // 順位
            const tdRank = document.createElement("td");
            const rankNum = row.rank_no;
            if (rankNum <= 3) {
                const medal = document.createElement("span");
                medal.className = `rank-medal rank-${rankNum}`;
                medal.textContent = rankNum;
                tdRank.appendChild(medal);
            } else {
                tdRank.textContent = rankNum;
            }
            tr.appendChild(tdRank);

            // 品番
            const tdHinban = document.createElement("td");
            tdHinban.innerHTML = highlightText(row.hinban || "", filter);
            tr.appendChild(tdHinban);

            // 品名
            const tdName = document.createElement("td");
            tdName.innerHTML = highlightText(row.hm_nm || "", filter);
            tr.appendChild(tdName);

            // 出荷数（バー付き）
            const tdQty = document.createElement("td");
            tdQty.className = "bar-cell";
            const pct = ((row.total_shukka_suu || 0) / maxQty) * 100;
            tdQty.innerHTML =
                `<span class="bar-bg" style="width:${pct}%"></span>` +
                `<span class="bar-value">${(row.total_shukka_suu || 0).toLocaleString()}</span>`;
            tr.appendChild(tdQty);

            // 出荷回数
            const tdCount = document.createElement("td");
            tdCount.textContent = (row.shukka_count || 0).toLocaleString();
            tr.appendChild(tdCount);

            fragment.appendChild(tr);
        });

        tbody.innerHTML = "";
        tbody.appendChild(fragment);
    }

    function highlightText(text, filter) {
        if (!filter) return escapeHTML(text);
        const escaped = escapeHTML(text);
        const regex = new RegExp(
            `(${escapeRegex(filter)})`,
            "gi"
        );
        return escaped.replace(regex, "<mark>$1</mark>");
    }

    function escapeHTML(str) {
        const div = document.createElement("div");
        div.textContent = str;
        return div.innerHTML;
    }

    function escapeRegex(str) {
        return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }

    // --- CSV Export ---
    function exportCSV() {
        if (rankingData.length === 0) {
            alert("データがありません");
            return;
        }

        const filter = filterInput.value.trim().toLowerCase();
        let data = rankingData;
        if (filter) {
            data = rankingData.filter(
                (r) =>
                    (r.hinban || "").toLowerCase().includes(filter) ||
                    (r.hm_nm || "").toLowerCase().includes(filter)
            );
        }

        const bom = "\uFEFF";
        let csv = bom + "順位,品番,品名,出荷数合計,出荷回数\r\n";
        data.forEach((r) => {
            csv +=
                [
                    r.rank_no,
                    `"${(r.hinban || "").replace(/"/g, '""')}"`,
                    `"${(r.hm_nm || "").replace(/"/g, '""')}"`,
                    r.total_shukka_suu || 0,
                    r.shukka_count || 0,
                ].join(",") + "\r\n";
        });

        const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        const period = summaryPeriod.textContent || "ranking";
        a.download = `出荷ランキング_${period}.csv`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }
})();
