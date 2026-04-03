let currentCenter = ""; // 👉 실제 의미: 대리점명
let performanceViewCenter = "관리자";

let dashboardDetailSort = {
  key: "",
  order: "asc",
  type: ""
};

let dashboardDetailRows = [];
let inventoryModelTurnoverRows = [];
let inventoryModelTurnoverSort = {
  key: "monthly_sales",
  order: "desc"
};
let activeAgingThreshold = 0;
let activeAdminAgencyDetail = "";
let adminAgencySortKey = "total_qty";
let adminAgencySortDir = "desc";
let agingStockDetailRows = [];
let agingStockDetailSort = {
  key: "aging_days",
  order: "desc"
};
let performanceTrendChart = null;
let performanceSearchCache = [];
let performanceSortState = { key: "", order: "asc" };

function $(id) {
  return document.getElementById(id);
}

function getPerformanceQueryCenter() {
  if (currentCenter === "관리자") {
    return performanceViewCenter || "관리자";
  }
  return currentCenter;
}

function setDisplay(id, display) {
  const el = $(id);
  if (el) el.style.display = display;
}

let activeDashboardDetailType = "";

// =========================
// 비밀번호
// =========================
const passwords = {
  "광주": "kimchizzang1219",
  "목포": "jindotgae00",
  "순천": "yeosuexpo8787",
  "전북": "saemangeum2233",
  "제주": "halubang9987",
  "관리자": "86731152020"
};

const MASTER_PASSWORD = "1252002"; // ⭐ 반드시 문자열

// =========================
// 설정
// =========================
const LOGIN_EXPIRE_MS = 24 * 60 * 60 * 1000; // 24시간
const API_URL = "https://nexus-inventory-site.onrender.com";

// =========================
// 자동 로그인
// =========================
window.addEventListener("load", () => {
  const saved = localStorage.getItem("loginInfo");
  if (!saved) return;

  try {
    const { center, time } = JSON.parse(saved);
    if (!center || !time) return;

    if (Date.now() - time > LOGIN_EXPIRE_MS) {
      localStorage.removeItem("loginInfo");
      return;
    }

    currentCenter = center;

    setDisplay("loginBox", "none");
    setDisplay("logoutBtn", "inline-block");

    const brandSub = $("brandSub");
    if (brandSub) {
      brandSub.innerText = center === "관리자" ? "관리자 모드" : `${center} 대리점 로그인됨`;
    }

    const hero = document.querySelector(".hero");
    if (hero) hero.style.display = "none";

    // ✅ 로그인 후에는 메뉴 화면으로
    setDisplay("menuBox", "block");
setDisplay("securityNoticeBox", "block");
setDisplay("uploadBox", "none");
setDisplay("searchBox", "none");
setDisplay("inventoryDash", "none");
setDisplay("performanceBox", "none");

    setDisplay("todayLoginInfo", "block");
    loadTodayLoginInfo();
  } catch (e) {
    console.error(e);
    localStorage.removeItem("loginInfo");
  }
});

// =========================
// 로그인 / 로그아웃
// =========================
function login() {
  const center = document.getElementById("centerSelect").value;
  const inputPassword = document.getElementById("password").value.trim();

  if (!center) return alert("센터(대리점)를 선택하세요.");

  const isMaster = inputPassword === MASTER_PASSWORD;
  const isCenterValid = inputPassword === passwords[center];

  if (!isMaster && !isCenterValid) {
    return alert("비밀번호가 틀렸습니다.");
  }

  currentCenter = center;

  localStorage.setItem(
    "loginInfo",
    JSON.stringify({
      center: center,
      time: Date.now()
    })
  );

  document.getElementById("loginBox").style.display = "none";
  document.getElementById("logoutBtn").style.display = "inline-block";
  document.getElementById("brandSub").innerText =
    center === "관리자" ? "관리자 모드" : `${center} 대리점 로그인됨`;

  document.querySelector(".hero").style.display = "none";

  document.getElementById("menuBox").style.display = "block";
document.getElementById("securityNoticeBox").style.display = "block";
document.getElementById("uploadBox").style.display = "none";
document.getElementById("searchBox").style.display = "none";
document.getElementById("inventoryDash").style.display = "none";
document.getElementById("performanceBox").style.display = "none";

  setDisplay("todayLoginInfo", "block");
  saveLoginLog();
  loadTodayLoginInfo();
}

function logout() {
  currentCenter = "";
  localStorage.removeItem("loginInfo");

  const info = document.getElementById("todayLoginInfo");
  if (info) {
    info.style.display = "none";
    info.textContent = "";
  }

  location.reload();
}

// =========================
// 엑셀 업로드 (관리자)
// =========================
async function uploadExcel() {
  if (currentCenter !== "관리자") {
    alert("관리자만 업로드할 수 있습니다.");
    return;
  }

  const fileInput = document.getElementById("excelFile");
  const uploadBtn = document.getElementById("uploadBtn");
  const status = document.getElementById("uploadStatus");

  if (!fileInput.files.length) {
    status.innerText = "엑셀 파일을 선택해주세요.";
    return;
  }

  uploadBtn.disabled = true;
  const formData = new FormData();
  formData.append("file", fileInput.files[0]);
  status.innerText = "업로드 중...";

  try {
    const resp = await fetch(`${API_URL}/upload`, {
      method: "POST",
      body: formData
    });

    const j = await resp.json();

    if (resp.ok && j.ok) {
      status.innerText = `✅ 업로드 완료! 총 ${j.count}건 반영됨`;
      fileInput.value = "";
    } else {
      status.innerText = `❌ 업로드 실패: ${j.message || "오류"}`;
    }
  } catch {
    status.innerText = "❌ 업로드 실패: 네트워크 오류";
  } finally {
    uploadBtn.disabled = false;
  }
}

// =========================
// 재고 검색 (🔥 대리점 기준 강제)
// =========================
async function runSearch() {
  const status = document.getElementById("status");
  const model = document.getElementById("model").value.trim();
  const address = document.getElementById("address").value.trim();
  const owner = document.getElementById("owner").value.trim();
  const nickname = document.getElementById("nickname").value.trim();

  if (!model && !address && !owner && !nickname) {
    alert("검색 조건을 하나 이상 입력하세요.");
    return;
  }

  status.innerText = "조회 중...";
  document.getElementById("result").innerHTML = "";

  try {
    const resp = await fetch(`${API_URL}/search`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        agency: currentCenter,
        model,
        address,
        owner,
        nickname
      })
    });

    const j = await resp.json();

    if (!resp.ok || !j.ok) {
      status.innerText = `조회 실패: ${j.message || "오류"}`;
      return;
    }

    status.innerText = `총 ${j.total}대 있습니다.`;

    const sortKey = document.getElementById("sortKey")?.value || "보유처";
    const sortOrder = document.getElementById("sortOrder")?.value || "asc";

    const keyMap = {
      "보유처": "store_name",
      "모델명": "model_name",
      "색상": "color"
    };

    const field = keyMap[sortKey] || "store_name";

    j.table.sort((a, b) => {
      const A = (a[field] || "").toString();
      const B = (b[field] || "").toString();

      const cmp = A.localeCompare(B, "ko");
      return sortOrder === "desc" ? -cmp : cmp;
    });

    const detail = document.getElementById("detailToggle")?.checked || false;

    const baseCols = ["store_name", "model_name", "color", "serial_no"];
    const detailCols = ["address", "pet_name", "nickname"];

    const cols = detail ? [...baseCols, ...detailCols] : baseCols;

    renderTable(j.table, cols);
  } catch {
    status.innerText = "조회 실패: 네트워크 오류";
  }
}

// =========================
// 테이블 렌더링
// =========================
function renderTable(rows, cols) {
  const wrap = document.createElement("div");
  wrap.className = "tableWrap compactTable";

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const trh = document.createElement("tr");

  const headerMap = {
    store_name: "접점명",
    model_name: "모델명",
    color: "색상",
    serial_no: "일련번호",
    address: "상세주소",
    pet_name: "펫네임",
    nickname: "애칭"
  };

  cols.forEach(c => {
    const th = document.createElement("th");
    th.textContent = headerMap[c] || c;
    trh.appendChild(th);
  });

  thead.appendChild(trh);

  const tbody = document.createElement("tbody");

  rows.forEach(r => {
    const tr = document.createElement("tr");
    cols.forEach(c => {
      const td = document.createElement("td");
      td.textContent = (r[c] ?? "").toString();
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  wrap.appendChild(table);

  const result = document.getElementById("result");
  result.innerHTML = "";
  result.appendChild(wrap);
}

// =========================
// 화면 전환
// =========================
function showOnly(ids) {
  const all = ["menuBox", "securityNoticeBox", "inventoryDash", "adminAgencyStockBox", "agingStockBox", "searchBox", "uploadBox", "performanceBox"];
  all.forEach(id => setDisplay(id, ids.includes(id) ? "block" : "none"));
}

function openInventory() {
  if (currentCenter === "관리자") {
  showOnly(["menuBox", "inventoryDash", "adminAgencyStockBox", "agingStockBox", "uploadBox", "searchBox"]);
} else {
  showOnly(["menuBox", "inventoryDash", "agingStockBox", "searchBox"]);
}

  const detailBox = document.getElementById("inventoryDetailPanels");
  if (detailBox) detailBox.style.display = "none";

  loadInventoryDashboard();
}

function openPerformance() {
  showOnly(["menuBox", "performanceBox"]);

  const filterWrap = document.getElementById("performanceAdminFilterWrap");
  const filterSelect = document.getElementById("performanceAdminCenterFilter");

  if (currentCenter === "관리자") {
    if (filterWrap) filterWrap.style.display = "block";
    if (filterSelect) {
      filterSelect.value = performanceViewCenter || "관리자";
    }
  } else {
    if (filterWrap) filterWrap.style.display = "none";
  }

  loadPerformanceDashboard();
}

function onPerformanceAdminFilterChange() {
  const select = document.getElementById("performanceAdminCenterFilter");
  if (!select) return;

  performanceViewCenter = select.value || "관리자";
  loadPerformanceDashboard();
}

// =========================
// 재고 대시보드 로드
// =========================
async function loadInventoryDashboard() {
  const cards = document.getElementById("invCards");
  if (cards) cards.innerHTML = "불러오는 중...";

  try {
    const sResp = await fetch(`${API_URL}/inventory/summary-extended`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ agency: currentCenter })
    });
    const s = await sResp.json();
    if (!sResp.ok || !s.ok) throw new Error(s.message || "요약 실패");

    renderInvCards(s.summary);

        // 관리자 모드: 센터별 총 재고 현황
    if (currentCenter === "관리자") {
      const agencyResp = await fetch(`${API_URL}/inventory/agency-summary`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({})
      });
      const agencyJson = await agencyResp.json();

      if (!agencyResp.ok || !agencyJson.ok) {
        throw new Error(agencyJson.message || "센터별 총 재고 실패");
      }

      renderAdminAgencyStockCards(agencyJson.rows || []);
    }

        // 2) 모델별 회전일
    const mResp = await fetch(`${API_URL}/inventory/by-model-turnover`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ agency: currentCenter })
    });
    const m = await mResp.json();
    if (!mResp.ok || !m.ok) throw new Error(m.message || "모델별 회전일 실패");

    inventoryModelTurnoverRows = Array.isArray(m.rows) ? m.rows : [];
    inventoryModelTurnoverSort = {
  key: "monthly_sales",
  order: "desc"
};

    renderInventoryModelTurnoverTable();

        // 3) 장기 입고 미소진 재고 요약
    const aResp = await fetch(`${API_URL}/inventory/aging-summary`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ agency: currentCenter })
    });
    const a = await aResp.json();
    if (!aResp.ok || !a.ok) throw new Error(a.message || "장기재고 요약 실패");

    renderAgingStockCards(a.summary || {});
  } catch (e) {
    if (cards) cards.innerHTML = `❌ 대시보드 로드 실패: ${e.message}`;
  }
}

function renderInvCards(summary) {
  const cards = document.getElementById("invCards");
  if (!cards) return;

  const { total_qty, store_cnt, model_cnt, warehouse_qty, store_qty } = summary;

  const pillAgency = document.getElementById("pillAgency");
  if (pillAgency) pillAgency.textContent = currentCenter;

  cards.innerHTML = `
    ${card("오늘 총 재고", total_qty, "대", "total")}
    ${card("판매점 수", store_cnt, "곳")}
    ${card("모델 수", model_cnt, "종")}
    ${card("창고 재고", warehouse_qty, "대", "warehouse")}
    ${card("판매점 재고", store_qty, "대", "store")}
  `;
}

function card(title, value, unit, detailType = "") {
  const clickable = detailType ? "clickable" : "";
  const onclick = detailType ? `onclick="loadDashboardDetail('${detailType}')"` : "";

  return `
    <div class="statCard ${clickable}" ${onclick}>
      <div class="statLabel">${title}</div>
      <div class="statValue">${Number(value || 0).toLocaleString()} ${unit}</div>
    </div>
  `;
}

function renderSimpleTable(targetId, rows, cols) {
  const wrap = document.getElementById(targetId);
  if (!wrap) return;

  const table = document.createElement("table");
  table.style.width = "100%";
  table.style.borderCollapse = "collapse";

  const thead = document.createElement("thead");
  const trh = document.createElement("tr");
  cols.forEach(c => {
    const th = document.createElement("th");
    th.textContent = c.label;
    th.style.textAlign = "left";
    th.style.borderBottom = "1px solid #ddd";
    th.style.padding = "8px";
    trh.appendChild(th);
  });
  thead.appendChild(trh);

  const tbody = document.createElement("tbody");
  rows.forEach(r => {
    const tr = document.createElement("tr");
    cols.forEach(c => {
      const td = document.createElement("td");
      const val = r[c.key];
      td.textContent =
        typeof val === "number" ? val.toLocaleString() : (val ?? "").toString();
      td.style.borderBottom = "1px solid #f0f0f0";
      td.style.padding = "8px";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);

  wrap.innerHTML = "";
  wrap.appendChild(table);
}

function renderInventoryModelTurnoverTable() {
  const wrap = document.getElementById("byModelTurnoverTable");
  if (!wrap) return;

  let rows = [...inventoryModelTurnoverRows];

  const key = inventoryModelTurnoverSort.key;
  const order = inventoryModelTurnoverSort.order;

  rows.sort((a, b) => {
    let av = a[key];
    let bv = b[key];

    const numericKeys = ["qty", "monthly_sales", "turnover_days"];
    const isNumeric = numericKeys.includes(key);

    if (isNumeric) {
      av = av === null || av === undefined ? -1 : Number(av);
      bv = bv === null || bv === undefined ? -1 : Number(bv);
      return order === "asc" ? av - bv : bv - av;
    } else {
      const cmp = String(av || "").localeCompare(String(bv || ""), "ko");
      return order === "asc" ? cmp : -cmp;
    }
  });

  let html = `
  <div class="tableWrap inventoryTurnoverWrap">
    <table class="inventoryTurnoverTable">
        <thead>
          <tr>
            <th class="sortable" onclick="toggleInventoryModelTurnoverSort('model_name')">
              모델명${getInventoryModelTurnoverSortIndicator('model_name')}
            </th>
            <th class="sortable" onclick="toggleInventoryModelTurnoverSort('qty')">
              수량${getInventoryModelTurnoverSortIndicator('qty')}
            </th>
            <th class="sortable" onclick="toggleInventoryModelTurnoverSort('monthly_sales')">
              당월판매${getInventoryModelTurnoverSortIndicator('monthly_sales')}
            </th>
            <th class="sortable" onclick="toggleInventoryModelTurnoverSort('turnover_days')">
              회전일${getInventoryModelTurnoverSortIndicator('turnover_days')}
            </th>
          </tr>
        </thead>
        <tbody>
  `;

  rows.forEach(r => {
  const turnover = r.turnover_days === null || r.turnover_days === undefined
    ? null
    : Number(r.turnover_days);

  const isSlow = turnover !== null && turnover >= 100;   // 부진재고
  const isShort = turnover !== null && turnover < 40;    // 부족재고

  let turnoverText = "-";
  if (turnover !== null) {
    if (isSlow) {
      turnoverText = `⚠️ ${turnover.toLocaleString()}일`;
    } else if (isShort) {
      turnoverText = `부족 ${turnover.toLocaleString()}일`;
    } else {
      turnoverText = `${turnover.toLocaleString()}일`;
    }
  }

  html += `
    <tr class="${isSlow ? "slowStockRow" : isShort ? "shortStockRow" : ""}">
      <td>
  ${
    currentCenter === "관리자"
      ? `<button type="button" class="linkLikeBtn" onclick='openModelTurnoverModal(${JSON.stringify(r.model_name || "")})'>${escapeHtml(r.model_name || "")}</button>`
      : `${escapeHtml(r.model_name || "")}`
  }
</td>
      <td>${Number(r.qty || 0).toLocaleString()}</td>
      <td>${Number(r.monthly_sales || 0).toLocaleString()}</td>
      <td class="${isSlow ? "slowStockText" : isShort ? "shortStockText" : ""}">
        ${turnoverText}
      </td>
    </tr>
  `;
});

  html += `
        </tbody>
      </table>
    </div>
  `;

  wrap.innerHTML = html;
}

async function openModelTurnoverModal(modelName) {
  const modal = document.getElementById("modelTurnoverModal");
  const titleEl = document.getElementById("modelTurnoverModalTitle");
  const bodyEl = document.getElementById("modelTurnoverModalBody");

  if (!modal || !titleEl || !bodyEl) return;

  titleEl.textContent = `${modelName} 센터별 회전일 상세`;
  bodyEl.innerHTML = "불러오는 중...";
  modal.style.display = "flex";

  try {
    const resp = await fetch(`${API_URL}/inventory/model-turnover-detail`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ model_name: modelName })
    });

    const j = await resp.json();

    if (!resp.ok || !j.ok) {
      bodyEl.innerHTML = `❌ ${j.message || "조회 실패"}`;
      return;
    }

    const rows = Array.isArray(j.rows) ? j.rows : [];

    let html = `
      <div class="tableWrap compactTable">
        <table>
          <thead>
            <tr>
              <th>센터</th>
              <th>수량</th>
              <th>당월판매</th>
              <th>회전일</th>
            </tr>
          </thead>
          <tbody>
    `;

    rows.forEach(r => {
  const turnover =
    r.turnover_days === null || r.turnover_days === undefined
      ? null
      : Number(r.turnover_days);

  let turnoverText = "-";
  let turnoverClass = "";

  if (turnover !== null) {
    if (turnover >= 100) {
      turnoverText = `⚠️ ${turnover.toLocaleString()}일`;
      turnoverClass = "slowStockText";
    } else if (turnover < 40) {
      turnoverText = `부족 ${turnover.toLocaleString()}일`;
      turnoverClass = "shortStockText";
    } else {
      turnoverText = `${turnover.toLocaleString()}일`;
    }
  }

  html += `
    <tr>
      <td>${escapeHtml(r.agency_name || "")}</td>
      <td>${Number(r.qty || 0).toLocaleString()}</td>
      <td>${Number(r.monthly_sales || 0).toLocaleString()}</td>
      <td class="${turnoverClass}">${turnoverText}</td>
    </tr>
  `;
});

    html += `
          </tbody>
        </table>
      </div>
    `;

    bodyEl.innerHTML = html;
  } catch (e) {
    console.error(e);
    bodyEl.innerHTML = "❌ 네트워크 오류";
  }
}

function closeModelTurnoverModal() {
  const modal = document.getElementById("modelTurnoverModal");
  if (modal) {
    modal.style.display = "none";
  }
}

function toggleInventoryModelTurnoverSort(key) {
  if (inventoryModelTurnoverSort.key === key) {
    inventoryModelTurnoverSort.order =
      inventoryModelTurnoverSort.order === "asc" ? "desc" : "asc";
  } else {
    inventoryModelTurnoverSort.key = key;
    inventoryModelTurnoverSort.order = "asc";
  }

  renderInventoryModelTurnoverTable();
}

function getInventoryModelTurnoverSortIndicator(key) {
  if (inventoryModelTurnoverSort.key !== key) return "";
  return inventoryModelTurnoverSort.order === "asc" ? " ▲" : " ▼";
}

function renderStoreTable(rows) {
  const wrap = document.getElementById("byStoreTable");
  if (!wrap) return;

  const table = document.createElement("table");
  table.style.width = "100%";
  table.style.borderCollapse = "collapse";

  const thead = document.createElement("thead");
  const trh = document.createElement("tr");
  ["접점명", "수량", "상세"].forEach(label => {
    const th = document.createElement("th");
    th.textContent = label;
    th.style.textAlign = "left";
    th.style.borderBottom = "1px solid #ddd";
    th.style.padding = "8px";
    trh.appendChild(th);
  });
  thead.appendChild(trh);

  const tbody = document.createElement("tbody");

  rows.forEach(r => {
    const tr = document.createElement("tr");

    const td1 = document.createElement("td");
    td1.textContent = r.store_name || "";
    td1.style.borderBottom = "1px solid #f0f0f0";
    td1.style.padding = "8px";

    const td2 = document.createElement("td");
    let qty = Number(r.qty || 0);
    let text = qty.toLocaleString();

    if (qty >= 30) text = "🔥 " + text;
    if (qty <= 3) text = "⚠️ " + text;

    td2.textContent = text;
    td2.style.borderBottom = "1px solid #f0f0f0";
    td2.style.padding = "8px";

    const td3 = document.createElement("td");
    td3.style.borderBottom = "1px solid #f0f0f0";
    td3.style.padding = "8px";
    const btn = document.createElement("button");
    btn.textContent = "보기";
    btn.onclick = () => openStoreDetail(r.store_code);
    td3.appendChild(btn);

    tr.appendChild(td1);
    tr.appendChild(td2);
    tr.appendChild(td3);

    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);

  wrap.innerHTML = "";
  wrap.appendChild(table);
}

// =========================
// 판매점 상세
// =========================
async function openStoreDetail(storeCode) {
  if (!storeCode) return alert("store_code가 없습니다.");

  try {
    const resp = await fetch(`${API_URL}/inventory/store-detail`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        agency: currentCenter,
        store_code: storeCode
      })
    });

    const j = await resp.json();
    if (!resp.ok || !j.ok) {
      return alert(j.message || "상세 조회 실패");
    }

    const rows = Array.isArray(j.rows) ? j.rows : [];
    const storeName = rows[0]?.store_name || "판매점";
    const address = rows[0]?.address || "-";

    const modelMap = {};
    const colorSet = new Set();

    rows.forEach(r => {
      const model = (r.model_name || "").trim();
      const color = (r.color || "").trim();
      const key = `${model}__${color}`;

      colorSet.add(color);

      if (!modelMap[key]) {
        modelMap[key] = {
          model_name: model,
          color: color,
          qty: 0
        };
      }
      modelMap[key].qty += 1;
    });

    const summaryRows = Object.values(modelMap).sort((a, b) => {
      if (b.qty !== a.qty) return b.qty - a.qty;
      return (a.model_name || "").localeCompare(b.model_name || "", "ko");
    });

    const modelCount = new Set(rows.map(r => (r.model_name || "").trim())).size;
    const totalCount = rows.length;

    document.getElementById("storeModal").style.display = "flex";
    document.getElementById("modalTitle").innerText = `${storeName} 상세 재고`;

    const summaryCards = `
      <div class="storeSummaryGrid">
        <div class="miniStat">
          <div class="miniStatLabel">총 재고</div>
          <div class="miniStatValue">${totalCount.toLocaleString()} 대</div>
        </div>
        <div class="miniStat">
          <div class="miniStatLabel">모델 수</div>
          <div class="miniStatValue">${modelCount.toLocaleString()} 종</div>
        </div>
        <div class="miniStat">
          <div class="miniStatLabel">색상 수</div>
          <div class="miniStatValue">${colorSet.size.toLocaleString()} 개</div>
        </div>
        <div class="miniStat">
          <div class="miniStatLabel">주소</div>
          <div class="miniStatValue small">${escapeHtml(address)}</div>
        </div>
      </div>
    `;

    let summaryTable = `
      <div class="modalSectionTitle">모델별 집계</div>
      <div class="tableWrap">
        <table>
          <thead>
            <tr>
              <th>모델명</th>
              <th>색상</th>
              <th>수량</th>
            </tr>
          </thead>
          <tbody>
    `;

    summaryRows.forEach(r => {
      let qtyText = Number(r.qty).toLocaleString();
      if (r.qty >= 5) qtyText = `🔥 ${qtyText}`;

      summaryTable += `
        <tr>
          <td>${escapeHtml(r.model_name)}</td>
          <td>${escapeHtml(r.color)}</td>
          <td>${qtyText}</td>
        </tr>
      `;
    });

    summaryTable += `
          </tbody>
        </table>
      </div>
    `;

    let detailTable = `
      <div class="modalSectionTitle" style="margin-top:20px;">원본 상세 리스트</div>
      <div class="tableWrap">
        <table>
          <thead>
            <tr>
              <th>모델명</th>
              <th>색상</th>
              <th>일련번호</th>
              <th>애칭</th>
            </tr>
          </thead>
          <tbody>
    `;

    rows.forEach(r => {
      detailTable += `
        <tr>
          <td>${escapeHtml(r.model_name || "")}</td>
          <td>${escapeHtml(r.color || "")}</td>
          <td>${escapeHtml(r.serial_no || "")}</td>
          <td>${escapeHtml(r.nickname || "")}</td>
        </tr>
      `;
    });

    detailTable += `
          </tbody>
        </table>
      </div>
    `;

    document.getElementById("modalBody").innerHTML =
      summaryCards + summaryTable + detailTable;
  } catch (e) {
    console.error(e);
    alert("상세 조회 실패");
  }
}

function closeStoreModal() {
  document.getElementById("storeModal").style.display = "none";
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function escapeJs(value) {
  return String(value ?? "")
    .replace(/\\/g, "\\\\")
    .replace(/'/g, "\\'")
    .replace(/"/g, '\\"');
}

function renderRecommend(rows) {
  const wrap = document.getElementById("recommendMove");
  if (!wrap) return;

  if (!rows || !rows.length) {
    wrap.innerHTML = "추천 이동 대상 없음";
    return;
  }

  let html = `
    <div class="tableWrap">
      <table>
        <thead>
          <tr>
            <th>모델명</th>
            <th>색상</th>
            <th>보내는 곳</th>
            <th>보유</th>
            <th>받는 곳</th>
            <th>보유</th>
            <th>차이</th>
          </tr>
        </thead>
        <tbody>
  `;

  rows.forEach(r => {
    const gapText = Number(r.gap || 0) >= 5
      ? `🔥 ${Number(r.gap || 0).toLocaleString()}`
      : Number(r.gap || 0).toLocaleString();

    html += `
      <tr>
        <td>${escapeHtml(r.model_name || "")}</td>
        <td>${escapeHtml(r.color || "")}</td>
        <td>${escapeHtml(r.from_store_name || "")}</td>
        <td>${Number(r.from_qty || 0).toLocaleString()}</td>
        <td>${escapeHtml(r.to_store_name || "")}</td>
        <td>${Number(r.to_qty || 0).toLocaleString()}</td>
        <td>${gapText}</td>
      </tr>
    `;
  });

  html += `
        </tbody>
      </table>
    </div>
  `;

  wrap.innerHTML = html;
}

function toggleInventoryDetails() {
  const box = document.getElementById("inventoryDetailPanels");
  if (!box) return;

  if (box.style.display === "none" || box.style.display === "") {
    box.style.display = "block";
  } else {
    box.style.display = "none";
  }
}

async function loadDashboardDetail(type) {
  const section = document.getElementById("dashboardDetailSection");
  const titleEl = document.getElementById("dashboardDetailTitle");
  const hintEl = document.getElementById("dashboardDetailHint");
  const tableEl = document.getElementById("dashboardDetailTable");

  if (!section || !titleEl || !hintEl || !tableEl) return;

  if (activeDashboardDetailType === type && section.style.display !== "none") {
    section.style.display = "none";
    tableEl.innerHTML = "";
    activeDashboardDetailType = "";
    dashboardDetailRows = [];
    dashboardDetailSort = {
      key: "",
      order: "asc",
      type: ""
    };
    highlightActiveDashboardCard("");
    return;
  }

  section.style.display = "block";
  tableEl.innerHTML = "불러오는 중...";

  let url = "";
  let title = "";
  let hint = "";

  if (type === "warehouse") {
    url = `${API_URL}/inventory/warehouse-detail`;
    title = "창고 재고 상세";
    hint = "모델명 / 색상 / 수량";
  } else if (type === "total") {
    url = `${API_URL}/inventory/total-detail`;
    title = "총 재고 상세";
    hint = "모델 / 색상 / 총재고 / 창고재고 / 판매점재고 / 창고비중";
  } else if (type === "store") {
    url = `${API_URL}/inventory/store-stock-detail`;
    title = "판매점 재고 상세";
    hint = "대리점명 / 판매점명 / 모델명 / 수량";
  } else {
    section.style.display = "none";
    activeDashboardDetailType = "";
    return;
  }

  titleEl.textContent = title;
  hintEl.textContent = hint;

  try {
    const resp = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ agency: currentCenter })
    });

    const j = await resp.json();

    if (!resp.ok || !j.ok) {
      tableEl.innerHTML = `❌ ${j.message || "조회 실패"}`;
      activeDashboardDetailType = "";
      return;
    }

    dashboardDetailRows = Array.isArray(j.rows) ? j.rows : [];

    if (dashboardDetailSort.type !== type) {
      dashboardDetailSort = {
        type,
        key: "",
        order: "asc"
      };
    }

    activeDashboardDetailType = type;
    renderDashboardDetailByType(type);
    highlightActiveDashboardCard(type);
  } catch (e) {
    console.error(e);
    tableEl.innerHTML = "❌ 네트워크 오류";
    activeDashboardDetailType = "";
  }
}

function renderDashboardWarehouseDetail(rows) {
  const wrap = document.getElementById("dashboardDetailTable");
  if (!wrap) return;

  if (!rows || !rows.length) {
    wrap.innerHTML = "창고 재고 내역이 없습니다.";
    return;
  }

  let sortedRows = [...rows];
  if (dashboardDetailSort.type === "warehouse" && dashboardDetailSort.key) {
    sortedRows = sortDashboardRows(
      sortedRows,
      dashboardDetailSort.key,
      dashboardDetailSort.order
    );
  }

  let html = `
    <div class="tableWrap compactTable warehouseCompactTable">
      <table>
        <thead>
          <tr>
            <th class="sortable" onclick="toggleDashboardSort('warehouse','model_name')">
              모델명${getSortIndicator("warehouse", "model_name")}
            </th>
            <th class="sortable" onclick="toggleDashboardSort('warehouse','color')">
              색상${getSortIndicator("warehouse", "color")}
            </th>
            <th class="sortable" onclick="toggleDashboardSort('warehouse','qty')">
              수량${getSortIndicator("warehouse", "qty")}
            </th>
          </tr>
        </thead>
        <tbody>
  `;

  sortedRows.forEach(r => {
    html += `
      <tr>
        <td>${escapeHtml(r.model_name || "")}</td>
        <td>${escapeHtml(r.color || "")}</td>
        <td>${Number(r.qty || 0).toLocaleString()}</td>
      </tr>
    `;
  });

  html += `
        </tbody>
      </table>
    </div>
  `;

  wrap.innerHTML = html;
}

function renderDashboardTotalDetail(rows) {
  const wrap = document.getElementById("dashboardDetailTable");
  if (!wrap) return;

  if (!rows || !rows.length) {
    wrap.innerHTML = "총 재고 상세 내역이 없습니다.";
    return;
  }

  let sortedRows = [...rows];
  if (dashboardDetailSort.type === "total" && dashboardDetailSort.key) {
    sortedRows = sortDashboardRows(
      sortedRows,
      dashboardDetailSort.key,
      dashboardDetailSort.order
    );
  }

  let html = `
  <div class="tableWrap adminAgencyDetailWrap">
    <table class="adminAgencyDetailTable">
        <thead>
          <tr>
            <th class="sortable" onclick="toggleDashboardSort('total','model_name')">
              모델명${getSortIndicator("total", "model_name")}
            </th>
            <th class="sortable" onclick="toggleDashboardSort('total','color')">
              색상${getSortIndicator("total", "color")}
            </th>
            <th class="sortable" onclick="toggleDashboardSort('total','total_qty')">
              총재고${getSortIndicator("total", "total_qty")}
            </th>
            <th class="sortable" onclick="toggleDashboardSort('total','warehouse_qty')">
              창고재고${getSortIndicator("total", "warehouse_qty")}
            </th>
            <th class="sortable" onclick="toggleDashboardSort('total','store_qty')">
              판매점재고${getSortIndicator("total", "store_qty")}
            </th>
            <th class="sortable" onclick="toggleDashboardSort('total','warehouse_ratio')">
              창고비중${getSortIndicator("total", "warehouse_ratio")}
            </th>
          </tr>
        </thead>
        <tbody>
  `;

  sortedRows.forEach(r => {
    html += `
      <tr>
        <td>${escapeHtml(r.model_name || "")}</td>
        <td>${escapeHtml(r.color || "")}</td>
        <td>${Number(r.total_qty || 0).toLocaleString()}</td>
        <td>${Number(r.warehouse_qty || 0).toLocaleString()}</td>
        <td>${Number(r.store_qty || 0).toLocaleString()}</td>
        <td>${Number(r.warehouse_ratio || 0).toLocaleString()}%</td>
      </tr>
    `;
  });

  html += `
        </tbody>
      </table>
    </div>
  `;

  wrap.innerHTML = html;
}

function renderDashboardStoreDetail(rows) {
  const wrap = document.getElementById("dashboardDetailTable");
  if (!wrap) return;

  if (!rows || !rows.length) {
    wrap.innerHTML = "판매점 재고 상세 내역이 없습니다.";
    return;
  }

  let sortedRows = [...rows];
  if (dashboardDetailSort.type === "store" && dashboardDetailSort.key) {
    sortedRows = sortDashboardRows(
      sortedRows,
      dashboardDetailSort.key,
      dashboardDetailSort.order
    );
  }

  let html = `
  <div class="tableWrap inventoryStoreDetailWrap">
    <table class="inventoryStoreDetailTable">
        <thead>
          <tr>
            <th class="sortable" onclick="toggleDashboardSort('store','agency_name')">
              대리점명${getSortIndicator("store", "agency_name")}
            </th>
            <th class="sortable" onclick="toggleDashboardSort('store','store_name')">
              판매점명${getSortIndicator("store", "store_name")}
            </th>
            <th class="sortable" onclick="toggleDashboardSort('store','model_name')">
              모델명${getSortIndicator("store", "model_name")}
            </th>
            <th class="sortable" onclick="toggleDashboardSort('store','qty')">
              수량${getSortIndicator("store", "qty")}
            </th>
          </tr>
        </thead>
        <tbody>
  `;

  sortedRows.forEach(r => {
    html += `
      <tr>
        <td>${escapeHtml(r.agency_name || "")}</td>
        <td>${escapeHtml(r.store_name || "")}</td>
        <td>${escapeHtml(r.model_name || "")}</td>
        <td>${Number(r.qty || 0).toLocaleString()}</td>
      </tr>
    `;
  });

  html += `
        </tbody>
      </table>
    </div>
  `;

  wrap.innerHTML = html;
}

function highlightActiveDashboardCard(type) {
  document.querySelectorAll("#invCards .statCard").forEach(cardEl => {
    cardEl.classList.remove("active");
  });

  if (!type) return;

  const selector = `#invCards .statCard[onclick="loadDashboardDetail('${type}')"]`;
  const activeCard = document.querySelector(selector);
  if (activeCard) {
    activeCard.classList.add("active");
  }
}

function sortDashboardRows(rows, key, order) {
  const copied = [...rows];

  copied.sort((a, b) => {
    const av = a[key];
    const bv = b[key];

    const aNum = Number(av);
    const bNum = Number(bv);

    const aIsNum = av !== null && av !== undefined && av !== "" && !Number.isNaN(aNum);
    const bIsNum = bv !== null && bv !== undefined && bv !== "" && !Number.isNaN(bNum);

    let cmp = 0;

    if (aIsNum && bIsNum) {
      cmp = aNum - bNum;
    } else {
      cmp = String(av ?? "").localeCompare(String(bv ?? ""), "ko");
    }

    return order === "desc" ? -cmp : cmp;
  });

  return copied;
}

function toggleDashboardSort(type, key) {
  if (dashboardDetailSort.type === type && dashboardDetailSort.key === key) {
    dashboardDetailSort.order = dashboardDetailSort.order === "asc" ? "desc" : "asc";
  } else {
    dashboardDetailSort.type = type;
    dashboardDetailSort.key = key;
    dashboardDetailSort.order = "asc";
  }

  renderDashboardDetailByType(type);
}

function getSortIndicator(type, key) {
  if (dashboardDetailSort.type !== type || dashboardDetailSort.key !== key) {
    return "";
  }
  return dashboardDetailSort.order === "asc" ? " ▲" : " ▼";
}

function renderDashboardDetailByType(type) {
  if (!dashboardDetailRows || !dashboardDetailRows.length) {
    const wrap = document.getElementById("dashboardDetailTable");
    if (wrap) wrap.innerHTML = "데이터가 없습니다.";
    return;
  }

  if (type === "warehouse") {
    renderDashboardWarehouseDetail(dashboardDetailRows);
  } else if (type === "total") {
    renderDashboardTotalDetail(dashboardDetailRows);
  } else if (type === "store") {
    renderDashboardStoreDetail(dashboardDetailRows);
  }
}

function renderAgingStockCards(summary) {
  const wrap = document.getElementById("agingStockCards");
  const detailSection = document.getElementById("agingStockDetailSection");
  const detailTable = document.getElementById("agingStockDetailTable");

  if (!wrap) return;

  wrap.innerHTML = `
    <div class="statCard clickable" onclick="toggleAgingStockDetail(360)">
      <div class="statLabel">360일 초과</div>
      <div class="statValue">${Number(summary.over_360 || 0).toLocaleString()}대</div>
    </div>

    <div class="statCard clickable" onclick="toggleAgingStockDetail(500)">
      <div class="statLabel">500일 초과</div>
      <div class="statValue">${Number(summary.over_500 || 0).toLocaleString()}대</div>
    </div>

    <div class="statCard clickable" onclick="toggleAgingStockDetail(720)">
      <div class="statLabel">720일 초과</div>
      <div class="statValue">${Number(summary.over_720 || 0).toLocaleString()}대</div>
    </div>
  `;

  if (detailSection) detailSection.style.display = "none";
  if (detailTable) detailTable.innerHTML = "";
  activeAgingThreshold = 0;
}

async function toggleAgingStockDetail(threshold) {
  const detailSection = document.getElementById("agingStockDetailSection");
  const detailTitle = document.getElementById("agingStockDetailTitle");
  const detailTable = document.getElementById("agingStockDetailTable");

  if (!detailSection || !detailTitle || !detailTable) return;

  if (activeAgingThreshold === threshold && detailSection.style.display !== "none") {
    detailSection.style.display = "none";
    detailTable.innerHTML = "";
    activeAgingThreshold = 0;
    return;
  }

  detailSection.style.display = "block";
  detailTitle.textContent = `${threshold}일 초과 장기 입고 미소진 재고`;
  detailTable.innerHTML = "불러오는 중...";

  try {
    const resp = await fetch(`${API_URL}/inventory/aging-detail`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        agency: currentCenter,
        threshold
      })
    });

    const j = await resp.json();

    if (!resp.ok || !j.ok) {
      detailTable.innerHTML = `❌ ${j.message || "상세 조회 실패"}`;
      activeAgingThreshold = 0;
      return;
    }

        const rows = Array.isArray(j.rows) ? j.rows : [];

    if (!rows.length) {
      detailTable.innerHTML = "조건에 해당하는 재고가 없습니다.";
      activeAgingThreshold = threshold;
      agingStockDetailRows = [];
      return;
    }

    agingStockDetailRows = rows;
    agingStockDetailSort = {
      key: "aging_days",
      order: "desc"
    };

    renderAgingStockDetailTable();
    activeAgingThreshold = threshold;
  } catch (e) {
    console.error(e);
    detailTable.innerHTML = "❌ 네트워크 오류";
    activeAgingThreshold = 0;
  }
}

function renderAgingStockDetailTable() {
  const detailTable = document.getElementById("agingStockDetailTable");
  if (!detailTable) return;

  let rows = [...agingStockDetailRows];
  const key = agingStockDetailSort.key;
  const order = agingStockDetailSort.order;

  rows.sort((a, b) => {
    let av = a[key];
    let bv = b[key];

    const numericKeys = ["aging_days"];
    const isNumeric = numericKeys.includes(key);

    let cmp = 0;

    if (isNumeric) {
      cmp = Number(av || 0) - Number(bv || 0);
    } else {
      cmp = String(av || "").localeCompare(String(bv || ""), "ko");
    }

    return order === "asc" ? cmp : -cmp;
  });

  let html = `
  <div class="tableWrap inventoryAgingWrap">
    <table class="inventoryAgingTable">
        <thead>
          <tr>
            <th class="sortable" onclick="toggleAgingStockDetailSort('agency_name')">
              센터${getAgingStockDetailSortIndicator('agency_name')}
            </th>
            <th class="sortable" onclick="toggleAgingStockDetailSort('model_name')">
              모델${getAgingStockDetailSortIndicator('model_name')}
            </th>
            <th class="sortable" onclick="toggleAgingStockDetailSort('color')">
              색상${getAgingStockDetailSortIndicator('color')}
            </th>
            <th class="sortable" onclick="toggleAgingStockDetailSort('serial_no')">
              일련번호${getAgingStockDetailSortIndicator('serial_no')}
            </th>
            <th class="sortable" onclick="toggleAgingStockDetailSort('aging_days')">
              입고경과일${getAgingStockDetailSortIndicator('aging_days')}
            </th>
          </tr>
        </thead>
        <tbody>
  `;

  rows.forEach(r => {
    html += `
      <tr>
        <td>${escapeHtml(r.agency_name || "")}</td>
        <td>${escapeHtml(r.model_name || "")}</td>
        <td>${escapeHtml(r.color || "")}</td>
        <td>${escapeHtml(r.serial_no || "")}</td>
        <td>${Number(r.aging_days || 0).toLocaleString()}일</td>
      </tr>
    `;
  });

  html += `
        </tbody>
      </table>
    </div>
  `;

  detailTable.innerHTML = html;
}

function toggleAgingStockDetailSort(key) {
  if (agingStockDetailSort.key === key) {
    agingStockDetailSort.order =
      agingStockDetailSort.order === "asc" ? "desc" : "asc";
  } else {
    agingStockDetailSort.key = key;
    agingStockDetailSort.order = "asc";
  }

  renderAgingStockDetailTable();
}

function getAgingStockDetailSortIndicator(key) {
  if (agingStockDetailSort.key !== key) return "";
  return agingStockDetailSort.order === "asc" ? " ▲" : " ▼";
}

function renderAdminAgencyStockCards(rows) {
  const wrap = document.getElementById("adminAgencyStockCards");
  if (!wrap) return;

  const order = ["광주", "목포", "순천", "전북", "제주"];
  const map = {};

  rows.forEach(r => {
    map[r.agency_name] = Number(r.total_qty || 0);
  });

  wrap.innerHTML = order.map(name => `
    <div
      class="statCard clickable ${activeAdminAgencyDetail === name ? "active" : ""}"
      onclick="toggleAdminAgencyDetail('${name}')"
    >
      <div class="statLabel">${name}</div>
      <div class="statValue">${Number(map[name] || 0).toLocaleString()}대</div>
    </div>
  `).join("");
}

async function toggleAdminAgencyDetail(agencyName) {
  const section = document.getElementById("adminAgencyDetailSection");
  const titleEl = document.getElementById("adminAgencyDetailTitle");
  const tableEl = document.getElementById("adminAgencyDetailTable");

  if (!section || !titleEl || !tableEl) return;

  // 같은 카드 다시 누르면 닫기
  if (activeAdminAgencyDetail === agencyName && section.style.display !== "none") {
    section.style.display = "none";
    tableEl.innerHTML = "";
    activeAdminAgencyDetail = "";
    reloadAdminAgencyCardHighlight();
    return;
  }

  section.style.display = "block";
  titleEl.textContent = `${agencyName} 총 재고 상세`;
  tableEl.innerHTML = "불러오는 중...";

  try {
    const resp = await fetch(`${API_URL}/inventory/agency-detail`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ agency_name: agencyName })
    });

    const j = await resp.json();

    if (!resp.ok || !j.ok) {
      tableEl.innerHTML = `❌ ${j.message || "조회 실패"}`;
      activeAdminAgencyDetail = "";
      reloadAdminAgencyCardHighlight();
      return;
    }

    const rows = Array.isArray(j.rows) ? j.rows : [];
    sortAdminAgencyRows(rows);

    if (!rows.length) {
      tableEl.innerHTML = "데이터가 없습니다.";
      activeAdminAgencyDetail = agencyName;
      reloadAdminAgencyCardHighlight();
      return;
    }

    let html = `
      <div class="tableWrap inventoryTotalDetailWrap">
        <table class="inventoryTotalDetailTable">
          <thead>
  <th class="sortable" onclick="changeAdminAgencySort('model_name')">
  모델명 ${getSortArrow('model_name')}
</th>

<th>색상</th>

<th class="sortable" onclick="changeAdminAgencySort('total_qty')">
  총재고 ${getSortArrow('total_qty')}
</th>

<th class="sortable" onclick="changeAdminAgencySort('warehouse_qty')">
  창고재고 ${getSortArrow('warehouse_qty')}
</th>

<th class="sortable" onclick="changeAdminAgencySort('store_qty')">
  판매점재고 ${getSortArrow('store_qty')}
</th>

<th class="sortable" onclick="changeAdminAgencySort('warehouse_ratio')">
  창고비중 ${getSortArrow('warehouse_ratio')}
</th>
</thead>
          <tbody>
    `;

    rows.forEach(r => {
      html += `
        <tr>
          <td>${escapeHtml(r.model_name || "")}</td>
          <td>${escapeHtml(r.color || "")}</td>
          <td>${Number(r.total_qty || 0).toLocaleString()}</td>
          <td>${Number(r.warehouse_qty || 0).toLocaleString()}</td>
          <td>${Number(r.store_qty || 0).toLocaleString()}</td>
          <td>${Number(r.warehouse_ratio || 0).toLocaleString()}%</td>
        </tr>
      `;
    });

    html += `
          </tbody>
        </table>
      </div>
    `;

    tableEl.innerHTML = html;
    activeAdminAgencyDetail = agencyName;
    reloadAdminAgencyCardHighlight();
  } catch (e) {
    console.error(e);
    tableEl.innerHTML = "❌ 네트워크 오류";
    activeAdminAgencyDetail = "";
    reloadAdminAgencyCardHighlight();
  }
}

function sortAdminAgencyRows(rows) {
  return rows.sort((a, b) => {
    const key = adminAgencySortKey;

    let av = a[key];
    let bv = b[key];

    if (typeof av === "string") av = av.toLowerCase();
    if (typeof bv === "string") bv = bv.toLowerCase();

    if (av < bv) return adminAgencySortDir === "asc" ? -1 : 1;
    if (av > bv) return adminAgencySortDir === "asc" ? 1 : -1;
    return 0;
  });
}

function changeAdminAgencySort(key) {
  if (adminAgencySortKey === key) {
    adminAgencySortDir = adminAgencySortDir === "asc" ? "desc" : "asc";
  } else {
    adminAgencySortKey = key;
    adminAgencySortDir = "desc";
  }

  // 다시 렌더
  if (activeAdminAgencyDetail) {
    toggleAdminAgencyDetail(activeAdminAgencyDetail);
  }
}

function reloadAdminAgencyCardHighlight() {
  const cards = document.querySelectorAll("#adminAgencyStockCards .statCard");
  cards.forEach(card => card.classList.remove("active"));

  if (!activeAdminAgencyDetail) return;

  const target = document.querySelector(
    `#adminAgencyStockCards .statCard[onclick="toggleAdminAgencyDetail('${activeAdminAgencyDetail}')"]`
  );

  if (target) {
    target.classList.add("active");
  }
}

// =========================
// 접속자 로그
// =========================
async function saveLoginLog() {
  if (!currentCenter) return;

  try {
    await fetch(`${API_URL}/login-log`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ agency: currentCenter })
    });
  } catch (e) {
    console.error("로그인 기록 저장 실패", e);
  }
}

async function loadTodayLoginInfo() {
  if (!currentCenter) return;

  const info = document.getElementById("todayLoginInfo");
  if (!info) return;

  try {
    const resp = await fetch(
      `${API_URL}/login-today-summary?agency=${encodeURIComponent(currentCenter)}`
    );
    const j = await resp.json();

    if (!resp.ok || !j.ok) {
      info.textContent = "오늘 접속자 조회 실패";
      return;
    }

    if (currentCenter === "관리자") {
      const order = ["광주", "목포", "순천", "전북", "제주"];
      const map = {};

      (j.rows || []).forEach(r => {
        map[r.agency_name] = Number(r.cnt || 0);
      });

      info.textContent = order
        .map(name => `${name} ${Number(map[name] || 0)}명`)
        .join(" · ");
    } else {
      info.textContent = `오늘 접속자 ${Number(j.count || 0)}명`;
    }
  } catch (e) {
    console.error("오늘 접속자 조회 실패", e);
    info.textContent = "오늘 접속자 조회 실패";
  }
}

// =========================
// 실적조회 화면 초기화
// =========================
async function loadPerformanceDashboard() {
  const cardWrap = document.getElementById("performanceDashCards");
  const chartWrap = document.getElementById("perfChartArea");
  const chartTitle = document.getElementById("perfChartTitle");
  const resultWrap = document.getElementById("performanceSearchResult");
  const statusWrap = document.getElementById("performanceSearchStatus");
  const monthWrap = document.getElementById("performanceDashMonth");

  if (cardWrap) {
    cardWrap.innerHTML = "불러오는 중...";
  }

  if (chartWrap) {
    chartWrap.innerHTML = "카드를 클릭하면 그래프가 표시됩니다.";
  }

  if (chartTitle) {
    chartTitle.textContent = "📈 최근 6개월 추이";
  }

  if (resultWrap) {
    resultWrap.innerHTML = "";
  }

  if (statusWrap) {
    statusWrap.innerText = "";
  }

  if (monthWrap) {
    monthWrap.textContent = "기준월: -";
  }

  try {
    const queryCenter = getPerformanceQueryCenter();

const resp = await fetch(
  `${API_URL}/performance/dashboard-summary?agency=${encodeURIComponent(queryCenter)}`
);
    const j = await resp.json();

    if (!resp.ok || !j.ok) {
      if (cardWrap) cardWrap.innerHTML = `❌ ${j.message || "실적 대시보드 조회 실패"}`;
      return;
    }

    const s = j.summary || {};
const latestMonth = j.latest_month || "-";
const latestDate = j.latest_date || "";
const progressRate = Number(j.progress_rate || 0);

    if (monthWrap) {
  const displayDate = latestDate ? String(latestDate).slice(0, 10) : "";
  monthWrap.textContent = displayDate
    ? `기준일: ${displayDate} · 표준 진척율: ${progressRate}%`
    : "기준일: - · 표준 진척율: 0%";
}

    if (cardWrap) {
  cardWrap.innerHTML = `
    <div class="statCard clickable performanceDashCard" onclick="loadPerformanceTrend('후불', '후불 최근 6개월 추이')">
      <div class="statLabel">후불</div>
      <div class="statValue">
        ${Number(s.postpaid || 0).toLocaleString()}
        <span class="rateText">(${Number(s.postpaid_rate || 0)}%)</span>
      </div>
    </div>

    <div class="statCard clickable performanceDashCard" onclick="loadPerformanceTrend('순신규', '순신규 최근 6개월 추이')">
      <div class="statLabel">순신규</div>
      <div class="statValue">
        ${Number(s.pure_new || 0).toLocaleString()}
        <span class="rateText">(${Number(s.pure_new_rate || 0)}%)</span>
      </div>
    </div>

    <div class="statCard clickable performanceDashCard" onclick="loadPerformanceTrend('약정갱신', '약정갱신 최근 6개월 추이')">
      <div class="statLabel">약정갱신</div>
      <div class="statValue">
        ${Number(s.renewal || 0).toLocaleString()}
        <span class="rateText">(${Number(s.renewal_rate || 0)}%)</span>
      </div>
    </div>

    <div class="statCard clickable performanceDashCard" onclick="loadPerformanceTrend('후불실적점', '후불 실적점 최근 6개월 추이')">
      <div class="statLabel">후불 실적점</div>
      <div class="statValue">
        ${Number(s.postpaid_store_count || 0).toLocaleString()}
        <span class="rateText">(${Number(s.postpaid_store_rate || 0)}%)</span>
      </div>
    </div>

    <div class="statCard clickable performanceDashCard" onclick="loadPerformanceTrend('MIT', 'MIT 최근 6개월 추이')">
      <div class="statLabel">MIT 당월 실적</div>
      <div class="statValue">
        ${Number(s.mit || 0).toLocaleString()}
        <span class="rateText">(${Number(s.mit_rate || 0)}%)</span>
      </div>
    </div>
  `;
}
  } catch (e) {
    console.error(e);
    if (cardWrap) cardWrap.innerHTML = "❌ 네트워크 오류";
  }
}

async function loadPerformanceTrend(metric, title) {
  const chartWrap = document.getElementById("perfChartArea");
  const chartTitle = document.getElementById("perfChartTitle");

  if (!chartWrap) return;

  chartWrap.innerHTML = "불러오는 중...";
  if (chartTitle) chartTitle.textContent = `📈 ${title}`;

  try {
    const queryCenter = getPerformanceQueryCenter();

const resp = await fetch(
  `${API_URL}/performance/dashboard-trend?metric=${encodeURIComponent(metric)}&agency=${encodeURIComponent(queryCenter)}`
);
    const j = await resp.json();

    if (!resp.ok || !j.ok) {
      chartWrap.innerHTML = `❌ ${j.message || "추이 조회 실패"}`;
      return;
    }

    const rows = Array.isArray(j.rows) ? j.rows : [];

    if (!rows.length) {
      chartWrap.innerHTML = "표시할 데이터가 없습니다.";
      return;
    }

    if (performanceTrendChart) {
      performanceTrendChart.destroy();
      performanceTrendChart = null;
    }

    chartWrap.innerHTML = `
      <div class="perfChartCanvasWrap">
        <canvas id="performanceTrendCanvas"></canvas>
      </div>
    `;

    const canvas = document.getElementById("performanceTrendCanvas");
    if (!canvas) {
      chartWrap.innerHTML = "차트 영역 생성 실패";
      return;
    }

    // ✅ X축은 월만 표시
    const labels = rows.map(r => {
      const monthText = String(r.month || "");
      const mm = monthText.slice(5, 7);
      return `${mm}월`;
    });

    const values = rows.map(r => Number(r.value || 0));

    const ctx = canvas.getContext("2d");

    performanceTrendChart = new Chart(ctx, {
      data: {
        labels,
        datasets: [
          {
            type: "bar",
            label: title,
            data: values,
            borderWidth: 1,
            borderRadius: 6,
            barThickness: 28,
            maxBarThickness: 36,
            backgroundColor: "rgba(91, 141, 239, 0.55)",
            borderColor: "rgba(91, 141, 239, 0.95)"
          },
          {
            type: "line",
            label: "추세",
            data: values,
            borderColor: "rgba(91, 141, 239, 0.35)",
            backgroundColor: "rgba(91, 141, 239, 0.12)",
            borderWidth: 2,
            pointRadius: 3,
            pointHoverRadius: 4,
            pointBackgroundColor: "rgba(91, 141, 239, 0.45)",
            pointBorderColor: "rgba(91, 141, 239, 0.45)",
            tension: 0.35,
            fill: false
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        animation: {
          duration: 400
        },
        plugins: {
          legend: {
            display: false
          },
          tooltip: {
            callbacks: {
              label: function(context) {
                return `${Number(context.raw || 0).toLocaleString()}`;
              }
            }
          }
        },
        scales: {
          x: {
            ticks: {
              maxRotation: 0,
              minRotation: 0,
              autoSkip: false,
              font: {
                size: 12
              }
            },
            grid: {
              display: false
            }
          },
          y: {
            beginAtZero: true,
            ticks: {
              callback: function(value) {
                return Number(value).toLocaleString();
              },
              font: {
                size: 12
              }
            }
          }
        }
      }
    });
  } catch (e) {
    console.error(e);
    chartWrap.innerHTML = "❌ 네트워크 오류";
  }
}

// =========================
// 실적조회 검색 (백엔드 API 연결 전 임시)
// =========================
async function searchPerformanceSales() {
  const startMonth = document.getElementById("salesStartMonth")?.value.trim();
  const endMonth = document.getElementById("salesEndMonth")?.value.trim();
  const region = document.getElementById("salesRegion")?.value.trim() || "";
  const agencyName = document.getElementById("salesAgency")?.value.trim() || "";
  const storeName = document.getElementById("salesStore")?.value.trim() || "";

  const status = document.getElementById("performanceSearchStatus");
  const result = document.getElementById("performanceSearchResult");

  if (!startMonth || !endMonth) {
    status.innerText = "조회기간 필수";
    return;
  }

  status.innerText = "조회 중...";
  result.innerHTML = "";

  try {
    const resp = await fetch(`${API_URL}/performance/search`, {
      method: "POST",
      headers: {"Content-Type":"application/json"},
      body: JSON.stringify({
        start_month: startMonth,
        end_month: endMonth,
        region,
        agency_name: agencyName,
        store_name: storeName
      })
    });

    const j = await resp.json();

    if (!j.ok) {
      status.innerText = j.message || "조회 실패";
      return;
    }

    performanceSearchCache = Array.isArray(j.rows) ? j.rows : [];
    performanceSortState = { key: "", order: "asc" };

    renderPerformanceSearchTable();
    status.innerText = `총 ${performanceSearchCache.length}건`;
  } catch (e) {
    console.error(e);
    status.innerText = "네트워크 오류";
  }
}

async function loadDetail(baseMonth, market, agency) {
  const result = document.getElementById("performanceSearchResult");

  try {
    const resp = await fetch(`${API_URL}/performance/detail`, {
      method: "POST",
      headers: {"Content-Type":"application/json"},
      body: JSON.stringify({
        base_month: baseMonth,
        market,
        agency_name: agency
      })
    });

    const j = await resp.json();

    if (!j.ok) {
      alert("상세 조회 실패");
      return;
    }

    let html = `
      <div style="margin-bottom:12px;">
        <button class="btn ghost" type="button" onclick="renderPerformanceSearchTable()">
          ← 뒤로 가기
        </button>
      </div>

      <div class="tableWrap performanceDetailTableWrap">
  <table class="performanceDetailTable">
          <thead>
            <tr>
              <th>기간</th>
              <th>지역</th>
              <th>대리점</th>
              <th>판매점코드</th>
              <th>판매점명</th>
              <th>후불</th>
              <th>순신규</th>
            </tr>
          </thead>
          <tbody>
    `;

    j.rows.forEach(r => {
      html += `
        <tr>
          <td>${escapeHtml(r.base_month || "")}</td>
          <td>${escapeHtml(r.market || "-")}</td>
          <td>${escapeHtml(r.agency_name || "")}</td>
          <td>${escapeHtml(r.store_code || "")}</td>
          <td>${escapeHtml(r.store_name || "")}</td>
          <td>${Number(r.postpaid || 0).toLocaleString()}</td>
          <td>${Number(r.pure_new || 0).toLocaleString()}</td>
        </tr>
      `;
    });

    html += `
          </tbody>
        </table>
      </div>
    `;

    result.innerHTML = html;
  } catch (e) {
    console.error(e);
    alert("상세 조회 실패");
  }
}

function renderPerformanceSearchTable() {
  const result = document.getElementById("performanceSearchResult");
  if (!result) return;

  let rows = [...performanceSearchCache];

  if (performanceSortState.key) {
    rows.sort((a, b) => {
      const key = performanceSortState.key;
      let av = a[key];
      let bv = b[key];

      const numericKeys = ["postpaid", "pure_new", "renewal", "mit"];
      const isNumeric = numericKeys.includes(key);

      let cmp = 0;

      if (isNumeric) {
        cmp = Number(av || 0) - Number(bv || 0);
      } else {
        cmp = String(av || "").localeCompare(String(bv || ""), "ko");
      }

      return performanceSortState.order === "desc" ? -cmp : cmp;
    });
  }

  let html = `
    <div class="tableWrap">
      <table>
        <thead>
          <tr>
            <th class="sortable" onclick="togglePerformanceSort('base_month')">
              기간${getPerformanceSortIndicator("base_month")}
            </th>
            <th class="sortable" onclick="togglePerformanceSort('market')">
              지역${getPerformanceSortIndicator("market")}
            </th>
            <th class="sortable" onclick="togglePerformanceSort('agency_name')">
              대리점${getPerformanceSortIndicator("agency_name")}
            </th>
            <th class="sortable" onclick="togglePerformanceSort('postpaid')">
              후불${getPerformanceSortIndicator("postpaid")}
            </th>
            <th class="sortable" onclick="togglePerformanceSort('pure_new')">
              순신규${getPerformanceSortIndicator("pure_new")}
            </th>
            <th class="sortable" onclick="togglePerformanceSort('renewal')">
              약정갱신${getPerformanceSortIndicator("renewal")}
            </th>
            <th class="sortable" onclick="togglePerformanceSort('mit')">
              MIT${getPerformanceSortIndicator("mit")}
            </th>
            <th>상세</th>
          </tr>
        </thead>
        <tbody>
  `;

  rows.forEach(r => {
    const market = r.market || "";
    const agency = r.agency_name || "";

    html += `
      <tr>
        <td>${escapeHtml(r.base_month || "")}</td>
        <td>${escapeHtml(market || "-")}</td>
        <td>${escapeHtml(agency)}</td>
        <td>${Number(r.postpaid || 0).toLocaleString()}</td>
        <td>${Number(r.pure_new || 0).toLocaleString()}</td>
        <td>${Number(r.renewal || 0).toLocaleString()}</td>
        <td>${Number(r.mit || 0).toLocaleString()}</td>
        <td>
          <button onclick="loadDetail('${escapeJs(r.base_month || "")}','${escapeJs(market)}','${escapeJs(agency)}')">
            보기
          </button>
        </td>
      </tr>
    `;
  });

  html += `
        </tbody>
      </table>
    </div>
  `;

  result.innerHTML = html;
}

function togglePerformanceSort(key) {
  if (performanceSortState.key === key) {
    performanceSortState.order =
      performanceSortState.order === "asc" ? "desc" : "asc";
  } else {
    performanceSortState.key = key;
    performanceSortState.order = "asc";
  }

  renderPerformanceSearchTable();
}

function getPerformanceSortIndicator(key) {
  if (performanceSortState.key !== key) return "";
  return performanceSortState.order === "asc" ? " ▲" : " ▼";
}

async function searchStoreFinder() {
  const keyword = document.getElementById("storeFinderKeyword")?.value.trim();
  const status = document.getElementById("storeFinderStatus");
  const result = document.getElementById("storeFinderResult");

  if (!keyword) {
    if (status) status.innerText = "검색어를 입력하세요.";
    return;
  }

  if (status) status.innerText = "검색 중...";
  if (result) result.innerHTML = "";

  try {
    const resp = await fetch(`${API_URL}/store-master/search`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ keyword })
    });

    const j = await resp.json();

    if (!resp.ok || !j.ok) {
      if (status) status.innerText = j.message || "판매점 찾기 실패";
      return;
    }

    const rows = Array.isArray(j.rows) ? j.rows : [];

    if (status) {
      status.innerText = `총 ${rows.length}건`;
    }

    if (!rows.length) {
      if (result) {
        result.innerHTML = `
          <div class="panel">
            <div class="panelHeader">
              <div class="panelTitle">검색 결과 없음</div>
              <div class="panelHint">판매점코드 / 판매점명 / 주소로 다시 검색해보세요.</div>
            </div>
          </div>
        `;
      }
      return;
    }

    let html = `
      <div class="storeFinderGuide">
        판매점코드를 누르면 복사됩니다.
      </div>

      <div class="tableWrap storeFinderTableWrap">
        <table class="storeFinderTable">
          <thead>
            <tr>
              <th>판매점코드</th>
              <th>판매점명</th>
              <th>판매점주소</th>
            </tr>
          </thead>
          <tbody>
    `;

    rows.forEach(r => {
      const code = r.store_code || "";
      const name = r.store_name || "";
      const address = r.address || "";

      html += `
        <tr>
          <td>
            <button
              class="copyCodeBtn"
              type="button"
              onclick="copyStoreCode('${escapeJs(code)}')"
            >
              ${escapeHtml(code)}
            </button>
          </td>
          <td>${escapeHtml(name)}</td>
          <td>${escapeHtml(address)}</td>
        </tr>
      `;
    });

    html += `
          </tbody>
        </table>
      </div>
    `;

    if (result) result.innerHTML = html;
  } catch (e) {
    console.error(e);
    if (status) status.innerText = "네트워크 오류";
  }
}

async function copyStoreCode(code) {
  try {
    await navigator.clipboard.writeText(code);
    alert(`판매점코드 복사됨: ${code}`);
  } catch (e) {
    console.error(e);
    alert("복사 실패");
  }
}

function getSortArrow(key) {
  if (adminAgencySortKey !== key) return "";
  return adminAgencySortDir === "asc" ? " ▲" : " ▼";
}