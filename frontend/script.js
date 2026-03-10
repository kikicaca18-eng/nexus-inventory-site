let currentCenter = ""; // 👉 실제 의미: 대리점명

function $(id) {
  return document.getElementById(id);
}

function setDisplay(id, display) {
  const el = $(id);
  if (el) el.style.display = display;
}

// =========================
// 비밀번호
// =========================
const passwords = {
  "광주": "20405",
  "목포": "20001",
  "순천": "20404",
  "전북": "20407",
  "제주": "20403",
  "관리자": "8673"
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
    setDisplay("uploadBox", "none");
    setDisplay("searchBox", "none");
    setDisplay("inventoryDash", "none");
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
document.getElementById("uploadBox").style.display = "none";
document.getElementById("searchBox").style.display = "none";
document.getElementById("inventoryDash").style.display = "none";
}

function logout() {
  currentCenter = "";
  localStorage.removeItem("loginInfo");
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
        agency: currentCenter, // 🔥 핵심: 로그인한 대리점명
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

    // 기본 정렬: 접점명(판매점/창고)
    j.table.sort((a, b) =>
      (a.store_name || "").localeCompare(b.store_name || "", "ko")
    );

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
  wrap.className = "tableWrap";

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
  const all = ["menuBox", "inventoryDash", "searchBox", "uploadBox"];
  all.forEach(id => setDisplay(id, ids.includes(id) ? "block" : "none"));
}


function openInventory() {
  if (currentCenter === "관리자") {
    showOnly(["menuBox", "inventoryDash", "uploadBox", "searchBox"]);
  } else {
    showOnly(["menuBox", "inventoryDash", "searchBox"]);
  }

  const detailBox = document.getElementById("inventoryDetailPanels");
  if (detailBox) detailBox.style.display = "none";

  loadInventoryDashboard();
}

function openPerformance() {
  showOnly([]);
  alert("실적조회는 다음 단계에서 붙일게! 이제 바로 실적 화면 만들자.");
}

// =========================
// 재고 대시보드 로드
// =========================
async function loadInventoryDashboard() {
  // 카드 영역 초기화
  const cards = document.getElementById("invCards");
  if (cards) cards.innerHTML = "불러오는 중...";

  try {
    // 1) 요약
    const sResp = await fetch(`${API_URL}/inventory/summary-extended`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ agency: currentCenter })
    });
    const s = await sResp.json();
    if (!sResp.ok || !s.ok) throw new Error(s.message || "요약 실패");

    renderInvCards(s.summary);

    // 2) 모델 TOP
    const mResp = await fetch(`${API_URL}/inventory/by-model`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ agency: currentCenter, limit: 20 })
    });
    const m = await mResp.json();
    if (!mResp.ok || !m.ok) throw new Error(m.message || "모델 TOP 실패");

    renderSimpleTable("byModelTable", m.rows, [
      { key: "model_name", label: "모델명" },
      { key: "qty", label: "수량" }
    ]);

    // 3) 판매점 TOP
    const stResp = await fetch(`${API_URL}/inventory/by-store`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ agency: currentCenter, limit: 30 })
    });
    const st = await stResp.json();
    if (!stResp.ok || !st.ok) throw new Error(st.message || "판매점 TOP 실패");

    renderStoreTable(st.rows);

    // 4) 추천 이동 모델
const rResp = await fetch(`${API_URL}/inventory/recommend-move`, {
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify({ agency: currentCenter })
});
const rec = await rResp.json();
if (rResp.ok && rec.ok) {
  renderRecommend(rec.rows);
}

  } catch (e) {
    if (cards) cards.innerHTML = `❌ 대시보드 로드 실패: ${e.message}`;
  }
}

function renderInvCards(summary) {
  const cards = document.getElementById("invCards");
  if (!cards) return;

  const { total_qty, store_cnt, model_cnt, warehouse_qty, store_qty } = summary;

  // 상단 pill 표시 (있으면)
  const pillAgency = document.getElementById("pillAgency");
  if (pillAgency) pillAgency.textContent = currentCenter;

  cards.innerHTML = `
    ${card("오늘 총 재고", total_qty, "대")}
    ${card("판매점 수", store_cnt, "곳")}
    ${card("모델 수", model_cnt, "종")}
    ${card("창고 재고", warehouse_qty, "대")}
    ${card("판매점 재고", store_qty, "대")}
  `;
}

function card(title, value, unit) {
  return `
    <div class="statCard">
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
let displayVal =
  typeof val === "number" ? val.toLocaleString() : (val ?? "").toString();

if (c.key === "qty" && typeof val === "number") {
  if (val >= 30) {
    displayVal = "🔥 " + displayVal;
  } else if (val <= 2) {
    displayVal = "⚠️ " + displayVal;
  }
}

td.textContent = displayVal;

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

// 판매점 상세(일단 alert로 보여주고, 다음에 모델로 예쁘게)
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

    // 1. 모델별 집계
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

    // 2. 모달 제목
    document.getElementById("storeModal").style.display = "flex";
    document.getElementById("modalTitle").innerText = `${storeName} 상세 재고`;

    // 3. 상단 요약 카드
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

    // 4. 모델별 집계 테이블
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

    // 5. 원본 상세 리스트
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