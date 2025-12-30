let currentCenter = "";

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

// 🔥 중요: 백엔드 Render 주소
const API_URL = "https://nexus-inventory-site.onrender.com";

// =========================
// 자동 로그인 (하루 유지)
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

    document.getElementById("loginBox").style.display = "none";
    document.getElementById("logoutBtn").style.display = "inline-block";
    document.getElementById("brandSub").innerText =
      center === "관리자" ? "관리자 모드" : `${center}센터 로그인됨`;

    document.querySelector(".hero").style.display = "none";

    if (center === "관리자") {
      document.getElementById("uploadBox").style.display = "block";
      document.getElementById("searchBox").style.display = "none";
    } else {
      document.getElementById("searchBox").style.display = "block";
      document.getElementById("uploadBox").style.display = "none";
    }
  } catch {
    localStorage.removeItem("loginInfo");
  }
});

// =========================
// 로그인 / 로그아웃
// =========================
function login() {
  const center = document.getElementById("centerSelect").value;
  const inputPassword = document.getElementById("password").value.trim();

  if (!center) return alert("센터를 선택하세요.");

  const isMaster = inputPassword === MASTER_PASSWORD;
  const isCenterValid = inputPassword === passwords[center];

  if (!isMaster && !isCenterValid) {
    return alert("비밀번호가 틀렸습니다.");
  }

  // ⭐ 센터는 선택값 그대로 유지
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
    center === "관리자" ? "관리자 모드" : `${center}센터 로그인됨`;

  document.querySelector(".hero").style.display = "none";

  if (center === "관리자") {
    document.getElementById("uploadBox").style.display = "block";
    document.getElementById("searchBox").style.display = "none";
  } else {
    document.getElementById("searchBox").style.display = "block";
    document.getElementById("uploadBox").style.display = "none";
  }
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
// 재고 검색 (🔥 기본 정렬: 보유처)
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
        center: currentCenter,
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

    // ✅ 기본 정렬: 보유처 기준 오름차순
    j.table.sort((a, b) =>
      (a["보유처"] || "").localeCompare(b["보유처"] || "", "ko")
    );

    // 상세 보기 여부
    const detail = document.getElementById("detailToggle")?.checked || false;

    const baseCols = ["보유처", "모델명", "색상", "일련번호"];
    const detailCols = ["상권주소", "펫네임", "애칭"];

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

  cols.forEach(c => {
    const th = document.createElement("th");
    th.textContent = c;
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
