let currentCenter = "";

const passwords = {
  "광주": "yc20405",
  "목포": "yd20001",
  "순천": "yc20404",
  "전북": "yc20407",
  "제주": "yc20403",
  "관리자": "41218673"
};

const API_URL = "https://nexus-inventory-site.onrender.com";

function login() {
  const center = document.getElementById("centerSelect").value;
  const pw = document.getElementById("password").value;

  if (!center) return alert("센터를 선택하세요.");
  if (pw !== passwords[center]) return alert("비밀번호가 틀렸습니다.");

  currentCenter = center;

  document.getElementById("loginBox").style.display = "none";
  document.getElementById("logoutBtn").style.display = "inline-block";
  document.getElementById("brandSub").innerText = (center === "관리자") ? "관리자 모드" : `${center}센터 로그인됨`;

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
  document.getElementById("loginBox").style.display = "block";
  document.getElementById("searchBox").style.display = "none";
  document.getElementById("uploadBox").style.display = "none";
  document.getElementById("logoutBtn").style.display = "none";
  document.getElementById("brandSub").innerText = "";
  document.getElementById("password").value = "";
  document.getElementById("result").innerHTML = "";
  document.getElementById("status").innerText = "";
}

async function uploadExcel() {
  const fileInput = document.getElementById("excelFile");
  const status = document.getElementById("uploadStatus");

  if (!fileInput.files.length) {
    status.innerText = "엑셀 파일을 선택해주세요.";
    return;
  }

  const formData = new FormData();
  formData.append("file", fileInput.files[0]);

  status.innerText = "업로드 중... (Render 첫 요청이면 조금 느릴 수 있어요)";
  try {
    const resp = await fetch(`${API_URL}/upload`, { method: "POST", body: formData });
    const j = await resp.json();
    if (resp.ok && j.ok) status.innerText = `업로드 완료! 총 ${j.count}건 반영됨`;
    else status.innerText = `업로드 실패: ${j.message || "오류"}`;
  } catch (e) {
    status.innerText = "업로드 실패: 네트워크 오류";
  }
}

async function runSearch() {
  const status = document.getElementById("status");
  const model = document.getElementById("model").value.trim();
  const address = document.getElementById("address").value.trim();
  const owner = document.getElementById("owner").value.trim();
  const nickname = document.getElementById("nickname").value.trim();
  const detail = document.getElementById("detailToggle").checked;

  if (!model) {
    alert("모델(필수)을 입력하세요.");
    return;
  }

  status.innerText = "조회 중...";
  document.getElementById("result").innerHTML = "";

  try {
    const resp = await fetch(`${API_URL}/search`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        center: currentCenter,   // ✅ 센터는 로그인 센터로 강제
        model,
        address: address || "",
        owner: owner || "",
        nickname: nickname || ""
      })
    });

    const j = await resp.json();
    if (!resp.ok || !j.ok) {
      status.innerText = `조회 실패: ${j.message || "오류"}`;
      return;
    }

    status.innerText = `총 ${j.total}대 있습니다.`;

    // 기본 표 컬럼(불필요한 정보 최소화)
    const baseCols = ["보유처", "모델명", "색상", "일련번호"];
    const detailCols = ["상권주소", "펫네임", "애칭"];
    const cols = detail ? [...baseCols, ...detailCols] : baseCols;

    renderTable(j.table, cols);
  } catch (e) {
    status.innerText = "조회 실패: 네트워크 오류";
  }
}

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
