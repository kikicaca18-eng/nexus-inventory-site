let currentCenter = "";

// 센터 비밀번호
const passwords = {
  "광주": "yc20405",
  "목포": "yd20001",
  "순천": "yc20404",
  "전북": "yc20407",
  "제주": "yc20403",
  "관리자": "41218673"  // 엑셀 업로드 관리자
};

// 백엔드 API 주소
const API_URL = "https://nexus-inventory-site.onrender.com";

function login() {
  const center = document.getElementById("centerSelect").value;
  const pw = document.getElementById("password").value;

  if (!center) return alert("센터를 선택하세요.");
  if (pw !== passwords[center]) return alert("비밀번호가 틀렸습니다.");

  currentCenter = center;

  document.getElementById("loginBox").style.display = "none";

  if (center === "관리자") {
      document.getElementById("uploadBox").style.display = "block";
  } else {
      document.getElementById("searchBox").style.display = "block";
      document.getElementById("centerTitle").innerText = `[ ${center} 센터 ]`;
  }
}

async function search() {
  const keyword = document.getElementById("keyword").value;

  const response = await fetch(API_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ center: currentCenter, keyword })
  });

  const data = await response.json();

  const resultDiv = document.getElementById("result");
  resultDiv.innerHTML = "";

  data.results.forEach(row => {
    const item = document.createElement("div");
    item.className = "item";
    item.innerHTML = Object.entries(row)
      .map(([k, v]) => `<b>${k}</b>: ${v}`)
      .join("<br>");
    resultDiv.appendChild(item);
  });

  if (data.results.length === 0) {
    resultDiv.innerHTML = "<p>검색 결과 없음</p>";
  }
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

    status.innerText = "업로드 중...";

    try {
        const response = await fetch(`${API_URL}/upload`, {
            method: "POST",
            body: formData
        });

        if (response.ok) {
            status.innerText = "업로드 완료!";
        } else {
            status.innerText = "업로드 실패: 서버 오류";
        }
    } catch (error) {
        status.innerText = "업로드 실패: 네트워크 오류";
    }
}

function logout() {
    currentCenter = "";
    document.getElementById("loginBox").style.display = "block";
    document.getElementById("searchBox").style.display = "none";
    document.getElementById("uploadBox").style.display = "none";
    document.getElementById("password").value = "";
}
