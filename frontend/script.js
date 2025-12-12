let currentCenter = "";

// 센터 비밀번호 (원하면 나중에 바꿀 수 있음)
const passwords = {
  "광주": "1111",
  "목포": "2222",
  "순천": "3333",
  "전북": "4444",
  "제주": "5555"
};

function login() {
  const center = document.getElementById("centerSelect").value;
  const pw = document.getElementById("password").value;

  if (!center) return alert("센터를 선택하세요.");
  if (pw !== passwords[center]) return alert("비밀번호가 틀렸습니다.");

  currentCenter = center;

  document.getElementById("loginBox").style.display = "none";
  document.getElementById("searchBox").style.display = "block";

  document.getElementById("centerTitle").innerText = `[ ${center} 센터 ]`;
}

async function search() {
  const keyword = document.getElementById("keyword").value;

  const response = await fetch("http://localhost:3000/search", {
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
    item.innerHTML = JSON.stringify(row);
    resultDiv.appendChild(item);
  });

  if (data.results.length === 0) {
    resultDiv.innerHTML = "<p>검색 결과 없음</p>";
  }
}
