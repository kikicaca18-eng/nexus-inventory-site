let currentCenter = "";

const passwords = {
  "광주": "yc20405",
  "목포": "yd20001",
  "순천": "yc20404",
  "전북": "yc20407",
  "제주": "yc20403",
  "관리자": "41218673"
};

const API_URL = "https://nexus-inventory-site.onrender.com"; // 백엔드 Render URL

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

function logout() {
  currentCenter = "";
  document.getElementById("loginBox").style.display = "block";
  document.getElementById("searchBox").style.display = "none";
  document.getElementById("uploadBox").style.display = "none";
  document.getElementById("password").value = "";
  document.getElementById("result").innerHTML = "";
}

async function uploadExcel() {
  const fileInput = document.getElementById("excelFile");
  const status = document.getElementById("uploadStatus");
  if (!fileInput.files.length) { status.innerText = "엑셀 파일을 선택해주세요."; return; }
  const formData = new FormData();
  formData.append("file", fileInput.files[0]);
  status.innerText = "업로드 중...";
  try {
    const resp = await fetch(`${API_URL}/upload`, { method:'POST', body: formData });
    const j = await resp.json();
    if (j && j.ok) status.innerText = `업로드 완료! 총 ${j.count || j.rows || '??'}건`;
    else if (resp.ok) status.innerText = `업로드 완료! (${j.count||'?'})`;
    else status.innerText = "업로드 실패";
  } catch (e) {
    status.innerText = "업로드 실패: 네트워크 오류";
  }
}

// 기존 단순 키워드 검색 (빠른 검색)
async function search() {
  const keyword = document.getElementById("keyword").value;
  const resp = await fetch(`${API_URL}/query`, {
    method:'POST',
    headers:{ 'Content-Type':'application/json' },
    body: JSON.stringify({ q: `${currentCenter} ${keyword}` })
  });
  const j = await resp.json();
  renderResponse(j);
}

// 자연어 질문 전송
async function ask() {
  const q = document.getElementById("nlQuery").value;
  if (!q) return alert('질문을 입력하세요.');
  // 자동으로 센터가 포함되지 않으면 현재 센터를 기본으로 붙임 (권한/범위)
  const annotated = currentCenter && !q.includes(currentCenter) && currentCenter !== '관리자'
                    ? `${currentCenter} ${q}` : q;
  const resp = await fetch(`${API_URL}/query`, {
    method:'POST',
    headers:{ 'Content-Type':'application/json' },
    body: JSON.stringify({ q: annotated })
  });
  const j = await resp.json();
  renderResponse(j);
}

// renderResponse: 백엔드 JSON을 보기 좋게 변환
function renderResponse(res) {
  const container = document.getElementById("result");
  container.innerHTML = "";

  if (!res) { container.innerText = "서버 오류"; return; }

  if (res.action === 'count') {
    container.innerHTML = `<p>총 ${res.total}대 있습니다.</p>`;
    return;
  }

  if (res.action === 'model') {
    container.innerHTML = `<p>모델 ${res.model} — 총 ${res.total}대</p>`;
    // center summary
    if (res.centerSummary && res.centerSummary.length) {
      container.innerHTML += `<h4>센터별 요약</h4>${renderTable(res.centerSummary, ['센터','수량'])}`;
    }
    if (res.table) container.innerHTML += `<h4>상세 리스트</h4>${renderTable(res.table, ['센터','보유처','모델명','색상','일련번호'])}`;
    return;
  }

  if (res.action === 'group') {
    container.innerHTML = `<p>총 ${res.total}대 (접점별)</p>`;
    container.innerHTML += `<h4>접점별(보유처/모델/색상/수량)</h4>${renderTable(res.groups, ['보유처','모델명','색상','수량'])}`;
    return;
  }

  if (res.action === 'list' || res.action === 'fallback') {
    container.innerHTML = `<p>총 ${res.total}건</p>`;
    if (res.table && res.table.length) {
      container.innerHTML += renderTable(res.table, ['센터','상권주소','보유처','모델명','색상','일련번호']);
    } else {
      container.innerHTML += `<p>${res.hint||'결과 없음'}</p>`;
    }
    return;
  }

  // 기타
  if (res.hint) container.innerHTML = `<p>${res.hint}</p>`;
  else container.innerHTML = JSON.stringify(res, null, 2);
}

// helper: render array of objects as HTML table using specified columns
function renderTable(rows, cols) {
  if (!rows || rows.length === 0) return '<p>결과 없음</p>';
  let html = '<table><thead><tr>';
  cols.forEach(c => html += `<th>${c}</th>`);
  html += '</tr></thead><tbody>';
  rows.forEach(r => {
    html += '<tr>';
    cols.forEach(c => {
      // support column names that may not exist (fallback '')
      html += `<td>${(r[c]!==undefined ? escapeHtml(r[c]) : '')}</td>`;
    });
    html += '</tr>';
  });
  html += '</tbody></table>';
  return html;
}

function escapeHtml(s){ if(s===null||s===undefined) return ''; return String(s).replace(/[&<>"]/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }
