// backend/server.js (CommonJS)
// 덮어쓰기: 기존 server.js를 아래로 교체하세요.
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

let inventoryData = []; // 업로드된 엑셀의 각 행(객체)이 저장됨

const upload = multer({ dest: 'uploads/' });

// 유틸: 안전한 문자열 (소문자)
function norm(s){ return (s || '').toString().trim().toLowerCase(); }

// 엑셀 업로드: /upload (multipart form-data, field name = file)
app.post('/upload', upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ ok:false, message:'파일 필요' });
    const workbook = xlsx.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = xlsx.utils.sheet_to_json(sheet, { defval: '' });

    // 컬럼 정규화(간단): ensure expected columns exist; keep original keys
    inventoryData = json.map(row => {
      return {
        센터: row['센터'] || row['center'] || '',
        상권주소: row['상권주소'] || row['상권'] || '',
        보유처: row['보유처'] || row['보유처명'] || row['판매점'] || '',
        펫네임: row['펫네임'] || row['펫 이름'] || row['펫네임'] || '',
        모델명: row['모델명'] || row['모델'] || '',
        색상: row['색상'] || row['컬러'] || '',
        일련번호: row['일련번호'] || row['시리얼'] || row['serial'] || '',
        // 원본 전체도 남겨둠
        __raw: row
      };
    });

    // remove temp file
    fs.unlinkSync(req.file.path);
    return res.json({ ok:true, count: inventoryData.length });
  } catch (e) {
    console.error('upload error', e);
    return res.status(500).json({ ok:false, message:'업로드중 오류' });
  }
});

// /meta : 간단 통계, 센터 목록
app.get('/meta', (req,res) => {
  const centers = Array.from(new Set(inventoryData.map(r => r.센터).filter(x=>x))).sort();
  res.json({ count: inventoryData.length, centers });
});

/**
 * 핵심: /query
 * body: { q: string } - 사용자가 입력한 자연어(또는 간단 키워드)
 * 응답 형식:
 *  { action: 'count'|'list'|'group'|'none',
 *    summary: '문장',
 *    count: number,
 *    table: [ {센터,상권주소,보유처,모델명,색상,일련번호} ],
 *    groups: [ { key: '보유처A', count: n, model:'M366', color:'블랙' } ]
 *  }
 */
app.post('/query', (req,res) => {
  try {
    const q = (req.body.q || '').toString().trim();
    if (!q) return res.json({ action:'none', message:'질문을 입력하세요.' });

    const ql = q.toLowerCase();

    // helper: filter by partial match over specified columns
    function partialFilter(data, token, columns){
      if (!token) return data;
      const t = token.toLowerCase();
      return data.filter(row => {
        return columns.some(col => norm(row[col]).includes(t));
      });
    }

    // detect center name (광주/목포/순천/전북/제주) in query
    const centersPossible = ['광주','목포','순천','전북','제주'];
    let centerFilter = null;
    for (const c of centersPossible){
      if (ql.includes(c)) { centerFilter = c; break; }
      // also allow '제주센터' etc:
      if (ql.includes(c + '센터')) { centerFilter = c; break; }
    }

    // detect "일련번호" request with numeric token
    let serialToken = null;
    const serialMatch = q.match(/(일련번호|serial|시리얼)[: ]*([0-9A-Za-z\-]+)/i);
    if (serialMatch) serialToken = serialMatch[2];
    // also detect "일련번호 355" style
    if (!serialToken) {
      const m = q.match(/일련번호\s*([0-9A-Za-z\-]+)/i);
      if (m) serialToken = m[1];
    }

    // detect "리스트", "보여줘", "어디 있어" -> list
    const wantsList = /리스트|보여줘|어디 있어|어디있어|상세|자세히|목록/i.test(q);
    // detect "몇대" or "몇 대" or "대 있어" -> count
    const wantsCount = /몇\s*대|대\s*있어|총\s*\d+\s*대|수량/i.test(q);
    // detect "접점별" or "판매점별" or "보유처별" -> group by 보유처
    const wantsGroupByOwner = /접점별|보유처별|판매점별|접점/i.test(q);
    // detect model token like 'M366' or names in query
    const tokenModelMatch = q.match(/([A-Za-z]{1,2}[0-9]{1,4}|[A-Za-z0-9\-]{2,10})/g);
    // try to identify likely model tokens by matching against data
    let modelToken = null;
    if (tokenModelMatch && inventoryData.length>0){
      // pick tokens that appear in 모델명 or 펫네임
      for (const tk of tokenModelMatch){
        const tkl = tk.toLowerCase();
        const found = inventoryData.some(r => norm(r.모델명).includes(tkl) || norm(r.펫네임).includes(tkl));
        if (found && tk.length>=2) { modelToken = tk; break; }
      }
    }

    // detect address place names (simple): if query contains a token that matches any 상권주소 substring
    let areaToken = null;
    if (inventoryData.length>0){
      // collect unique place tokens (split 상권주소 by space/comma)
      const addrSamples = inventoryData.slice(0,1000).map(r => r.상권주소 || '').join(' ').toLowerCase();
      // check if any known city/군/구 substring appears in query (simple heuristic)
      const placeCandidates = ['군산','목포','광주','순천','제주','전북','전주','송정','남구','동구','서구'];
      for (const pc of placeCandidates){
        if (ql.includes(pc)) { areaToken = pc; break; }
      }
    }

    // Start filtering pipeline
    let filtered = inventoryData.slice();
    if (centerFilter) filtered = filtered.filter(r => norm(r.센터).includes(norm(centerFilter)));
    // if serial token -> restrict to 일련번호
    if (serialToken) filtered = filtered.filter(r => norm(r.일련번호).includes(norm(serialToken)));
    // if modelToken -> restrict to 모델명 or 펫네임
    if (modelToken) filtered = filtered.filter(r =>
      norm(r.모델명).includes(norm(modelToken)) || norm(r.펫네임).includes(norm(modelToken))
    );
    // if areaToken -> restrict to 상권주소
    if (areaToken) filtered = filtered.filter(r => norm(r.상권주소).includes(norm(areaToken)));

    // if query includes a specific 보유처 name (heuristic: check tokens)
    // look for any token that matches 보유처 values
    const owners = Array.from(new Set(inventoryData.map(r=>r.보유처).filter(x=>x))).slice(0,1000);
    let ownerToken = null;
    for (const o of owners){
      if (!o) continue;
      if (ql.includes(o.toLowerCase())) { ownerToken = o; break; }
    }
    if (ownerToken) filtered = filtered.filter(r => norm(r.보유처).includes(norm(ownerToken)));

    // If nothing filtered but the user asked a model-like question with explicit model in text, try broad search across columns
    if (filtered.length === 0 && modelToken) {
      filtered = inventoryData.filter(r => Object.values(r).some(v => norm(v).includes(norm(modelToken))));
    }

    // Decide response type
    // If user asked for group by owner
    if (wantsGroupByOwner) {
      // group by 보유처 / 모델명 / 색상 -> count
      const map = {};
      filtered.forEach(r => {
        const key = `${r.보유처}||${r.모델명}||${r.색상}`;
        map[key] = (map[key]||0)+1;
      });
      const groups = Object.entries(map).map(([k,v]) => {
        const [보유처, 모델명, 색상] = k.split('||');
        return { 보유처, 모델명, 색상, 수량: v };
      }).sort((a,b)=>b.수량 - a.수량);
      return res.json({ action:'group', groups, total: filtered.length });
    }

    // If user asked for list
    if (wantsList) {
      const table = filtered.map(r => ({
        센터: r.센터, 상권주소: r.상권주소, 보유처: r.보유처,
        모델명: r.모델명, 색상: r.색상, 일련번호: r.일련번호
      }));
      return res.json({ action:'list', total: table.length, table });
    }

    // If user asked for count (몇대)
    if (wantsCount || /\b몇\b/.test(ql) && /\b대\b/.test(ql)) {
      return res.json({ action:'count', total: filtered.length });
    }

    // Default fallback: if query includes '일련번호' -> list detail
    if (serialToken) {
      const table = filtered.map(r => ({
        센터: r.센터, 상권주소: r.상권주소, 보유처: r.보유처,
        모델명: r.모델명, 색상: r.색상, 일련번호: r.일련번호
      }));
      return res.json({ action:'list', total: table.length, table });
    }

    // If user asked for model-specific (e.g., 'M366')
    if (modelToken) {
      // default: show count + groups by center and owner
      // produce center summary
      const centerMap = {};
      filtered.forEach(r => { centerMap[r.센터] = (centerMap[r.센터]||0)+1; });
      const centerSummary = Object.entries(centerMap).map(([k,v]) => ({ 센터:k, 수량:v }));
      const table = filtered.map(r=>({센터:r.센터, 보유처:r.보유처, 모델명:r.모델명, 색상:r.색상, 일련번호:r.일련번호}));
      return res.json({ action:'model', model:modelToken, total: filtered.length, centerSummary, table });
    }

    // final fallback: return small table + hint
    const table = filtered.slice(0,200).map(r => ({
      센터: r.센터, 상권주소: r.상권주소, 보유처: r.보유처,
      모델명: r.모델명, 색상: r.색상, 일련번호: r.일련번호
    }));

    const hint = '질문을 구체적으로 해보세요. 예: "M366 재고 몇대", "광주에 재고 몇대", "리스트 보여줘"';
    return res.json({ action:'fallback', total: table.length, table, hint });
  } catch (e) {
    console.error('/query error', e);
    return res.status(500).json({ ok:false, message:'서버 오류' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, ()=>console.log('Server run on', PORT));
