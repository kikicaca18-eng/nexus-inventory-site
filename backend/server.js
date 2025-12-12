// backend/server.js (CommonJS)
const express = require("express");
const cors = require("cors");
const multer = require("multer");
const xlsx = require("xlsx");
const fs = require("fs");

const app = express();
app.use(cors());
app.use(express.json({ limit: "20mb" }));

let inventoryData = []; // 업로드된 엑셀 데이터(최신 1개)

const upload = multer({ dest: "uploads/" });

// 문자열 정규화: 앞뒤 공백 제거 + 연속 공백 축소 + 소문자
function norm(v) {
  return (v ?? "")
    .toString()
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

// 부분일치
function contains(hay, needle) {
  if (!needle) return true;
  return norm(hay).includes(norm(needle));
}

// ✅ 관리자 업로드: 엑셀 업로드하면 "덮어쓰기"로 최신화
app.post("/upload", upload.single("file"), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ ok: false, message: "파일이 없습니다." });

    const wb = xlsx.readFile(req.file.path);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 파일 검증: 헤더가 최소한 존재하는지
    const requiredCols = ["센터", "상권주소", "보유처", "펫네임", "모델명", "색상", "일련번호", "애칭"];
    const sample = rows[0] || {};
    const missing = requiredCols.filter(c => !(c in sample));
    if (missing.length) {
      fs.unlinkSync(req.file.path);
      return res.status(400).json({
        ok: false,
        message: `엑셀 헤더가 다릅니다. 누락: ${missing.join(", ")}`
      });
    }

    // ✅ 최신 파일로 덮어쓰기
    inventoryData = rows.map(r => ({
      센터: r["센터"],
      상권주소: r["상권주소"],
      보유처: r["보유처"],
      펫네임: r["펫네임"],
      모델명: r["모델명"],
      색상: r["색상"],
      일련번호: r["일련번호"],
      애칭: r["애칭"],
    }));

    fs.unlinkSync(req.file.path);
    return res.json({ ok: true, count: inventoryData.length });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "업로드 처리 중 오류" });
  }
});

// ✅ 센터/모델/상권주소/보유처/애칭 검색 (모델 필수)
app.post("/search", (req, res) => {
  try {
    const { center, model, address, owner, nickname } = req.body || {};

    if (!center) return res.status(400).json({ ok: false, message: "center가 필요합니다." });
    if (!model || !model.toString().trim()) {
      return res.status(400).json({ ok: false, message: "모델(model)은 필수입니다." });
    }

    // 1) 센터는 무조건 A열(센터)에서만 필터
    let filtered = inventoryData.filter(r => contains(r["센터"], center));

    // 2) 모델은 모델명/펫네임 컬럼에서만 매칭 (정확도 핵심)
    filtered = filtered.filter(r =>
      contains(r["모델명"], model) || contains(r["펫네임"], model)
    );

    // 3) 선택 조건들
    if (address && address.toString().trim()) {
      filtered = filtered.filter(r => contains(r["상권주소"], address));
    }
    if (owner && owner.toString().trim()) {
      filtered = filtered.filter(r => contains(r["보유처"], owner));
    }
    if (nickname && nickname.toString().trim()) {
      filtered = filtered.filter(r => contains(r["애칭"], nickname));
    }

    // 기본 표에 필요한 정보만 내려줌 (불필요한 컬럼 제거)
    const table = filtered.map(r => ({
      보유처: r["보유처"],
      모델명: r["모델명"],
      색상: r["색상"],
      일련번호: r["일련번호"],
      // 필요 시 프론트에서 “상세 보기”로 보여줄 용도
      상권주소: r["상권주소"],
      펫네임: r["펫네임"],
      애칭: r["애칭"],
      센터: r["센터"],
    }));

    return res.json({ ok: true, total: table.length, table });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "검색 중 오류" });
  }
});

app.get("/meta", (req, res) => {
  return res.json({ ok: true, count: inventoryData.length });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("Backend running on", PORT));
