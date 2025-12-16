const express = require("express");
const cors = require("cors");
const multer = require("multer");
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();
app.use(cors());
app.use(express.json({ limit: "20mb" }));

// 📌 데이터 저장 경로 (서버 디스크)
const DATA_DIR = path.join(__dirname, "data");
const EXCEL_PATH = path.join(DATA_DIR, "latest.xlsx");

let inventoryData = [];

// ===== 유틸 =====
function norm(v) {
  return (v ?? "")
    .toString()
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function contains(hay, needle) {
  if (!needle) return true;
  return norm(hay).includes(norm(needle));
}

// ===== 서버 시작 시: 기존 엑셀 자동 로딩 =====
function loadExcelFromDisk() {
  try {
    if (!fs.existsSync(EXCEL_PATH)) {
      console.log("ℹ️ 저장된 엑셀 없음");
      return;
    }

    const wb = xlsx.readFile(EXCEL_PATH);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

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

    console.log(`✅ 엑셀 자동 로딩 완료 (${inventoryData.length}건)`);
  } catch (e) {
    console.error("❌ 엑셀 로딩 실패:", e);
  }
}

// 서버 시작 시 실행
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}
loadExcelFromDisk();

// ===== 파일 업로드 설정 =====
const upload = multer({ dest: "uploads/" });

// ===== 엑셀 업로드 (덮어쓰기 저장) =====
app.post("/upload", upload.single("file"), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ ok: false, message: "파일이 없습니다." });
    }

    // uploads 임시파일 → data/latest.xlsx 로 이동
    fs.copyFileSync(req.file.path, EXCEL_PATH);
    fs.unlinkSync(req.file.path);

    // 새 엑셀 로딩
    loadExcelFromDisk();

    return res.json({
      ok: true,
      count: inventoryData.length,
      message: "엑셀 업로드 및 저장 완료"
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "업로드 실패" });
  }
});

// ===== 재고 검색 API =====
app.post("/search", (req, res) => {
  const { center, model, address, owner, nickname } = req.body;

  if (!center) {
    return res.status(400).json({ ok: false, message: "센터 정보가 없습니다." });
  }
  if (!model) {
    return res.status(400).json({ ok: false, message: "모델은 필수입니다." });
  }

  let filtered = inventoryData.filter(r => contains(r.센터, center));

  filtered = filtered.filter(r =>
    contains(r.모델명, model) || contains(r.펫네임, model)
  );

  if (address) filtered = filtered.filter(r => contains(r.상권주소, address));
  if (owner) filtered = filtered.filter(r => contains(r.보유처, owner));
  if (nickname) filtered = filtered.filter(r => contains(r.애칭, nickname));

  const table = filtered.map(r => ({
    보유처: r.보유처,
    모델명: r.모델명,
    색상: r.색상,
    일련번호: r.일련번호,
    상권주소: r.상권주소,
    펫네임: r.펫네임,
    애칭: r.애칭,
  }));

  res.json({ ok: true, total: table.length, table });
});

// ===== 상태 확인 =====
app.get("/meta", (req, res) => {
  res.json({
    ok: true,
    count: inventoryData.length,
    saved: fs.existsSync(EXCEL_PATH)
  });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("🚀 Backend running on port", PORT);
});
