const express = require("express");
const cors = require("cors");
const multer = require("multer");
const XLSX = require("xlsx");

const app = express();
app.use(cors());
app.use(express.json({ limit: "20mb" }));

/**
 * 🔹 메모리 기반 재고 데이터
 * 서버 재시작 시 초기화됨
 */
let inventoryData = [];

/**
 * 🔹 문자열 정규화 유틸
 */
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

/**
 * 🔹 multer (메모리 저장)
 */
const upload = multer({
  storage: multer.memoryStorage()
});

/**
 * =========================
 *  엑셀 업로드 & 파싱
 * =========================
 */
app.post("/upload", upload.single("file"), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ ok: false, message: "파일이 없습니다." });
    }

    // 1. 메모리에서 엑셀 읽기
    const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // 2. 파싱 (엑셀 → JSON)
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    // 3. 메모리 DB 교체
    inventoryData = rows;

    return res.json({
      ok: true,
      count: inventoryData.length,
      message: "엑셀 업로드 및 파싱 완료"
    });

  } catch (e) {
    console.error("❌ 엑셀 업로드 오류:", e);
    return res.status(500).json({ ok: false, message: "업로드 실패" });
  }
});

/**
 * =========================
 *  재고 검색 API
 * =========================
 */
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
    애칭: r.애칭
  }));

  return res.json({
    ok: true,
    total: table.length,
    table
  });
});

/**
 * =========================
 *  서버 상태 확인
 * =========================
 */
app.get("/meta", (req, res) => {
  res.json({
    ok: true,
    count: inventoryData.length
  });
});

/**
 * =========================
 *  서버 시작
 * =========================
 */
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("🚀 Backend running on port", PORT);
});
