const express = require("express");
const cors = require("cors");
const multer = require("multer");
const XLSX = require("xlsx");

const app = express();
app.use(cors());
app.use(express.json({ limit: "20mb" }));

/**
 * =========================
 *  메모리 기반 재고 데이터
 * =========================
 */
let inventoryData = [];

/**
 * =========================
 *  문자열 유틸
 * =========================
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
 * =========================
 *  multer (메모리 저장)
 * =========================
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

    const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    inventoryData = rows;

    return res.json({
      ok: true,
      count: inventoryData.length,
      message: "엑셀 업로드 및 파싱 완료"
    });
  } catch (e) {
    console.error("❌ 업로드 오류:", e);
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

  // ✅ 센터는 필수
  if (!center) {
    return res.status(400).json({
      ok: false,
      message: "센터 정보가 없습니다."
    });
  }

  // 🔥 핵심: 검색 조건 최소 1개 필수
  if (!model && !address && !owner && !nickname) {
    return res.status(400).json({
      ok: false,
      message: "검색 조건을 하나 이상 입력하세요."
    });
  }

  // 1️⃣ 센터 필터
  let filtered = inventoryData.filter(r =>
    contains(r.센터, center)
  );

  // 2️⃣ 모델 (있을 때만)
  if (model) {
    filtered = filtered.filter(r =>
      contains(r.모델명, model) || contains(r.펫네임, model)
    );
  }

  // 3️⃣ 기타 조건
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
