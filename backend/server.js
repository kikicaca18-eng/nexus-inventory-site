const express = require("express");
const cors = require("cors");
const multer = require("multer");
const XLSX = require("xlsx");
const cron = require("node-cron");
const pool = require("./db"); // PostgreSQL 연결

/**
 * =========================
 *  🔥 서버 버전 식별 로그
 * =========================
 */
console.log("🔥🔥🔥 server.js VERSION 2025-01-19 / PostgreSQL MULTI UPLOAD + SEARCH FIX 🔥🔥🔥");

const app = express();

/**
 * =========================
 *  ✅ CORS (multipart 업로드 대응)
 * =========================
 */
app.use(
  cors({
    origin: "*",
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Content-Type"]
  })
);
app.options("*", cors());

app.use(express.json({ limit: "20mb" }));

/**
 * =========================
 *  루트 테스트
 * =========================
 */
app.get("/", (req, res) => {
  res.send("✅ BACKEND OK - PostgreSQL connected");
});

/**
 * =========================
 *  multer (메모리 업로드)
 * =========================
 */
const upload = multer({
  storage: multer.memoryStorage()
});

/**
 * =========================
 *  엑셀 업로드 (다중 업로드)
 * =========================
 */
app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ ok: false, message: "파일이 없습니다." });
    }

    const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    // 1️⃣ 업로드 파일 기록
    const fileResult = await pool.query(
      `
      INSERT INTO upload_files (filename)
      VALUES ($1)
      RETURNING id
      `,
      [req.file.originalname]
    );

    const uploadFileId = fileResult.rows[0].id;

    // 2️⃣ 재고 데이터 저장 (누적)
    for (const r of rows) {
      await pool.query(
        `
        INSERT INTO inventory (
          upload_file_id,
          센터,
          보유처,
          모델명,
          펫네임,
          애칭,
          색상,
          일련번호,
          상권주소
        ) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9)
        `,
        [
          uploadFileId,
          r.센터,
          r.보유처,
          r.모델명,
          r.펫네임,
          r.애칭,
          r.색상,
          r.일련번호,
          r.상권주소
        ]
      );
    }

    return res.json({
      ok: true,
      count: rows.length,
      uploadFileId,
      message: "엑셀 업로드 완료 (PostgreSQL)"
    });
  } catch (err) {
    console.error("❌ 업로드 오류:", err);
    return res.status(500).json({ ok: false, message: "업로드 실패" });
  }
});

/**
 * =========================
 *  재고 검색 (🔥 핵심 수정 부분)
 * =========================
 */
app.post("/search", async (req, res) => {
  const { center, model, address, owner, nickname } = req.body;

  // 센터는 필수
  if (!center) {
    return res.status(400).json({ ok: false, message: "센터 정보가 없습니다." });
  }

  // 🔥 검색 조건 1개 이상 필수 (중요)
  if (!model && !address && !owner && !nickname) {
    return res.status(400).json({
      ok: false,
      message: "검색 조건을 하나 이상 입력하세요."
    });
  }

  let sql = `
    SELECT *
    FROM inventory
    WHERE upload_file_id = (
      SELECT id
      FROM upload_files
      ORDER BY uploaded_at DESC
      LIMIT 1
    )
    AND 센터 ILIKE $1
  `;

  const params = [`%${center}%`];
  let idx = 2;

  // 모델은 선택
  if (model) {
    sql += ` AND (모델명 ILIKE $${idx} OR 펫네임 ILIKE $${idx})`;
    params.push(`%${model}%`);
    idx++;
  }

  if (address) {
    sql += ` AND 상권주소 ILIKE $${idx}`;
    params.push(`%${address}%`);
    idx++;
  }

  if (owner) {
    sql += ` AND 보유처 ILIKE $${idx}`;
    params.push(`%${owner}%`);
    idx++;
  }

  if (nickname) {
    sql += ` AND 애칭 ILIKE $${idx}`;
    params.push(`%${nickname}%`);
  }

  const result = await pool.query(sql, params);

  res.json({
    ok: true,
    total: result.rows.length,
    table: result.rows
  });
});

/**
 * =========================
 *  마지막 업로드 상태
 * =========================
 */
app.get("/upload-status", async (req, res) => {
  const result = await pool.query(`
    SELECT filename, uploaded_at
    FROM upload_files
    ORDER BY uploaded_at DESC
    LIMIT 1
  `);

  res.json({
    ok: true,
    lastUpload: result.rows[0] || null
  });
});

/**
 * =========================
 *  상태 로그용 CRON
 * =========================
 */
cron.schedule("0 * * * *", async () => {
  try {
    const result = await pool.query(`SELECT COUNT(*) FROM inventory`);
    console.log(`🕐 [CRON] 현재 inventory row 수: ${result.rows[0].count}`);
  } catch (err) {
    console.error("❌ [CRON] 오류:", err);
  }
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
