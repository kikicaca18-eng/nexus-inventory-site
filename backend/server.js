const path = require("path");
const express = require("express");
const cors = require("cors");
const multer = require("multer");
const XLSX = require("xlsx");
const Database = require("better-sqlite3");
const cron = require("node-cron");

const app = express();
app.use(cors());
app.use(express.json({ limit: "20mb" }));

/**
 * =========================
 *  SQLite DB 연결
 * =========================
 */
const dbPath = path.join(__dirname, "inventory.db");
const db = new Database(dbPath);
console.log("📦 SQLite DB connected:", dbPath);

/**
 * =========================
 *  테이블 생성
 * =========================
 */
db.prepare(`
  CREATE TABLE IF NOT EXISTS inventory (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    센터 TEXT,
    보유처 TEXT,
    모델명 TEXT,
    펫네임 TEXT,
    애칭 TEXT,
    색상 TEXT,
    일련번호 TEXT,
    상권주소 TEXT
  )
`).run();

db.prepare(`
  CREATE TABLE IF NOT EXISTS upload_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    uploaded_at_utc TEXT NOT NULL,
    uploaded_at_kst TEXT NOT NULL,
    row_count INTEGER NOT NULL,
    filename TEXT
  )
`).run();

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
 *  엑셀 업로드 (전체 덮어쓰기)
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

    // 기존 재고 삭제
    db.prepare("DELETE FROM inventory").run();

    const insert = db.prepare(`
      INSERT INTO inventory (
        센터, 보유처, 모델명, 펫네임, 애칭, 색상, 일련번호, 상권주소
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    `);

    const tx = db.transaction(data => {
      for (const r of data) {
        insert.run(
          r.센터,
          r.보유처,
          r.모델명,
          r.펫네임,
          r.애칭,
          r.색상,
          r.일련번호,
          r.상권주소
        );
      }
    });

    tx(rows);

    // 업로드 로그 저장
    const now = new Date();
    const uploadedAtUtc = now.toISOString();
    const uploadedAtKst = new Intl.DateTimeFormat("ko-KR", {
      timeZone: "Asia/Seoul",
      dateStyle: "medium",
      timeStyle: "medium"
    }).format(now);

    db.prepare(`
      INSERT INTO upload_log (uploaded_at_utc, uploaded_at_kst, row_count, filename)
      VALUES (?, ?, ?, ?)
    `).run(
      uploadedAtUtc,
      uploadedAtKst,
      rows.length,
      req.file.originalname || null
    );

    return res.json({
      ok: true,
      count: rows.length,
      message: "엑셀 업로드 완료"
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

  if (!center) {
    return res.status(400).json({ ok: false, message: "센터 정보가 없습니다." });
  }

  if (!model && !address && !owner && !nickname) {
    return res.status(400).json({
      ok: false,
      message: "검색 조건을 하나 이상 입력하세요."
    });
  }

  let sql = `SELECT * FROM inventory WHERE 센터 LIKE ?`;
  let params = [`%${center}%`];

  if (model) {
    sql += ` AND (모델명 LIKE ? OR 펫네임 LIKE ?)`;
    params.push(`%${model}%`, `%${model}%`);
  }
  if (address) {
    sql += ` AND 상권주소 LIKE ?`;
    params.push(`%${address}%`);
  }
  if (owner) {
    sql += ` AND 보유처 LIKE ?`;
    params.push(`%${owner}%`);
  }
  if (nickname) {
    sql += ` AND 애칭 LIKE ?`;
    params.push(`%${nickname}%`);
  }

  const rows = db.prepare(sql).all(...params);

  res.json({
    ok: true,
    total: rows.length,
    table: rows
  });
});

/**
 * =========================
 *  마지막 업로드 상태
 * =========================
 */
app.get("/upload-status", (req, res) => {
  const last = db.prepare(`
    SELECT uploaded_at_kst, row_count, filename
    FROM upload_log
    ORDER BY id DESC
    LIMIT 1
  `).get();

  res.json({
    ok: true,
    lastUpload: last || null
  });
});

/**
 * =========================
 *  1시간마다 SQLite 데이터 리프레시
 * =========================
 * - Render 재시작과 무관
 * - 엑셀 재업로드 ❌
 * - 현재 DB 상태 점검용
 */

cron.schedule("0 * * * *", () => {
  try {
    console.log("🕐 [CRON] 1시간 주기 SQLite 상태 점검 시작");

    // 현재 재고 row 수 확인
    const row = db
      .prepare("SELECT COUNT(*) as cnt FROM inventory")
      .get();

    console.log(`📦 [CRON] 현재 inventory row 수: ${row.cnt}`);

    // 마지막 업로드 기록 확인
    const lastUpload = db
      .prepare(`
        SELECT uploaded_at_kst, row_count
        FROM upload_log
        ORDER BY id DESC
        LIMIT 1
      `)
      .get();

    if (!lastUpload) {
      console.warn("⚠️ [CRON] 업로드 이력이 없습니다.");
    } else {
      console.log(
        `📄 [CRON] 마지막 업로드: ${lastUpload.uploaded_at_kst} / ${lastUpload.row_count}건`
      );
    }

    console.log("✅ [CRON] SQLite 리프레시 완료");
  } catch (err) {
    console.error("❌ [CRON] SQLite 리프레시 오류:", err);
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
