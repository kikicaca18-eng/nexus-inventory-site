const express = require("express");
const cors = require("cors");
const multer = require("multer");
const XLSX = require("xlsx");
const cron = require("node-cron");
const pool = require("./db"); // ✅ Supabase(Postgres) 연결은 db.js에서 처리

console.log("🔥 server.js VERSION 2026-01-27 / SUPABASE inventory_items SNAPSHOT + SEARCH 🔥");

const app = express();

app.use(
  cors({
    origin: "*",
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Content-Type"]
  })
);
app.options("*", cors());

app.use(express.json({ limit: "20mb" }));

app.get("/", (req, res) => {
  res.send("✅ BACKEND OK - Supabase PostgreSQL connected");
});

const upload = multer({ storage: multer.memoryStorage() });

/**
 * 한국시간(Asia/Seoul) 기준 YYYY-MM-DD
 */
function todayKST() {
  return new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Seoul" }).format(new Date());
}

function toText(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

/**
 * =========================
 * ✅ 업로드: /upload  (기존 프론트 유지)
 * - 새 재고 포맷(0127) 업로드
 * - 오늘 날짜 스냅샷으로 "덮어쓰기"
 * - 저장 테이블: inventory_items
 * =========================
 */
app.post("/upload", upload.single("file"), async (req, res) => {
  const snapshotDate = todayKST();

  try {
    if (!req.file) {
      return res.status(400).json({ ok: false, message: "파일이 없습니다." });
    }

    const workbook = XLSX.read(req.file.buffer, { type: "buffer" });

    // 시트 1개만 운영한다고 했으니 첫 시트 사용
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    if (!rows.length) {
      return res.status(400).json({ ok: false, message: "시트에 데이터가 없습니다." });
    }

    // ✅ 기대 컬럼(너가 올린 재고 파일 기준)
    const expected = [
      "대리점코드",
      "대리점명",
      "세부상권",
      "상세주소",
      "접점코드",
      "접점명",
      "펫네임",
      "모델명",
      "색상",
      "일련번호",
      "애칭"
    ];

    for (const c of expected) {
      if (!Object.prototype.hasOwnProperty.call(rows[0], c)) {
        return res.status(400).json({
          ok: false,
          message: `엑셀 헤더에 '${c}' 컬럼이 없습니다. (첫줄 헤더 확인)`
        });
      }
    }

    // ✅ 데이터 가공(일련번호 없는 줄 제외)
    const data = rows
      .filter(r => toText(r["일련번호"]) !== "")
      .map(r => ({
        agency_code: toText(r["대리점코드"]),
        agency_name: toText(r["대리점명"]),
        sub_market: toText(r["세부상권"]),
        address: toText(r["상세주소"]),
        store_code: toText(r["접점코드"]),
        store_name: toText(r["접점명"]),
        pet_name: toText(r["펫네임"]),
        model_name: toText(r["모델명"]),
        color: toText(r["색상"]),
        serial_no: toText(r["일련번호"]),
        nickname: toText(r["애칭"])
      }));

    const client = await pool.connect();
    try {
      await client.query("BEGIN");

      // ✅ 오늘 스냅샷 덮어쓰기(안전/단순/왕초보 운영 최적)
      await client.query("DELETE FROM inventory_items WHERE snapshot_date = $1", [snapshotDate]);

      // ✅ 대량 insert(빠르게)
      // 한 번에 너무 크게 넣지 않도록 800개씩 끊어서 넣음
      const chunkSize = 800;

      for (let i = 0; i < data.length; i += chunkSize) {
        const chunk = data.slice(i, i + chunkSize);

        const values = [];
        const placeholders = chunk
          .map((row, idx) => {
            const base = idx * 12;
            values.push(
              snapshotDate,
              row.agency_code,
              row.agency_name,
              row.sub_market,
              row.address,
              row.store_code,
              row.store_name,
              row.pet_name,
              row.model_name,
              row.color,
              row.serial_no,
              row.nickname
            );
            // 12 columns
            return `($${base + 1},$${base + 2},$${base + 3},$${base + 4},$${base + 5},$${base + 6},$${base + 7},$${base + 8},$${base + 9},$${base + 10},$${base + 11},$${base + 12})`;
          })
          .join(",");

        const sql = `
          INSERT INTO inventory_items
          (snapshot_date, agency_code, agency_name, sub_market, address, store_code, store_name, pet_name, model_name, color, serial_no, nickname)
          VALUES ${placeholders}
        `;

        await client.query(sql, values);
      }

      await client.query("COMMIT");

      return res.json({
        ok: true,
        count: data.length,
        snapshot_date: snapshotDate,
        sheet: sheetName,
        message: "재고 업로드 완료 (Supabase snapshot overwrite)"
      });
    } catch (e) {
      await client.query("ROLLBACK");
      console.error("❌ 업로드 DB 오류:", e);
      return res.status(500).json({ ok: false, message: "업로드 실패(DB)" });
    } finally {
      client.release();
    }
  } catch (err) {
    console.error("❌ 업로드 오류:", err);
    return res.status(500).json({ ok: false, message: "업로드 실패" });
  }
});

/**
 * =========================
 * ✅ 검색: /search  (기존 프론트 유지)
 * - 기존 파라미터 그대로 받음: center(필수), model/address/owner/nickname(선택)
 * - 오늘 스냅샷(snapshot_date=오늘) 기준 검색
 *
 * 매핑:
 *  center   -> store_name(접점명)에서 검색 (필수)
 *  model    -> model_name 또는 pet_name 검색
 *  address  -> address(상세주소) 검색
 *  owner    -> agency_name(대리점명) 검색
 *  nickname -> nickname(애칭) 검색
 * =========================
 */
app.post("/search", async (req, res) => {
  const snapshotDate = todayKST();
  const { center, model, address, owner, nickname } = req.body;

  if (!center) {
    return res.status(400).json({ ok: false, message: "센터(center) 정보가 없습니다." });
  }

  if (!model && !address && !owner && !nickname) {
    return res.status(400).json({
      ok: false,
      message: "검색 조건을 하나 이상 입력하세요."
    });
  }

  let sql = `
    SELECT
      snapshot_date,
      agency_code, agency_name,
      sub_market, address,
      store_code, store_name,
      pet_name, model_name, color,
      serial_no, nickname
    FROM inventory_items
    WHERE snapshot_date = $1
      AND store_name ILIKE $2
  `;

  const params = [snapshotDate, `%${center}%`];
  let idx = 3;

  if (model) {
    sql += ` AND (model_name ILIKE $${idx} OR pet_name ILIKE $${idx})`;
    params.push(`%${model}%`);
    idx++;
  }

  if (address) {
    sql += ` AND address ILIKE $${idx}`;
    params.push(`%${address}%`);
    idx++;
  }

  if (owner) {
    sql += ` AND agency_name ILIKE $${idx}`;
    params.push(`%${owner}%`);
    idx++;
  }

  if (nickname) {
    sql += ` AND nickname ILIKE $${idx}`;
    params.push(`%${nickname}%`);
    idx++;
  }

  // 너무 많이 뿌리면 프론트가 느려질 수 있어서 제한
  sql += ` ORDER BY store_name ASC, model_name ASC LIMIT 2000`;

  try {
    const result = await pool.query(sql, params);
    return res.json({
      ok: true,
      snapshot_date: snapshotDate,
      total: result.rows.length,
      table: result.rows
    });
  } catch (err) {
    console.error("❌ 검색 오류:", err);
    return res.status(500).json({ ok: false, message: "검색 실패" });
  }
});

/**
 * =========================
 * ✅ 마지막 업로드 상태: /upload-status (기존 프론트 유지)
 * - "오늘 스냅샷에 데이터가 있냐"로 판단
 * =========================
 */
app.get("/upload-status", async (req, res) => {
  const snapshotDate = todayKST();

  try {
    const r = await pool.query(
      `SELECT COUNT(*)::int AS cnt FROM inventory_items WHERE snapshot_date = $1`,
      [snapshotDate]
    );

    res.json({
      ok: true,
      snapshot_date: snapshotDate,
      today_count: r.rows[0]?.cnt ?? 0
    });
  } catch (err) {
    console.error("❌ 업로드 상태 오류:", err);
    res.status(500).json({ ok: false, message: "상태 조회 실패" });
  }
});

/**
 * =========================
 *  상태 로그용 CRON (선택)
 * =========================
 */
cron.schedule("0 * * * *", async () => {
  const snapshotDate = todayKST();
  try {
    const result = await pool.query(
      `SELECT COUNT(*)::int AS cnt FROM inventory_items WHERE snapshot_date = $1`,
      [snapshotDate]
    );
    console.log(`🕐 [CRON] 오늘(${snapshotDate}) inventory_items row 수: ${result.rows[0].cnt}`);
  } catch (err) {
    console.error("❌ [CRON] 오류:", err);
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("🚀 Backend running on port", PORT);
});
