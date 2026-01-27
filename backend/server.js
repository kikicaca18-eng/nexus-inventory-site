const express = require("express");
const cors = require("cors");
const multer = require("multer");
const XLSX = require("xlsx");
const cron = require("node-cron");
const pool = require("./db");

console.log("🔥 server.js FINAL / SUPABASE inventory_items + AGENCY AUTH 🔥");

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
 * 한국시간 기준 날짜
 */
function todayKST() {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Seoul"
  }).format(new Date());
}

function toText(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

/**
 * =========================
 * ✅ 업로드: /upload (관리자만)
 * =========================
 */
app.post("/upload", upload.single("file"), async (req, res) => {
  const snapshotDate = todayKST();

  try {
    if (!req.file) {
      return res.status(400).json({ ok: false, message: "파일이 없습니다." });
    }

    const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    if (!rows.length) {
      return res.status(400).json({ ok: false, message: "시트에 데이터가 없습니다." });
    }

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
          message: `엑셀 헤더에 '${c}' 컬럼이 없습니다.`
        });
      }
    }

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

      await client.query(
        "DELETE FROM inventory_items WHERE snapshot_date = $1",
        [snapshotDate]
      );

      const chunkSize = 800;

      for (let i = 0; i < data.length; i += chunkSize) {
        const chunk = data.slice(i, i + chunkSize);
        const values = [];
        const placeholders = chunk.map((row, idx) => {
          const b = idx * 12;
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
          return `($${b + 1},$${b + 2},$${b + 3},$${b + 4},$${b + 5},$${b + 6},$${b + 7},$${b + 8},$${b + 9},$${b + 10},$${b + 11},$${b + 12})`;
        });

        await client.query(
          `
          INSERT INTO inventory_items
          (snapshot_date, agency_code, agency_name, sub_market, address, store_code, store_name, pet_name, model_name, color, serial_no, nickname)
          VALUES ${placeholders.join(",")}
          `,
          values
        );
      }

      await client.query("COMMIT");

      res.json({
        ok: true,
        count: data.length,
        snapshot_date: snapshotDate
      });
    } catch (e) {
      await client.query("ROLLBACK");
      console.error(e);
      res.status(500).json({ ok: false, message: "DB 오류" });
    } finally {
      client.release();
    }
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, message: "업로드 실패" });
  }
});

/**
 * =========================
 * ✅ 재고 검색: /search
 * - agency 기준 강제 제한
 * =========================
 */
app.post("/search", async (req, res) => {
  const snapshotDate = todayKST();
  const { agency, model, address, owner, nickname } = req.body;

  if (!agency) {
    return res.status(400).json({ ok: false, message: "agency 정보가 없습니다." });
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
      agency_name,
      store_name,
      model_name,
      pet_name,
      color,
      serial_no,
      address,
      nickname
    FROM inventory_items
    WHERE snapshot_date = $1
  `;

  const params = [snapshotDate];
  let idx = 2;

  // 🔐 관리자 아니면 본인 대리점만
  if (agency !== "관리자") {
    sql += ` AND agency_name = $${idx}`;
    params.push(agency);
    idx++;
  }

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
    sql += ` AND store_name ILIKE $${idx}`;
    params.push(`%${owner}%`);
    idx++;
  }

  if (nickname) {
    sql += ` AND nickname ILIKE $${idx}`;
    params.push(`%${nickname}%`);
    idx++;
  }

  sql += ` ORDER BY store_name ASC, model_name ASC LIMIT 2000`;

  try {
    const result = await pool.query(sql, params);
    res.json({
      ok: true,
      agency,
      snapshot_date: snapshotDate,
      total: result.rows.length,
      table: result.rows
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, message: "검색 실패" });
  }
});

/**
 * =========================
 * 업로드 상태
 * =========================
 */
app.get("/upload-status", async (req, res) => {
  const snapshotDate = todayKST();
  const r = await pool.query(
    "SELECT COUNT(*)::int AS cnt FROM inventory_items WHERE snapshot_date = $1",
    [snapshotDate]
  );

  res.json({
    ok: true,
    snapshot_date: snapshotDate,
    today_count: r.rows[0]?.cnt ?? 0
  });
});

/**
 * =========================
 * CRON 로그
 * =========================
 */
cron.schedule("0 * * * *", async () => {
  const d = todayKST();
  const r = await pool.query(
    "SELECT COUNT(*)::int AS cnt FROM inventory_items WHERE snapshot_date = $1",
    [d]
  );
  console.log(`🕐 [CRON] ${d} 재고 ${r.rows[0].cnt}건`);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("🚀 Backend running on port", PORT);
});
