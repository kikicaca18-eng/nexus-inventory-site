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
 * ✅ 재고 대시보드: 요약 카드
 * POST /inventory/summary
 * body: { agency }
 * =========================
 */
app.post("/inventory/summary", async (req, res) => {
  const snapshotDate = todayKST();
  const { agency } = req.body;

  if (!agency) {
    return res.status(400).json({ ok: false, message: "agency 정보가 없습니다." });
  }

  const params = [snapshotDate];
  let idx = 2;

  let where = `WHERE snapshot_date = $1`;

  // 🔐 관리자 아니면 본인 대리점만
  if (agency !== "관리자") {
    where += ` AND agency_name = $${idx}`;
    params.push(agency);
    idx++;
  }

  try {
    const q = `
      SELECT
        COUNT(*)::int AS total_qty,
        COUNT(DISTINCT store_code)::int AS store_cnt,
        COUNT(DISTINCT model_name)::int AS model_cnt
      FROM inventory_items
      ${where}
    `;

    const r = await pool.query(q, params);

    return res.json({
      ok: true,
      snapshot_date: snapshotDate,
      agency,
      summary: r.rows[0]
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "요약 조회 실패" });
  }
});

/**
 * =========================
 * ✅ 재고 대시보드: 창고/판매점 분리 요약
 * =========================
 */
app.post("/inventory/summary-extended", async (req, res) => {
  const snapshotDate = todayKST();
  const { agency } = req.body;

  if (!agency) {
    return res.status(400).json({ ok: false, message: "agency 정보가 없습니다." });
  }

  const params = [snapshotDate];
  let idx = 2;

  let where = `WHERE snapshot_date = $1`;

  if (agency !== "관리자") {
    where += ` AND agency_name = $${idx}`;
    params.push(agency);
    idx++;
  }

  try {
    const q = `
      SELECT
        COUNT(*)::int AS total_qty,
        COUNT(DISTINCT store_code)::int AS store_cnt,
        COUNT(DISTINCT model_name)::int AS model_cnt,
        SUM(CASE WHEN store_name ILIKE '%창고%' THEN 1 ELSE 0 END)::int AS warehouse_qty,
        SUM(CASE WHEN store_name NOT ILIKE '%창고%' THEN 1 ELSE 0 END)::int AS store_qty
      FROM inventory_items
      ${where}
    `;

    const r = await pool.query(q, params);

    res.json({
      ok: true,
      snapshot_date: snapshotDate,
      summary: r.rows[0]
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, message: "확장 요약 실패" });
  }
});

/**
 * =========================
 * ✅ 재고 대시보드: 모델별 TOP
 * POST /inventory/by-model
 * body: { agency, limit }
 * =========================
 */
app.post("/inventory/by-model", async (req, res) => {
  const snapshotDate = todayKST();
  const { agency, limit } = req.body;

  if (!agency) {
    return res.status(400).json({ ok: false, message: "agency 정보가 없습니다." });
  }

  const topN = Number(limit) > 0 ? Math.min(Number(limit), 100) : 20;

  const params = [snapshotDate];
  let idx = 2;

  let where = `WHERE snapshot_date = $1`;

  if (agency !== "관리자") {
    where += ` AND agency_name = $${idx}`;
    params.push(agency);
    idx++;
  }

  try {
    const q = `
      SELECT
        model_name,
        COUNT(*)::int AS qty
      FROM inventory_items
      ${where}
      GROUP BY model_name
      ORDER BY qty DESC, model_name ASC
      LIMIT ${topN}
    `;

    const r = await pool.query(q, params);

    return res.json({
      ok: true,
      snapshot_date: snapshotDate,
      agency,
      rows: r.rows
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "모델별 조회 실패" });
  }
});

/**
 * =========================
 * ✅ 재고 대시보드: 판매점(접점)별 TOP
 * POST /inventory/by-store
 * body: { agency, limit }
 * =========================
 */
app.post("/inventory/by-store", async (req, res) => {
  const snapshotDate = todayKST();
  const { agency, limit } = req.body;

  if (!agency) {
    return res.status(400).json({ ok: false, message: "agency 정보가 없습니다." });
  }

  const topN = Number(limit) > 0 ? Math.min(Number(limit), 200) : 30;

  const params = [snapshotDate];
  let idx = 2;

  let where = `WHERE snapshot_date = $1`;

  if (agency !== "관리자") {
    where += ` AND agency_name = $${idx}`;
    params.push(agency);
    idx++;
  }

  try {
    const q = `
      SELECT
        store_code,
        store_name,
        COUNT(*)::int AS qty
      FROM inventory_items
      ${where}
      GROUP BY store_code, store_name
      ORDER BY qty DESC, store_name ASC
      LIMIT ${topN}
    `;

    const r = await pool.query(q, params);

    return res.json({
      ok: true,
      snapshot_date: snapshotDate,
      agency,
      rows: r.rows
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "판매점별 조회 실패" });
  }
});

/**
 * =========================
 * ✅ 재고 추천 이동 모델
 * =========================
 */
app.post("/inventory/recommend-move", async (req, res) => {
  const snapshotDate = todayKST();
  const { agency } = req.body;

  if (!agency) {
    return res.status(400).json({ ok: false, message: "agency 필요" });
  }

  const params = [snapshotDate];
  let idx = 2;
  let where = `WHERE snapshot_date = $1`;

  if (agency !== "관리자") {
    where += ` AND agency_name = $${idx}`;
    params.push(agency);
    idx++;
  }

  try {
    const q = `
      SELECT
        agency_name,
        store_code,
        store_name,
        model_name,
        color,
        COUNT(*)::int AS qty
      FROM inventory_items
      ${where}
      AND store_name NOT ILIKE '%창고%'
      GROUP BY agency_name, store_code, store_name, model_name, color
    `;

    const r = await pool.query(q, params);
    const rows = r.rows || [];

    // 모델+색상별로 묶기
    const grouped = {};
    for (const row of rows) {
      const model = row.model_name || "";
      const color = row.color || "";
      const key = `${model}__${color}`;

      if (!grouped[key]) grouped[key] = [];
      grouped[key].push({
        agency_name: row.agency_name,
        store_code: row.store_code,
        store_name: row.store_name,
        model_name: model,
        color,
        qty: Number(row.qty || 0)
      });
    }

    const recommendations = [];

    Object.values(grouped).forEach(list => {
      // 많은 곳 / 적은 곳 정렬
      const sortedDesc = [...list].sort((a, b) => b.qty - a.qty);
      const sortedAsc = [...list].sort((a, b) => a.qty - b.qty);

      const high = sortedDesc[0];
      const low = sortedAsc[0];

      // 룰:
      // 공급 가능: 5대 이상
      // 부족: 1대 이하
      // 차이 3대 이상일 때만 추천
      if (
        high &&
        low &&
        high.store_code !== low.store_code &&
        high.qty >= 5 &&
        low.qty <= 1 &&
        high.qty - low.qty >= 3
      ) {
        recommendations.push({
          model_name: high.model_name,
          color: high.color,
          from_store_code: high.store_code,
          from_store_name: high.store_name,
          from_qty: high.qty,
          to_store_code: low.store_code,
          to_store_name: low.store_name,
          to_qty: low.qty,
          gap: high.qty - low.qty
        });
      }
    });

    recommendations.sort((a, b) => {
      if (b.gap !== a.gap) return b.gap - a.gap;
      return (a.model_name || "").localeCompare(b.model_name || "", "ko");
    });

    return res.json({
      ok: true,
      snapshot_date: snapshotDate,
      agency,
      rows: recommendations.slice(0, 30)
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "추천 조회 실패" });
  }
});

/**
 * =========================
 * ✅ 판매점 상세: 특정 판매점 재고 리스트
 * POST /inventory/store-detail
 * body: { agency, store_code }
 * =========================
 */
app.post("/inventory/store-detail", async (req, res) => {
  const snapshotDate = todayKST();
  const { agency, store_code } = req.body;

  if (!agency) {
    return res.status(400).json({ ok: false, message: "agency 정보가 없습니다." });
  }
  if (!store_code) {
    return res.status(400).json({ ok: false, message: "store_code가 없습니다." });
  }

  const params = [snapshotDate];
  let idx = 2;

  let where = `WHERE snapshot_date = $1 AND store_code = $${idx}`;
  params.push(store_code);
  idx++;

  if (agency !== "관리자") {
    where += ` AND agency_name = $${idx}`;
    params.push(agency);
    idx++;
  }

  try {
    const q = `
      SELECT
        store_code, store_name,
        model_name, pet_name, color,
        serial_no, nickname,
        address
      FROM inventory_items
      ${where}
      ORDER BY model_name ASC, color ASC, serial_no ASC
      LIMIT 3000
    `;

    const r = await pool.query(q, params);

    return res.json({
      ok: true,
      snapshot_date: snapshotDate,
      agency,
      total: r.rows.length,
      rows: r.rows
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "판매점 상세 조회 실패" });
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
