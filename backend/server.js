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
 * =========================
 * 공통 유틸
 * =========================
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

function toNumber(v) {
  if (v === null || v === undefined || v === "") return 0;
  const n = Number(String(v).replace(/,/g, "").trim());
  return Number.isNaN(n) ? 0 : n;
}

function toDateText(v) {
  if (!v) return null;

  if (v instanceof Date && !isNaN(v)) {
    return new Intl.DateTimeFormat("en-CA", {
      timeZone: "Asia/Seoul"
    }).format(v);
  }

  const s = String(v).trim();
  const normalized = s.replace(/\./g, "-").replace(/\//g, "-");
  const d = new Date(normalized);

  if (!isNaN(d)) {
    return new Intl.DateTimeFormat("en-CA", {
      timeZone: "Asia/Seoul"
    }).format(d);
  }

  return null;
}

function getClientIp(req) {
  const forwarded = req.headers["x-forwarded-for"];
  if (forwarded) {
    return String(forwarded).split(",")[0].trim();
  }

  return (
    req.ip ||
    req.connection?.remoteAddress ||
    req.socket?.remoteAddress ||
    ""
  );
}

function normalizeSheetRows(rows, metricType, dataScope, baseMonth) {
  return rows.map(r => {
    const common = {
      data_scope: dataScope,
      metric_type: metricType,
      record_date: toDateText(r["일자"]),
      base_month: baseMonth || null,
      is_ms: toText(r["M&S여부"]),
      agency_code: toText(r["대리점코드"]),
      agency_name: toText(r["대리점명"]),
      store_code: toText(r["판매점코드"]),
      store_name: toText(r["판매점명"]),
      market: toText(r["상권"]),
      manager_name: toText(r["영업M"]),
      model_name: null,
      product_name: null,
      wireless_type: null,
      total_score: toNumber(r["총실적"]),
      new_010: 0,
      mnp: 0,
      change_device: 0
    };

    if (metricType === "후불") {
      common.model_name = toText(r["단말기모델"]);
      common.new_010 = toNumber(r["010신규"]);
      common.mnp = toNumber(r["MNP"]);
      common.change_device = toNumber(r["기변"]);
    }

    if (metricType === "순신규" || metricType === "약정갱신") {
      common.product_name = toText(r["상품"]);
    }

    if (metricType === "MIT") {
      common.product_name = toText(r["상품"]);
      common.wireless_type = toText(r["무선유형"]);
    }

    return common;
  });
}

function makeSearchableText(row, sheetName, baseMonth) {
  const pairs = Object.entries(row || {})
    .map(([k, v]) => `${k} ${String(v ?? "").trim()}`)
    .join(" ");

  return `기준월 ${baseMonth} 시트 ${sheetName} ${pairs}`.toLowerCase();
}

/**
 * =========================
 * 원본 raw 검색용 유틸
 * =========================
 */
function normalizeQuestionForSearch(question) {
  let q = (question || "").toLowerCase().trim();

  // 26년 2월 -> 2026-02, 2026년 2월 -> 2026-02
  q = q.replace(/(\d{2})년\s*(\d{1,2})월/g, (_, yy, mm) => {
    return `20${yy}-${String(mm).padStart(2, "0")}`;
  });

  q = q.replace(/(20\d{2})년\s*(\d{1,2})월/g, (_, yyyy, mm) => {
    return `${yyyy}-${String(mm).padStart(2, "0")}`;
  });

  // m&s -> ms
  q = q.replace(/m\s*&\s*s/gi, "ms");

  return q;
}

function extractMeaningfulKeywords(question) {
  const normalized = normalizeQuestionForSearch(question);

  const stopwords = new Set([
    "알려줘", "알려주", "좀", "해주세요", "해줘", "뭐야", "무엇", "어떻게",
    "의", "가", "이", "은", "는", "을", "를", "좀", "한번", "한", "번",
    "실적좀", "실적", "데이터", "내용", "기반", "업로드", "엑셀", "해주세요"
  ]);

  return normalized
    .split(/[\s,()/]+/)
    .map(v => v.trim())
    .filter(v => v.length >= 2)
    .filter(v => !stopwords.has(v));
}

/**
 * =========================
 * 실적 업로드 공통
 * =========================
 */
async function insertSalesData({
  fileBuffer,
  fileName,
  uploadType,
  baseMonth,
  uploadedBy
}) {
  const workbook = XLSX.read(fileBuffer, { type: "buffer" });

  const requiredSheets = ["후불", "순신규", "약정갱신", "MIT", "판매점LIST"];
  for (const sheetName of requiredSheets) {
    if (!workbook.SheetNames.includes(sheetName)) {
      throw new Error(`엑셀에 '${sheetName}' 시트가 없습니다.`);
    }
  }

  const postpaidRows = XLSX.utils.sheet_to_json(workbook.Sheets["후불"], { defval: "" });
  const pureNewRows = XLSX.utils.sheet_to_json(workbook.Sheets["순신규"], { defval: "" });
  const renewalRows = XLSX.utils.sheet_to_json(workbook.Sheets["약정갱신"], { defval: "" });
  const mitRows = XLSX.utils.sheet_to_json(workbook.Sheets["MIT"], { defval: "" });
  const storeRows = XLSX.utils.sheet_to_json(workbook.Sheets["판매점LIST"], { defval: "" });

  // 기존 집계/요약용 records
  const allSalesRows = [
    ...normalizeSheetRows(postpaidRows, "후불", uploadType, baseMonth),
    ...normalizeSheetRows(pureNewRows, "순신규", uploadType, baseMonth),
    ...normalizeSheetRows(renewalRows, "약정갱신", uploadType, baseMonth),
    ...normalizeSheetRows(mitRows, "MIT", uploadType, baseMonth)
  ];

  // 판매점 마스터
  const storeMasterRows = storeRows
    .map(r => ({
      store_code: toText(r["판매점코드"]),
      store_name: toText(r["판매점명"]),
      market_category: toText(r["상권구분"]),
      address: toText(r["주소"])
    }))
    .filter(r => r.store_code !== "");

  // Gemma 검색용 경량 텍스트 rows
  const aiRows = [];

  function pushAiRows(rows, sheetName) {
    rows.forEach((r) => {
      const searchableText = makeSearchableText(r, sheetName, baseMonth);

      aiRows.push({
        sheet_name: sheetName,
        is_ms: toText(r["M&S여부"]),
        agency_name: toText(r["대리점명"]),
        store_code: toText(r["판매점코드"]),
        store_name: toText(r["판매점명"]),
        searchable_text: searchableText
      });
    });
  }

  pushAiRows(postpaidRows, "후불");
  pushAiRows(pureNewRows, "순신규");
  pushAiRows(renewalRows, "약정갱신");
  pushAiRows(mitRows, "MIT");
  pushAiRows(storeRows, "판매점LIST");

  const client = await pool.connect();

  try {
    await client.query("BEGIN");

    const batchResult = await client.query(
      `
      INSERT INTO sales_upload_batches
      (upload_type, base_month, file_name, uploaded_by)
      VALUES ($1, $2, $3, $4)
      RETURNING id
      `,
      [uploadType, baseMonth || null, fileName || null, uploadedBy || "관리자"]
    );

    const batchId = String(batchResult.rows[0].id);

    // 같은 월/구분 재업로드 시 기존 데이터 삭제
    if (baseMonth) {
      await client.query(
        `
        DELETE FROM sales_records
        WHERE data_scope = $1
          AND base_month = $2
        `,
        [uploadType, baseMonth]
      );

      await client.query(
        `
        DELETE FROM sales_ai_rows
        WHERE data_scope = $1
          AND base_month = $2
        `,
        [uploadType, baseMonth]
      );
    }

    // 판매점 마스터 갱신
    if (storeMasterRows.length > 0) {
      await client.query(`DELETE FROM sales_store_master`);

      for (let i = 0; i < storeMasterRows.length; i += 300) {
        const chunk = storeMasterRows.slice(i, i + 300);
        const values = [];

        const placeholders = chunk.map((row, idx) => {
          const b = idx * 4;
          values.push(
            row.store_code,
            row.store_name,
            row.market_category,
            row.address
          );
          return `($${b + 1}, $${b + 2}, $${b + 3}, $${b + 4})`;
        });

        await client.query(
          `
          INSERT INTO sales_store_master
          (store_code, store_name, market_category, address)
          VALUES ${placeholders.join(",")}
          `,
          values
        );
      }
    }

    // sales_records 저장
    for (let i = 0; i < allSalesRows.length; i += 300) {
      const chunk = allSalesRows.slice(i, i + 300);
      const values = [];

      const placeholders = chunk.map((row, idx) => {
        const b = idx * 19;

        values.push(
          batchId,
          row.data_scope,
          row.metric_type,
          row.record_date,
          row.base_month,
          row.is_ms,
          row.agency_code,
          row.agency_name,
          row.store_code,
          row.store_name,
          row.market,
          row.manager_name,
          row.model_name,
          row.product_name,
          row.wireless_type,
          row.total_score,
          row.new_010,
          row.mnp,
          row.change_device
        );

        return `(
          $${b + 1}, $${b + 2}, $${b + 3}, $${b + 4}, $${b + 5},
          $${b + 6}, $${b + 7}, $${b + 8}, $${b + 9}, $${b + 10},
          $${b + 11}, $${b + 12}, $${b + 13}, $${b + 14}, $${b + 15},
          $${b + 16}, $${b + 17}, $${b + 18}, $${b + 19}
        )`;
      });

      await client.query(
        `
        INSERT INTO sales_records
        (
          batch_id,
          data_scope,
          metric_type,
          record_date,
          base_month,
          is_ms,
          agency_code,
          agency_name,
          store_code,
          store_name,
          market,
          manager_name,
          model_name,
          product_name,
          wireless_type,
          total_score,
          new_010,
          mnp,
          change_device
        )
        VALUES ${placeholders.join(",")}
        `,
        values
      );
    }

    // sales_ai_rows 저장 (가벼운 텍스트만)
    for (let i = 0; i < aiRows.length; i += 300) {
      const chunk = aiRows.slice(i, i + 300);
      const values = [];

      const placeholders = chunk.map((row, idx) => {
        const b = idx * 9;

        values.push(
          batchId,
          uploadType,
          baseMonth,
          row.sheet_name,
          row.is_ms,
          row.agency_name,
          row.store_code,
          row.store_name,
          row.searchable_text
        );

        return `(
          $${b + 1}, $${b + 2}, $${b + 3}, $${b + 4}, $${b + 5},
          $${b + 6}, $${b + 7}, $${b + 8}, $${b + 9}
        )`;
      });

      await client.query(
        `
        INSERT INTO sales_ai_rows
        (
          batch_id,
          data_scope,
          base_month,
          sheet_name,
          is_ms,
          agency_name,
          store_code,
          store_name,
          searchable_text
        )
        VALUES ${placeholders.join(",")}
        `,
        values
      );
    }

    await client.query("COMMIT");

    return {
      batch_id: batchId,
      sales_count: allSalesRows.length,
      store_count: storeMasterRows.length,
      ai_count: aiRows.length
    };
  } catch (e) {
    await client.query("ROLLBACK");
    throw e;
  } finally {
    client.release();
  }
}

/**
 * =========================
 * 재고 업로드 / 검색
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

      for (let i = 0; i < data.length; i += 500) {
        const chunk = data.slice(i, i + 500);
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

          return `(
            $${b + 1}, $${b + 2}, $${b + 3}, $${b + 4}, $${b + 5}, $${b + 6},
            $${b + 7}, $${b + 8}, $${b + 9}, $${b + 10}, $${b + 11}, $${b + 12}
          )`;
        });

        await client.query(
          `
          INSERT INTO inventory_items
          (
            snapshot_date, agency_code, agency_name, sub_market, address,
            store_code, store_name, pet_name, model_name, color, serial_no, nickname
          )
          VALUES ${placeholders.join(",")}
          `,
          values
        );
      }

      await client.query("COMMIT");

      return res.json({
        ok: true,
        count: data.length,
        snapshot_date: snapshotDate
      });
    } catch (e) {
      await client.query("ROLLBACK");
      console.error(e);
      return res.status(500).json({ ok: false, message: "DB 오류" });
    } finally {
      client.release();
    }
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "업로드 실패" });
  }
});

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

    return res.json({
      ok: true,
      agency,
      snapshot_date: snapshotDate,
      total: result.rows.length,
      table: result.rows
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "검색 실패" });
  }
});

/**
 * =========================
 * 재고 대시보드
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

    return res.json({
      ok: true,
      snapshot_date: snapshotDate,
      summary: r.rows[0]
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "확장 요약 실패" });
  }
});

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
      const sortedDesc = [...list].sort((a, b) => b.qty - a.qty);
      const sortedAsc = [...list].sort((a, b) => a.qty - b.qty);

      const high = sortedDesc[0];
      const low = sortedAsc[0];

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
        store_code,
        store_name,
        model_name,
        pet_name,
        color,
        serial_no,
        nickname,
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

app.post("/inventory/warehouse-detail", async (req, res) => {
  const snapshotDate = todayKST();
  const { agency } = req.body;

  if (!agency) {
    return res.status(400).json({ ok: false, message: "agency 정보가 없습니다." });
  }

  const params = [snapshotDate];
  let idx = 2;
  let where = `WHERE snapshot_date = $1 AND store_name ILIKE '%창고%'`;

  if (agency !== "관리자") {
    where += ` AND agency_name = $${idx}`;
    params.push(agency);
    idx++;
  }

  try {
    const q = `
      SELECT
        model_name,
        color,
        COUNT(*)::int AS qty
      FROM inventory_items
      ${where}
      GROUP BY model_name, color
      ORDER BY qty DESC, model_name ASC, color ASC
      LIMIT 3000
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
    return res.status(500).json({ ok: false, message: "창고 재고 상세 조회 실패" });
  }
});

app.post("/inventory/total-detail", async (req, res) => {
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
        model_name,
        color,
        COUNT(*)::int AS total_qty,
        SUM(CASE WHEN store_name ILIKE '%창고%' THEN 1 ELSE 0 END)::int AS warehouse_qty,
        SUM(CASE WHEN store_name NOT ILIKE '%창고%' THEN 1 ELSE 0 END)::int AS store_qty,
        ROUND(
          (
            SUM(CASE WHEN store_name ILIKE '%창고%' THEN 1 ELSE 0 END)::numeric
            / NULLIF(COUNT(*)::numeric, 0)
          ) * 100, 1
        ) AS warehouse_ratio
      FROM inventory_items
      ${where}
      GROUP BY model_name, color
      ORDER BY total_qty DESC, model_name ASC, color ASC
      LIMIT 3000
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
    return res.status(500).json({ ok: false, message: "총 재고 상세 조회 실패" });
  }
});

app.post("/inventory/store-stock-detail", async (req, res) => {
  const snapshotDate = todayKST();
  const { agency } = req.body;

  if (!agency) {
    return res.status(400).json({ ok: false, message: "agency 정보가 없습니다." });
  }

  const params = [snapshotDate];
  let idx = 2;
  let where = `WHERE snapshot_date = $1 AND store_name NOT ILIKE '%창고%'`;

  if (agency !== "관리자") {
    where += ` AND agency_name = $${idx}`;
    params.push(agency);
    idx++;
  }

  try {
    const q = `
      SELECT
        agency_name,
        store_name,
        model_name,
        COUNT(*)::int AS qty
      FROM inventory_items
      ${where}
      GROUP BY agency_name, store_name, model_name
      ORDER BY agency_name ASC, store_name ASC, qty DESC, model_name ASC
      LIMIT 5000
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
    return res.status(500).json({ ok: false, message: "판매점 재고 상세 조회 실패" });
  }
});

app.get("/upload-status", async (req, res) => {
  const snapshotDate = todayKST();

  try {
    const r = await pool.query(
      "SELECT COUNT(*)::int AS cnt FROM inventory_items WHERE snapshot_date = $1",
      [snapshotDate]
    );

    return res.json({
      ok: true,
      snapshot_date: snapshotDate,
      today_count: r.rows[0]?.cnt ?? 0
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "업로드 상태 조회 실패" });
  }
});

/**
 * =========================
 * 접속자 카운트
 * =========================
 */
app.post("/login-log", async (req, res) => {
  const loginDate = todayKST();
  const { agency } = req.body;
  const ipAddress = getClientIp(req);
  const userAgent = req.headers["user-agent"] || "";

  if (!agency) {
    return res.status(400).json({ ok: false, message: "agency 정보가 없습니다." });
  }

  if (agency === "관리자") {
    return res.json({ ok: true, skipped: true });
  }

  try {
    await pool.query(
      `
      INSERT INTO login_logs (login_date, agency_name, ip_address, user_agent)
      VALUES ($1, $2, $3, $4)
      ON CONFLICT (login_date, agency_name, ip_address) DO NOTHING
      `,
      [loginDate, agency, ipAddress, userAgent]
    );

    return res.json({ ok: true, login_date: loginDate });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "로그인 기록 저장 실패" });
  }
});

app.get("/login-today-summary", async (req, res) => {
  const loginDate = todayKST();
  const agency = req.query.agency;

  if (!agency) {
    return res.status(400).json({ ok: false, message: "agency 정보가 없습니다." });
  }

  try {
    if (agency === "관리자") {
      const r = await pool.query(
        `
        SELECT agency_name, COUNT(*)::int AS cnt
        FROM login_logs
        WHERE login_date = $1
        GROUP BY agency_name
        ORDER BY
          CASE agency_name
            WHEN '광주' THEN 1
            WHEN '목포' THEN 2
            WHEN '순천' THEN 3
            WHEN '전북' THEN 4
            WHEN '제주' THEN 5
            ELSE 99
          END
        `,
        [loginDate]
      );

      return res.json({
        ok: true,
        login_date: loginDate,
        agency,
        rows: r.rows
      });
    }

    const r = await pool.query(
      `
      SELECT COUNT(*)::int AS cnt
      FROM login_logs
      WHERE login_date = $1
        AND agency_name = $2
      `,
      [loginDate, agency]
    );

    return res.json({
      ok: true,
      login_date: loginDate,
      agency,
      count: r.rows[0]?.cnt ?? 0
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "접속자 조회 실패" });
  }
});

/**
 * =========================
 * 실적 업로드
 * =========================
 */
app.post("/sales/upload-monthly", upload.single("file"), async (req, res) => {
  try {
    console.log("📥 /sales/upload-monthly 시작");
    console.log("base_month:", req.body?.base_month);
    console.log("uploaded_by:", req.body?.uploaded_by);
    console.log("file name:", req.file?.originalname);
    console.log("file size:", req.file?.size);

    const baseMonth = toText(req.body.base_month);
    const uploadedBy = toText(req.body.uploaded_by) || "관리자";

    if (!req.file) {
      return res.status(400).json({ ok: false, message: "파일이 없습니다." });
    }

    if (!baseMonth) {
      return res.status(400).json({
        ok: false,
        message: "base_month가 없습니다. 예: 2026-02"
      });
    }

    const result = await insertSalesData({
      fileBuffer: req.file.buffer,
      fileName: req.file.originalname,
      uploadType: "monthly",
      baseMonth,
      uploadedBy
    });

    console.log("✅ /sales/upload-monthly 완료", result);

    return res.json({
      ok: true,
      message: "월 누적 실적 업로드 완료",
      base_month: baseMonth,
      batch_id: result.batch_id,
      sales_count: result.sales_count,
      store_count: result.store_count,
      raw_count: result.raw_count
    });
  } catch (e) {
    console.error("❌ /sales/upload-monthly 실패", e);
    return res.status(500).json({
      ok: false,
      message: e.message || "월 누적 실적 업로드 실패"
    });
  }
});

app.post("/sales/upload-daily", upload.single("file"), async (req, res) => {
  try {
    console.log("📥 /sales/upload-daily 시작");
    console.log("base_month:", req.body?.base_month);
    console.log("uploaded_by:", req.body?.uploaded_by);
    console.log("file name:", req.file?.originalname);
    console.log("file size:", req.file?.size);

    const baseMonth = toText(req.body.base_month);
    const uploadedBy = toText(req.body.uploaded_by) || "관리자";

    if (!req.file) {
      return res.status(400).json({ ok: false, message: "파일이 없습니다." });
    }

    if (!baseMonth) {
      return res.status(400).json({
        ok: false,
        message: "base_month가 없습니다. 예: 2026-03"
      });
    }

    const result = await insertSalesData({
      fileBuffer: req.file.buffer,
      fileName: req.file.originalname,
      uploadType: "daily",
      baseMonth,
      uploadedBy
    });

    console.log("✅ /sales/upload-daily 완료", result);

    return res.json({
      ok: true,
      message: "당월 실적 업로드 완료",
      base_month: baseMonth,
      batch_id: result.batch_id,
      sales_count: result.sales_count,
      store_count: result.store_count,
      raw_count: result.raw_count
    });
  } catch (e) {
    console.error("❌ /sales/upload-daily 실패", e);
    return res.status(500).json({
      ok: false,
      message: e.message || "당월 실적 업로드 실패"
    });
  }
});

/**
 * =========================
 * 실적 조회
 * =========================
 */
app.get("/sales/upload-history", async (req, res) => {
  try {
    const r = await pool.query(
      `
      SELECT
        id,
        upload_type,
        base_month,
        file_name,
        uploaded_by,
        uploaded_at
      FROM sales_upload_batches
      ORDER BY uploaded_at DESC
      LIMIT 20
      `
    );

    return res.json({
      ok: true,
      rows: r.rows
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "업로드 이력 조회 실패" });
  }
});

app.get("/sales/summary", async (req, res) => {
  const baseMonth = toText(req.query.base_month);

  if (!baseMonth) {
    return res.status(400).json({ ok: false, message: "base_month가 필요합니다." });
  }

  try {
    const totalQ = await pool.query(
      `
      SELECT
        COUNT(*)::int AS row_count,
        COALESCE(SUM(total_score), 0)::numeric AS total_score,
        COUNT(DISTINCT agency_name)::int AS agency_count,
        COUNT(DISTINCT store_code)::int AS store_count
      FROM sales_records
      WHERE base_month = $1
        AND is_ms = 'Y'
      `,
      [baseMonth]
    );

    const typeQ = await pool.query(
      `
      SELECT
        metric_type,
        COUNT(*)::int AS row_count,
        COALESCE(SUM(total_score), 0)::numeric AS total_score
      FROM sales_records
      WHERE base_month = $1
        AND is_ms = 'Y'
      GROUP BY metric_type
      ORDER BY
        CASE metric_type
          WHEN '후불' THEN 1
          WHEN '순신규' THEN 2
          WHEN '약정갱신' THEN 3
          WHEN 'MIT' THEN 4
          ELSE 99
        END
      `,
      [baseMonth]
    );

    const agencyQ = await pool.query(
      `
      SELECT
        agency_name,
        COALESCE(SUM(total_score), 0)::numeric AS total_score
      FROM sales_records
      WHERE base_month = $1
        AND is_ms = 'Y'
      GROUP BY agency_name
      ORDER BY total_score DESC, agency_name ASC
      `,
      [baseMonth]
    );

    return res.json({
      ok: true,
      base_month: baseMonth,
      summary: totalQ.rows[0],
      by_type: typeQ.rows,
      by_agency: agencyQ.rows
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "실적 요약 조회 실패" });
  }
});

app.post("/sales/query", async (req, res) => {
  const { base_month, agency, metric_type } = req.body;

  if (!base_month) {
    return res.status(400).json({ ok: false, message: "base_month 필요" });
  }

  try {
    let where = `WHERE base_month = $1 AND is_ms = 'Y'`;
    const params = [base_month];
    let idx = 2;

    if (agency) {
      where += ` AND agency_name = $${idx}`;
      params.push(agency);
      idx++;
    }

    if (metric_type) {
      where += ` AND metric_type = $${idx}`;
      params.push(metric_type);
      idx++;
    }

    const q = `
      SELECT
        agency_name,
        metric_type,
        COUNT(*)::int AS row_count,
        COALESCE(SUM(total_score), 0)::numeric AS total_score
      FROM sales_records
      ${where}
      GROUP BY agency_name, metric_type
      ORDER BY total_score DESC
    `;

    const r = await pool.query(q, params);

    return res.json({
      ok: true,
      base_month,
      rows: r.rows
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "질의 실패" });
  }
});

/**
 * =========================
 * AI 테스트
 * =========================
 */
app.post("/ai/test", async (req, res) => {
  try {
    const response = await fetch("http://127.0.0.1:11434/api/generate", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: "gemma3:4b",
        prompt: "너는 M&S 호남도매 실적 분석 AI야. 간단히 자기소개 해봐.",
        stream: false
      })
    });

    const data = await response.json();

    return res.json({
      ok: true,
      result: data.response
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "AI 호출 실패" });
  }
});

/**
 * =========================
 * 엑셀 원본 기반 AI 질의
 * =========================
 */
app.post("/sales/ask-ai-raw", async (req, res) => {
  const question = toText(req.body.question);
  const loginAgency = toText(req.body.agency);

  if (!question) {
    return res.status(400).json({ ok: false, message: "question 필요" });
  }

  try {
    const keywords = extractMeaningfulKeywords(question);

    if (!keywords.length) {
      return res.json({
        ok: true,
        answer: "질문에서 검색할 핵심 단어를 찾지 못했습니다."
      });
    }

    let where = `WHERE 1=1`;
    const params = [];
    let idx = 1;

    // 일반 센터 사용자는 자기 센터만
    if (loginAgency && loginAgency !== "관리자") {
      where += ` AND agency_name = $${idx}`;
      params.push(loginAgency);
      idx++;
    }

    const orParts = [];
    for (const keyword of keywords) {
      orParts.push(`searchable_text ILIKE $${idx}`);
      params.push(`%${keyword}%`);
      idx++;
    }

    if (orParts.length) {
      where += ` AND (${orParts.join(" OR ")})`;
    }

    const q = `
      SELECT
        sheet_name,
        base_month,
        agency_name,
        store_code,
        store_name,
        searchable_text
      FROM sales_ai_rows
      ${where}
      ORDER BY id DESC
      LIMIT 30
    `;

    const r = await pool.query(q, params);

    if (!r.rows.length) {
      return res.json({
        ok: true,
        answer: "엑셀 데이터에서 해당 조건을 찾지 못했습니다."
      });
    }

    const contextText = r.rows
      .map(row => row.searchable_text)
      .join("\n");

    const prompt = `
너는 H.O.S 실적 분석 AI다.

반드시 아래 규칙을 지켜라.
- 아래 데이터 안에서만 답해라.
- 데이터에 없는 내용은 절대 추측하지 마라.
- 질문과 관련 없는 말은 하지 마라.
- 한국어로 간단명료하게 답해라.
- 숫자, 센터, 판매점, 코드, 주소, 상품, 모델 등은 데이터 기준으로만 말해라.

[사용자 질문]
${question}

[검색 키워드]
${JSON.stringify(keywords)}

[엑셀 검색 결과]
${contextText}

위 데이터만 근거로 답변해라.
`;

    const response = await fetch("http://127.0.0.1:11434/api/generate", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: "gemma3:4b",
        prompt,
        stream: false
      })
    });

    const data = await response.json();

    return res.json({
      ok: true,
      answer: data.response || "응답이 없습니다."
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({
      ok: false,
      message: "AI raw 질의 실패"
    });
  }
});

/**
 * =========================
 * DB 연결 테스트
 * =========================
 */
pool.query("SELECT NOW()")
  .then(r => console.log("✅ DB 연결 성공:", r.rows[0]))
  .catch(err => console.error("❌ DB 연결 실패:", err));

/**
 * =========================
 * CRON 로그
 * =========================
 */
cron.schedule("0 * * * *", async () => {
  try {
    const d = todayKST();
    const r = await pool.query(
      "SELECT COUNT(*)::int AS cnt FROM inventory_items WHERE snapshot_date = $1",
      [d]
    );
    console.log(`🕐 [CRON] ${d} 재고 ${r.rows[0].cnt}건`);
  } catch (e) {
    console.error("❌ CRON 오류:", e);
  }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log("🚀 Backend running on port", PORT);
});