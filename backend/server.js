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

function mapCenterToSalesAgency(center) {
  const name = toText(center);

  if (!name || name === "관리자") return "";

  const mapping = {
    "광주": "M&S광주",
    "목포": "M&S목포",
    "순천": "M&S순천",
    "전북": "M&S전북",
    "제주": "M&S제주"
  };

  return mapping[name] || name;
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
      "애칭",
      "입고경과일"
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
        nickname: toText(r["애칭"]),
        aging_days: Number(toText(r["입고경과일"]).replace(/,/g, "")) || 0
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
          const b = idx * 13;
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
            row.nickname,
            row.aging_days
          );

          return `(
            $${b + 1}, $${b + 2}, $${b + 3}, $${b + 4}, $${b + 5}, $${b + 6},
            $${b + 7}, $${b + 8}, $${b + 9}, $${b + 10}, $${b + 11}, $${b + 12},
            $${b + 13}
          )`;
        });

        await client.query(
          `
          INSERT INTO inventory_items
          (
            snapshot_date, agency_code, agency_name, sub_market, address,
            store_code, store_name, pet_name, model_name, color, serial_no, nickname,
            aging_days
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

// ================= 실적 대시보드 =================
app.get("/api/performance/summary", async (req, res) => {
  try {
    const { baseMonth } = req.query;

    if (!baseMonth) {
      return res.status(400).json({ error: "baseMonth 필요" });
    }

    // 유형별 합계
    const result = await pool.query(`
      SELECT 
        metric_type,
        SUM(total_score) as total
      FROM sales_records
      WHERE base_month = $1
      GROUP BY metric_type
    `, [baseMonth]);

    // 결과 정리
    const data = {
      후불: 0,
      순신규: 0,
      약정갱신: 0,
      MIT: 0
    };

    result.rows.forEach(r => {
      data[r.metric_type] = Number(r.total);
    });

    // 실적점 (후불 기준 판매점 수)
    const storeCount = await pool.query(`
      SELECT COUNT(DISTINCT store_code)
      FROM sales_records
      WHERE base_month = $1
        AND metric_type = '후불'
    `, [baseMonth]);

    data["실적점"] = Number(storeCount.rows[0].count);

    res.json({ ok: true, data });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "서버 오류" });
  }
});

/**
 * =========================
 * 실적 대시보드 요약
 * GET /performance/dashboard-summary?agency=광주
 * =========================
 */
app.get("/performance/dashboard-summary", async (req, res) => {
  try {
    const agency = mapCenterToSalesAgency(req.query.agency);

    const monthQ = await pool.query(
      `
      SELECT MAX(base_month) AS latest_month
      FROM sales_records
      WHERE data_scope = 'daily'
      `
    );

    const latestMonth =
  monthQ.rows[0]?.latest_month ||
  (
    await pool.query(
      `
      SELECT MAX(base_month) AS latest_month
      FROM sales_records
      WHERE data_scope = 'monthly'
      `
    )
  ).rows[0]?.latest_month;

// 기준일 = 당월(후불) 마지막 마감일
const dateQ = await pool.query(
  `
  SELECT MAX(record_date) AS latest_date
  FROM sales_records
  WHERE data_scope = 'daily'
    AND metric_type = '후불'
  `
);

const latestDate = dateQ.rows[0]?.latest_date || null;
    // -------------------------
    // 표준 진척율 계산
    // 기준: 일요일 제외, 월~토 영업일
    // -------------------------
        let progressRate = 0;

    if (latestDate) {
      const dateObj = new Date(latestDate);

      const year = dateObj.getUTCFullYear();
      const month = dateObj.getUTCMonth() + 1; // 1~12
      const day = dateObj.getUTCDate();

      // 해당 월 마지막 날짜
      const lastDayOfMonth = new Date(Date.UTC(year, month, 0)).getUTCDate();

      let totalBusinessDays = 0;
      let passedBusinessDays = 0;

      for (let d = 1; d <= lastDayOfMonth; d++) {
        const current = new Date(Date.UTC(year, month - 1, d));
        const weekday = current.getUTCDay(); // 0=일요일

        // 일요일 제외
        if (weekday !== 0) {
          totalBusinessDays++;

          if (d <= day) {
            passedBusinessDays++;
          }
        }
      }

      if (totalBusinessDays > 0) {
        progressRate = Math.round((passedBusinessDays / totalBusinessDays) * 100);
      }
    }

    if (!latestMonth) {
      return res.json({
        ok: true,
        latest_month: latestMonth,
        latest_date : latestDate,
        progress_rate: progressRate,
        summary: {
          postpaid: 0,
          pure_new: 0,
          renewal: 0,
          mit: 0,
          postpaid_store_count: 0,
          postpaid_rate: 0,
          pure_new_rate: 0,
          renewal_rate: 0,
          mit_rate: 0,
          postpaid_store_rate: 0,
          postpaid_share: 0,
pure_new_share: 0,
renewal_share: 0,
mit_share: 0
        }
      });
    }

    const params = [latestMonth];
    let idx = 2;

    let where = `
      WHERE base_month = $1
        AND is_ms = 'Y'
    `;

    // 관리자 아니면 해당 센터만
    if (agency && agency !== "관리자") {
      where += ` AND agency_name = $${idx}`;
      params.push(agency);
      idx++;
    }

    const totalQ = await pool.query(
      `
      SELECT
        COALESCE(SUM(CASE WHEN metric_type = '후불' THEN total_score ELSE 0 END), 0)::numeric AS postpaid,
        COALESCE(SUM(CASE WHEN metric_type = '순신규' THEN total_score ELSE 0 END), 0)::numeric AS pure_new,
        COALESCE(SUM(CASE WHEN metric_type = '약정갱신' THEN total_score ELSE 0 END), 0)::numeric AS renewal,
        COALESCE(SUM(CASE WHEN metric_type = 'MIT' THEN total_score ELSE 0 END), 0)::numeric AS mit
      FROM sales_records
      ${where}
      `,
      params
    );

        // -------------------------
    // 전체 KT 실적(비중 계산용 분모)
    // 관리자=호남 전체 / 센터선택=해당 센터 전체
    // -------------------------
    const totalKtParams = [latestMonth];
    let totalKtWhere = `
      WHERE base_month = $1
    `;

    if (agency && agency !== "관리자") {
      totalKtWhere += ` AND agency_name = $2`;
      totalKtParams.push(agency);
    }

    const totalKtQ = await pool.query(
      `
      SELECT
        COALESCE(SUM(CASE WHEN metric_type = '후불' THEN total_score ELSE 0 END), 0)::numeric AS postpaid_total,
        COALESCE(SUM(CASE WHEN metric_type = '순신규' THEN total_score ELSE 0 END), 0)::numeric AS pure_new_total,
        COALESCE(SUM(CASE WHEN metric_type = '약정갱신' THEN total_score ELSE 0 END), 0)::numeric AS renewal_total,
        COALESCE(SUM(CASE WHEN metric_type = 'MIT' THEN total_score ELSE 0 END), 0)::numeric AS mit_total
      FROM sales_records
      ${totalKtWhere}
      `,
      totalKtParams
    );

    const storeQ = await pool.query(
      `
      SELECT COUNT(DISTINCT store_code)::int AS cnt
      FROM sales_records
      ${where}
        AND metric_type = '후불'
        AND COALESCE(total_score, 0) > 0
      `,
      params
    );

    // -------------------------
    // 목표 조회
    // -------------------------
    let targetQ;

    if (agency && agency !== "관리자") {
      // 센터 로그인: 해당 센터 목표
      targetQ = await pool.query(
        `
        SELECT
          metric_type,
          COALESCE(SUM(target_value), 0)::numeric AS target_value
        FROM sales_targets
        WHERE base_month = $1
          AND agency_name = $2
        GROUP BY metric_type
        `,
        [latestMonth, agency]
      );
    } else {
      // 관리자 로그인: 전체 센터 합산 목표
      targetQ = await pool.query(
        `
        SELECT
          metric_type,
          COALESCE(SUM(target_value), 0)::numeric AS target_value
        FROM sales_targets
        WHERE base_month = $1
        GROUP BY metric_type
        `,
        [latestMonth]
      );
    }

    const targetMap = {};
    (targetQ.rows || []).forEach(r => {
      targetMap[r.metric_type] = Number(r.target_value || 0);
    });

    const postpaid = Number(totalQ.rows[0]?.postpaid || 0);
    const pureNew = Number(totalQ.rows[0]?.pure_new || 0);
    const renewal = Number(totalQ.rows[0]?.renewal || 0);
    const mit = Number(totalQ.rows[0]?.mit || 0);
    const postpaidStoreCount = Number(storeQ.rows[0]?.cnt || 0);
    const postpaidTotal = Number(totalKtQ.rows[0]?.postpaid_total || 0);
    const pureNewTotal = Number(totalKtQ.rows[0]?.pure_new_total || 0);
    const renewalTotal = Number(totalKtQ.rows[0]?.renewal_total || 0);
    const mitTotal = Number(totalKtQ.rows[0]?.mit_total || 0);

    function calcRate(actual, target) {
      if (!target || Number(target) === 0) return 0;
      return Math.round((Number(actual) / Number(target)) * 100);
    }

        function calcShare(actual, total) {
      if (!total || Number(total) === 0) return 0;
      return Number(((Number(actual) / Number(total)) * 100).toFixed(1));
    }

    return res.json({
  ok: true,
  latest_month: latestMonth,
  latest_date: latestDate,
  progress_rate: progressRate,
  summary: {
    postpaid,
    pure_new: pureNew,
    renewal,
    mit,
    postpaid_store_count: postpaidStoreCount,

    postpaid_rate: calcRate(postpaid, targetMap["후불"]),
    pure_new_rate: calcRate(pureNew, targetMap["순신규"]),
    renewal_rate: calcRate(renewal, targetMap["약정갱신"]),
    mit_rate: calcRate(mit, targetMap["MIT"]),
    postpaid_store_rate: calcRate(postpaidStoreCount, targetMap["후불실적점"]),

    postpaid_share: calcShare(postpaid, postpaidTotal),
    pure_new_share: calcShare(pureNew, pureNewTotal),
    renewal_share: calcShare(renewal, renewalTotal),
    mit_share: calcShare(mit, mitTotal)
  }
});
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "실적 대시보드 조회 실패" });
  }
});

/**
 * =========================
 * 최근 6개월 추이
 * GET /performance/dashboard-trend?metric=후불&agency=광주
 * metric:
 * - 후불
 * - 순신규
 * - 약정갱신
 * - MIT
 * - 후불실적점
 * =========================
 */
app.get("/performance/dashboard-trend", async (req, res) => {
  try {
    const metric = toText(req.query.metric);
    const rawAgency = toText(req.query.agency);           // 관리자 / 광주 / 목포 / 순천 / 전북 / 제주
    const agency = mapCenterToSalesAgency(rawAgency);     // 관리자 / M&S광주 / M&S목포 ...

    if (!metric) {
      return res.status(400).json({ ok: false, message: "metric이 필요합니다." });
    }

    const monthQ = await pool.query(
      `
      SELECT DISTINCT base_month
      FROM sales_records
      WHERE base_month IS NOT NULL
      ORDER BY base_month DESC
      LIMIT 6
      `
    );

    const months = monthQ.rows.map(r => r.base_month).reverse();

    if (!months.length) {
      return res.json({ ok: true, rows: [] });
    }

    const params = [months];
    let idx = 2;

    let commonWhere = `
      WHERE base_month = ANY($1)
        AND is_ms = 'Y'
    `;

    // 관리자 아니면 해당 M&S 센터만
    if (agency && agency !== "관리자") {
      commonWhere += ` AND agency_name = $${idx}`;
      params.push(agency);
      idx++;
    }

    let rows = [];

    // -------------------------
    // 1) 기존 값 조회 (기존 기능 유지)
    // -------------------------
    if (metric === "후불실적점") {
      const q = await pool.query(
        `
        SELECT
          base_month,
          COUNT(DISTINCT store_code)::int AS value
        FROM sales_records
        ${commonWhere}
          AND metric_type = '후불'
          AND COALESCE(total_score, 0) > 0
        GROUP BY base_month
        ORDER BY base_month ASC
        `,
        params
      );

      const map = {};
      q.rows.forEach(r => {
        map[r.base_month] = Number(r.value || 0);
      });

      rows = months.map(m => ({
        month: m,
        value: Number(map[m] || 0),
        share_rate: 0
      }));
    } else {
      const metricParams = [...params, metric];

      const q = await pool.query(
        `
        SELECT
          base_month,
          COALESCE(SUM(total_score), 0)::numeric AS value
        FROM sales_records
        ${commonWhere}
          AND metric_type = $${metricParams.length}
        GROUP BY base_month
        ORDER BY base_month ASC
        `,
        metricParams
      );

      const map = {};
      q.rows.forEach(r => {
        map[r.base_month] = Number(r.value || 0);
      });

      // -------------------------
      // 2) 비중 계산용 KT 전체 분모 조회
      //    - 관리자(호남): 전체 KT
      //    - 센터선택: 해당 센터 전체 KT
      // -------------------------
      const totalParams = [months, metric];
      let totalWhere = `
        WHERE base_month = ANY($1)
          AND metric_type = $2
      `;

      // 센터별 분모 범위
      // market 컬럼에 센터/권역 정보가 들어있다는 전제
      if (rawAgency && rawAgency !== "관리자" && rawAgency !== "호남") {
        if (rawAgency === "광주") {
          totalWhere += ` AND market ILIKE $3`;
          totalParams.push(`%광주%`);
        } else if (rawAgency === "목포") {
          totalWhere += ` AND market ILIKE $3`;
          totalParams.push(`%목포%`);
        } else if (rawAgency === "순천") {
          totalWhere += ` AND market ILIKE $3`;
          totalParams.push(`%순천%`);
        } else if (rawAgency === "제주") {
          totalWhere += ` AND market ILIKE $3`;
          totalParams.push(`%제주%`);
        } else if (rawAgency === "전북") {
          totalWhere += `
            AND (
              market ILIKE $3
              OR market ILIKE $4
              OR market ILIKE $5
              OR market ILIKE $6
              OR market ILIKE $7
              OR market ILIKE $8
              OR market ILIKE $9
              OR market ILIKE $10
              OR market ILIKE $11
            )
          `;
          totalParams.push(`%전북%`);
          totalParams.push(`%전라북도%`);
          totalParams.push(`%전주%`);
          totalParams.push(`%군산%`);
          totalParams.push(`%익산%`);
          totalParams.push(`%정읍%`);
          totalParams.push(`%김제%`);
          totalParams.push(`%남원%`);
          totalParams.push(`%완주%`);
        }
      }

      const totalQ = await pool.query(
        `
        SELECT
          base_month,
          COALESCE(SUM(total_score), 0)::numeric AS total_value
        FROM sales_records
        ${totalWhere}
        GROUP BY base_month
        ORDER BY base_month ASC
        `,
        totalParams
      );

      const totalMap = {};
      totalQ.rows.forEach(r => {
        totalMap[r.base_month] = Number(r.total_value || 0);
      });

      function calcTrendShare(actual, total) {
        if (!total || Number(total) === 0) return 0;
        return Number(((Number(actual) / Number(total)) * 100).toFixed(1));
      }

      rows = months.map(m => {
        const value = Number(map[m] || 0);
        const total = Number(totalMap[m] || 0);

        return {
          month: m,
          value,
          share_rate: calcTrendShare(value, total)
        };
      });
    }

    return res.json({ ok: true, rows });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "최근 6개월 추이 조회 실패" });
  }
});

/**
 * =========================
 * 실적 필터 조회 (대리점 기준)
 * =========================
 */
app.post("/performance/search", async (req, res) => {
  try {
    const startMonth = toText(req.body.start_month);
    const endMonth = toText(req.body.end_month);
    const region = toText(req.body.region);
    const agencyName = toText(req.body.agency_name);
    const storeName = toText(req.body.store_name);

    if (!startMonth || !endMonth) {
      return res.status(400).json({
        ok: false,
        message: "조회 시작월과 종료월은 필수입니다."
      });
    }

    const params = [startMonth, endMonth];
    let idx = 3;

    let where = `
      WHERE base_month >= $1
        AND base_month <= $2
    `;

    // 🔥 M&S 제한 제거 (핵심!)
    // 기존 is_ms = 'Y' 삭제

    if (region) {
      where += ` AND market ILIKE $${idx}`;
      params.push(`%${region}%`);
      idx++;
    }

    if (agencyName) {
      where += ` AND agency_name ILIKE $${idx}`;
      params.push(`%${agencyName}%`);
      idx++;
    }

    if (storeName) {
      where += ` AND store_name ILIKE $${idx}`;
      params.push(`%${storeName}%`);
      idx++;
    }

    const q = `
      SELECT
        base_month,
        market,
        agency_name,
        COALESCE(SUM(CASE WHEN metric_type = '후불' THEN total_score ELSE 0 END), 0) AS postpaid,
        COALESCE(SUM(CASE WHEN metric_type = '순신규' THEN total_score ELSE 0 END), 0) AS pure_new,
        COALESCE(SUM(CASE WHEN metric_type = '약정갱신' THEN total_score ELSE 0 END), 0) AS renewal,
        COALESCE(SUM(CASE WHEN metric_type = 'MIT' THEN total_score ELSE 0 END), 0) AS mit
      FROM sales_records
      ${where}
      GROUP BY base_month, market, agency_name
      ORDER BY base_month ASC
    `;

    const r = await pool.query(q, params);

    return res.json({
      ok: true,
      rows: r.rows
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "실적 조회 실패" });
  }
});

/**
 * =========================
 * 판매점 상세 조회
 * =========================
 */
app.post("/performance/detail", async (req, res) => {
  try {
    const { base_month, market, agency_name } = req.body;

    if (!base_month || !agency_name) {
      return res.status(400).json({
        ok: false,
        message: "기준월, 대리점 필수"
      });
    }

    const params = [base_month, agency_name];
    let idx = 3;

    let where = `
      WHERE base_month = $1
        AND agency_name = $2
    `;

    if (market) {
      where += ` AND market = $${idx}`;
      params.push(market);
    }

    const q = `
      SELECT
        base_month,
        market,
        agency_name,
        store_code,
        store_name,
        COALESCE(SUM(CASE WHEN metric_type = '후불' THEN total_score ELSE 0 END), 0) AS postpaid,
        COALESCE(SUM(CASE WHEN metric_type = '순신규' THEN total_score ELSE 0 END), 0) AS pure_new
      FROM sales_records
      ${where}
      GROUP BY base_month, market, agency_name, store_code, store_name
      ORDER BY postpaid DESC
    `;

    const r = await pool.query(q, params);

    return res.json({
      ok: true,
      rows: r.rows
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false });
  }
});

/**
 * =========================
 * 판매점 찾기
 * POST /store-master/search
 * body:
 * {
 *   keyword: "명진"
 * }
 * =========================
 */
app.post("/store-master/search", async (req, res) => {
  try {
    const keyword = toText(req.body.keyword);

    if (!keyword) {
      return res.status(400).json({
        ok: false,
        message: "검색어를 입력하세요."
      });
    }

    const r = await pool.query(
      `
      SELECT
        store_code,
        store_name,
        address
      FROM sales_store_master
      WHERE
        store_code ILIKE $1
        OR store_name ILIKE $1
        OR address ILIKE $1
      ORDER BY store_name ASC
      LIMIT 200
      `,
      [`%${keyword}%`]
    );

    return res.json({
      ok: true,
      rows: r.rows
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({
      ok: false,
      message: "판매점 찾기 실패"
    });
  }
});

/**
 * =========================
 * 모델별 회전일
 * POST /inventory/by-model-turnover
 * body: { agency: "광주" }
 * =========================
 */
app.post("/inventory/by-model-turnover", async (req, res) => {
  const snapshotDate = todayKST();
  const { agency } = req.body;

  if (!agency) {
    return res.status(400).json({ ok: false, message: "agency 정보가 없습니다." });
  }

  try {
    const salesAgency = mapCenterToSalesAgency(agency);

    // -------------------------
    // 당월 기준월 / 기준일
    // -------------------------
    const monthQ = await pool.query(
      `
      SELECT MAX(base_month) AS latest_month
      FROM sales_records
      WHERE data_scope = 'daily'
        AND metric_type = '후불'
      `
    );

    const latestMonth = monthQ.rows[0]?.latest_month || null;

    const dateQ = await pool.query(
      `
      SELECT MAX(record_date) AS latest_date
      FROM sales_records
      WHERE data_scope = 'daily'
        AND metric_type = '후불'
      `
    );

    const latestDate = dateQ.rows[0]?.latest_date || null;

    // -------------------------
    // 기준일까지의 영업일 수 (일요일 제외)
    // -------------------------
    let businessDays = 0;

    if (latestDate) {
      const dateObj = new Date(latestDate);
      const year = dateObj.getUTCFullYear();
      const month = dateObj.getUTCMonth() + 1;
      const day = dateObj.getUTCDate();

      for (let d = 1; d <= day; d++) {
        const current = new Date(Date.UTC(year, month - 1, d));
        const weekday = current.getUTCDay(); // 0=일요일

        if (weekday !== 0) {
          businessDays++;
        }
      }
    }

    // -------------------------
    // 재고 모델별 집계
    // summary-extended와 동일하게
    // snapshot_date = todayKST() 기준 사용
    // -------------------------
    const inventoryParams = [snapshotDate];
    let inventoryWhere = `WHERE snapshot_date = $1`;

    if (agency !== "관리자") {
      inventoryWhere += ` AND agency_name = $2`;
      inventoryParams.push(agency);
    }

    const inventoryQ = await pool.query(
      `
      SELECT
        model_name,
        COUNT(*)::int AS qty
      FROM inventory_items
      ${inventoryWhere}
      GROUP BY model_name
      ORDER BY model_name ASC
      `,
      inventoryParams
    );

    // -------------------------
    // 당월 후불판매(모델별)
    // -------------------------
    let salesRows = [];

    if (latestMonth) {
      const salesParams = [latestMonth];
      let salesWhere = `
        WHERE base_month = $1
          AND data_scope = 'daily'
          AND metric_type = '후불'
          AND is_ms = 'Y'
      `;

      if (agency !== "관리자") {
        salesWhere += ` AND agency_name = $2`;
        salesParams.push(salesAgency);
      }

      const salesQ = await pool.query(
        `
        SELECT
          model_name,
          COALESCE(SUM(total_score), 0)::numeric AS sales_qty
        FROM sales_records
        ${salesWhere}
        GROUP BY model_name
        `,
        salesParams
      );

      salesRows = salesQ.rows || [];
    }

    const salesMap = {};
    salesRows.forEach(r => {
      salesMap[r.model_name] = Number(r.sales_qty || 0);
    });

    const rows = (inventoryQ.rows || []).map(r => {
      const qty = Number(r.qty || 0);
      const salesQty = Number(salesMap[r.model_name] || 0);

      let turnoverDays = null;

      if (salesQty > 0 && businessDays > 0) {
        const avgPerDay = salesQty / businessDays;
        turnoverDays = Number((qty / avgPerDay).toFixed(1));
      }

      return {
        model_name: r.model_name || "",
        qty,
        monthly_sales: salesQty,
        turnover_days: turnoverDays,
        base_month: latestMonth,
        business_days: businessDays,
        snapshot_date: snapshotDate
      };
    });

    return res.json({
      ok: true,
      snapshot_date: snapshotDate,
      latest_month: latestMonth,
      business_days: businessDays,
      rows
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({
      ok: false,
      message: "모델별 회전일 조회 실패"
    });
  }
});

app.post("/inventory/aging-summary", async (req, res) => {
  const snapshotDate = todayKST();
  const { agency } = req.body;

  if (!agency) {
    return res.status(400).json({ ok: false, message: "agency 정보가 없습니다." });
  }

  try {
    const params = [snapshotDate];
    let idx = 2;
    let where = `WHERE snapshot_date = $1`;

    if (agency !== "관리자") {
      where += ` AND agency_name = $${idx}`;
      params.push(agency);
      idx++;
    }

    const q = `
      SELECT
        SUM(CASE WHEN aging_days > 360 THEN 1 ELSE 0 END)::int AS over_360,
        SUM(CASE WHEN aging_days > 500 THEN 1 ELSE 0 END)::int AS over_500,
        SUM(CASE WHEN aging_days > 720 THEN 1 ELSE 0 END)::int AS over_720
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
    return res.status(500).json({ ok: false, message: "장기재고 요약 실패" });
  }
});

app.post("/inventory/aging-detail", async (req, res) => {
  const snapshotDate = todayKST();
  const { agency, threshold } = req.body;

  if (!agency) {
    return res.status(400).json({ ok: false, message: "agency 정보가 없습니다." });
  }

  const limitDays = Number(threshold);
  if (![360, 500, 720].includes(limitDays)) {
    return res.status(400).json({ ok: false, message: "threshold 값이 올바르지 않습니다." });
  }

  try {
    const params = [snapshotDate, limitDays];
    let idx = 3;
    let where = `
      WHERE snapshot_date = $1
        AND aging_days > $2
    `;

    if (agency !== "관리자") {
      where += ` AND agency_name = $${idx}`;
      params.push(agency);
      idx++;
    }

    const q = `
  SELECT
    agency_name,
    store_name,
    model_name,
    color,
    serial_no,
    aging_days
  FROM inventory_items
  ${where}
  ORDER BY aging_days DESC, model_name ASC
`;

    const r = await pool.query(q, params);

    return res.json({
      ok: true,
      snapshot_date: snapshotDate,
      threshold: limitDays,
      rows: r.rows
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, message: "장기재고 상세 조회 실패" });
  }
});

app.post("/inventory/agency-summary", async (req, res) => {
  const snapshotDate = todayKST();

  try {
    const q = `
      SELECT
        agency_name,
        COUNT(*)::int AS total_qty
      FROM inventory_items
      WHERE snapshot_date = $1
        AND agency_name IN ('광주', '목포', '순천', '전북', '제주')
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
    `;

    const r = await pool.query(q, [snapshotDate]);

    return res.json({
      ok: true,
      snapshot_date: snapshotDate,
      rows: r.rows
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({
      ok: false,
      message: "센터별 총 재고 조회 실패"
    });
  }
});

app.post("/inventory/agency-detail", async (req, res) => {
  const snapshotDate = todayKST();
  const { agency_name } = req.body;

  if (!agency_name) {
    return res.status(400).json({ ok: false, message: "agency_name 정보가 없습니다." });
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
          ) * 100,
          1
        ) AS warehouse_ratio
      FROM inventory_items
      WHERE snapshot_date = $1
        AND agency_name = $2
      GROUP BY model_name, color
      ORDER BY total_qty DESC, model_name ASC, color ASC
    `;

    const r = await pool.query(q, [snapshotDate, agency_name]);

    return res.json({
      ok: true,
      snapshot_date: snapshotDate,
      agency_name,
      rows: r.rows
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({
      ok: false,
      message: "센터별 총 재고 상세 조회 실패"
    });
  }
});

app.post("/inventory/model-turnover-detail", async (req, res) => {
  const snapshotDate = todayKST();
  const { model_name } = req.body;

  if (!model_name) {
    return res.status(400).json({ ok: false, message: "model_name 정보가 없습니다." });
  }

  try {
    // 당월 기준월 / 기준일
    const monthQ = await pool.query(
      `
      SELECT MAX(base_month) AS latest_month
      FROM sales_records
      WHERE data_scope = 'daily'
        AND metric_type = '후불'
      `
    );

    const latestMonth = monthQ.rows[0]?.latest_month || null;

    const dateQ = await pool.query(
      `
      SELECT MAX(record_date) AS latest_date
      FROM sales_records
      WHERE data_scope = 'daily'
        AND metric_type = '후불'
      `
    );

    const latestDate = dateQ.rows[0]?.latest_date || null;

    let businessDays = 0;

    if (latestDate) {
      const dateObj = new Date(latestDate);
      const year = dateObj.getUTCFullYear();
      const month = dateObj.getUTCMonth() + 1;
      const day = dateObj.getUTCDate();

      for (let d = 1; d <= day; d++) {
        const current = new Date(Date.UTC(year, month - 1, d));
        const weekday = current.getUTCDay(); // 0=일요일
        if (weekday !== 0) {
          businessDays++;
        }
      }
    }

    // 센터별 재고 수량
    const inventoryQ = await pool.query(
      `
      SELECT
        agency_name,
        COUNT(*)::int AS qty
      FROM inventory_items
      WHERE snapshot_date = $1
        AND model_name = $2
        AND agency_name IN ('광주', '목포', '순천', '전북', '제주')
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
      [snapshotDate, model_name]
    );

    // 센터별 당월판매
    let salesRows = [];
    if (latestMonth) {
      const salesQ = await pool.query(
        `
        SELECT
          agency_name,
          COALESCE(SUM(total_score), 0)::numeric AS sales_qty
        FROM sales_records
        WHERE base_month = $1
          AND data_scope = 'daily'
          AND metric_type = '후불'
          AND is_ms = 'Y'
          AND model_name = $2
        GROUP BY agency_name
        ORDER BY
          CASE agency_name
            WHEN 'M&S광주' THEN 1
            WHEN 'M&S목포' THEN 2
            WHEN 'M&S순천' THEN 3
            WHEN 'M&S전북' THEN 4
            WHEN 'M&S제주' THEN 5
            ELSE 99
          END
        `,
        [latestMonth, model_name]
      );

      salesRows = salesQ.rows || [];
    }

    const salesMap = {
      "광주": 0,
      "목포": 0,
      "순천": 0,
      "전북": 0,
      "제주": 0
    };

    salesRows.forEach(r => {
      const agency = String(r.agency_name || "");
      if (agency === "M&S광주") salesMap["광주"] = Number(r.sales_qty || 0);
      if (agency === "M&S목포") salesMap["목포"] = Number(r.sales_qty || 0);
      if (agency === "M&S순천") salesMap["순천"] = Number(r.sales_qty || 0);
      if (agency === "M&S전북") salesMap["전북"] = Number(r.sales_qty || 0);
      if (agency === "M&S제주") salesMap["제주"] = Number(r.sales_qty || 0);
    });

    const order = ["광주", "목포", "순천", "전북", "제주"];

    const inventoryMap = {
      "광주": 0,
      "목포": 0,
      "순천": 0,
      "전북": 0,
      "제주": 0
    };

    (inventoryQ.rows || []).forEach(r => {
      inventoryMap[r.agency_name] = Number(r.qty || 0);
    });

    const rows = order.map(name => {
      const qty = Number(inventoryMap[name] || 0);
      const salesQty = Number(salesMap[name] || 0);

      let turnoverDays = null;
      if (salesQty > 0 && businessDays > 0) {
        const avgPerDay = salesQty / businessDays;
        turnoverDays = Number((qty / avgPerDay).toFixed(1));
      }

      return {
        agency_name: name,
        qty,
        monthly_sales: salesQty,
        turnover_days: turnoverDays
      };
    });

    return res.json({
      ok: true,
      model_name,
      business_days: businessDays,
      rows
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({
      ok: false,
      message: "모델별 센터 회전일 상세 조회 실패"
    });
  }
});

app.post("/inventory/model-turnover-store-detail", async (req, res) => {
  const snapshotDate = todayKST();
  const { model_name, agency } = req.body;

  if (!model_name) {
    return res.status(400).json({ ok: false, message: "model_name 정보가 없습니다." });
  }

  if (!agency || agency === "관리자") {
    return res.status(400).json({ ok: false, message: "센터 agency 정보가 없습니다." });
  }

  try {
    const salesAgency = mapCenterToSalesAgency(agency);

    // 당월 기준월
    const monthQ = await pool.query(
      `
      SELECT MAX(base_month) AS latest_month
      FROM sales_records
      WHERE data_scope = 'daily'
        AND metric_type = '후불'
      `
    );

    const latestMonth = monthQ.rows[0]?.latest_month || null;

    // 판매점별 재고
    const inventoryQ = await pool.query(
      `
      SELECT
        store_name,
        COUNT(*)::int AS qty
      FROM inventory_items
      WHERE snapshot_date = $1
        AND agency_name = $2
        AND model_name = $3
      GROUP BY store_name
      ORDER BY qty DESC, store_name ASC
      `,
      [snapshotDate, agency, model_name]
    );

    // 판매점별 당월판매
    let salesRows = [];
    if (latestMonth) {
      const salesQ = await pool.query(
        `
        SELECT
          store_name,
          COALESCE(SUM(total_score), 0)::numeric AS sales_qty
        FROM sales_records
        WHERE base_month = $1
          AND data_scope = 'daily'
          AND metric_type = '후불'
          AND is_ms = 'Y'
          AND agency_name = $2
          AND model_name = $3
        GROUP BY store_name
        ORDER BY sales_qty DESC, store_name ASC
        `,
        [latestMonth, salesAgency, model_name]
      );

      salesRows = salesQ.rows || [];
    }

    const salesMap = {};
    salesRows.forEach(r => {
      salesMap[r.store_name] = Number(r.sales_qty || 0);
    });

    const inventoryMap = {};
    (inventoryQ.rows || []).forEach(r => {
      inventoryMap[r.store_name] = Number(r.qty || 0);
    });

    const storeNames = Array.from(
      new Set([
        ...Object.keys(inventoryMap),
        ...Object.keys(salesMap)
      ])
    ).sort((a, b) => a.localeCompare(b, "ko"));

    const rows = storeNames.map(name => ({
      store_name: name,
      qty: Number(inventoryMap[name] || 0),
      monthly_sales: Number(salesMap[name] || 0)
    }));

    // 기본은 재고 많은 순
    rows.sort((a, b) => {
      if (b.qty !== a.qty) return b.qty - a.qty;
      if (b.monthly_sales !== a.monthly_sales) return b.monthly_sales - a.monthly_sales;
      return String(a.store_name || "").localeCompare(String(b.store_name || ""), "ko");
    });

    return res.json({
      ok: true,
      model_name,
      agency,
      rows
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({
      ok: false,
      message: "판매점별 모델 상세 조회 실패"
    });
  }
});

app.listen(PORT, () => {
  console.log("🚀 Backend running on port", PORT);
});