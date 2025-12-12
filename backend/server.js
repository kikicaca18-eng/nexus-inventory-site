// server.js
import express from "express";
import cors from "cors";
import multer from "multer";
import xlsx from "xlsx";
import fs from "fs";

const app = express();
app.use(cors());
app.use(express.json());

// 메모리 저장 방식(파일로 저장하지 않음)
const upload = multer({ storage: multer.memoryStorage() });

// 엑셀 데이터 메모리에 저장(임시 DB 역할)
let inventoryData = [];

// 센터별 비밀번호 (너 회사 정책 맞게 변경 가능)
const centerPasswords = {
  "광주": "gwangju123",
  "목포": "mokpo123",
  "순천": "suncheon123",
  "전북": "jeonbuk123",
  "제주": "jeju123"
};

// 로그인 처리
app.post("/login", (req, res) => {
  const { center, password } = req.body;

  if (!centerPasswords[center]) {
    return res.status(400).json({ success: false, message: "존재하지 않는 센터입니다." });
  }

  if (centerPasswords[center] !== password) {
    return res.status(401).json({ success: false, message: "비밀번호가 틀렸습니다." });
  }

  return res.json({ success: true, center });
});

// 엑셀 업로드 (관리자만 접근)
app.post("/upload-excel", upload.single("file"), (req, res) => {
  try {
    const buffer = req.file.buffer;

    const workbook = xlsx.read(buffer, { type: "buffer" });
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = xlsx.utils.sheet_to_json(worksheet);

    inventoryData = jsonData;

    return res.json({ success: true, rows: inventoryData.length });
  } catch (error) {
    console.error(error);
    return res.status(500).json({ success: false, message: "엑셀 처리 중 오류 발생" });
  }
});

// 센터별 재고 조회
app.get("/search", (req, res) => {
  const { center, keyword } = req.query;

  if (!center) {
    return res.status(400).json({ success: false, message: "센터 정보 필요" });
  }

  let results = inventoryData.filter(row => row["센터"] === center);

  if (keyword && keyword.trim() !== "") {
    results = results.filter(row =>
      JSON.stringify(row).includes(keyword)
    );
  }

  return res.json({ success: true, results });
});

// 서버 실행
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`✅ Backend server running on port ${PORT}`);
});
