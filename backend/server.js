import express from "express";
import cors from "cors";
import multer from "multer";
import xlsx from "xlsx";
import fs from "fs";

const app = express();
app.use(cors());
app.use(express.json());

// 메모리에 저장되는 재고 데이터
let inventoryData = [];

// Multer 설정 (파일 업로드)
const upload = multer({ dest: "uploads/" });

// 엑셀 업로드 API
app.post("/upload", upload.single("file"), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ message: "파일이 없습니다." });
        }

        // 엑셀 파일 읽기
        const workbook = xlsx.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        inventoryData = xlsx.utils.sheet_to_json(sheet);

        // 임시 파일 삭제
        fs.unlinkSync(req.file.path);

        res.json({ message: "업로드 성공", count: inventoryData.length });
    } catch (error) {
        console.error("엑셀 업로드 오류:", error);
        res.status(500).json({ message: "서버 오류" });
    }
});

// 센터별 + 키워드 검색 API
app.post("/", (req, res) => {
    const { center, keyword } = req.body;

    if (!center) return res.status(400).json({ results: [] });

    const results = inventoryData.filter(row => {
        const matchCenter = row["센터"] === center;
        const matchKeyword =
            !keyword ||
            JSON.stringify(row).toLowerCase().includes(keyword.toLowerCase());

        return matchCenter && matchKeyword;
    });

    res.json({ results });
});

// 서버 실행
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`백엔드 서버 실행됨: ${PORT}`));
