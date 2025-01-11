const express = require("express");
const multer = require("multer");
const exceljs = require("exceljs");

const app = express();
const port = 8683;

// Cấu hình multer để lưu file tải lên vào bộ nhớ
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.post("/upload", upload.single("excelFile"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "Không có file nào được tải lên." });
    }

    const workbook = new exceljs.Workbook();
    // Đọc file từ bộ nhớ
    await workbook.xlsx.load(req.file.buffer);

    // Lấy worksheet đầu tiên
    const worksheet = workbook.getWorksheet(1);

    const data = {};
    // Bắt đầu đọc dữ liệu từ hàng thứ 9 (bỏ qua hàng tiêu đề)
    let rowNumber = 2;
    let currentRow = worksheet.getRow(rowNumber);
    const Temp_user = [];

    const Category = 2; // Cột B
    const Item = 3; // Cột C
    const Guideline = 5; // Cột E


    while (currentRow.getCell(1).value || currentRow.getCell(2).value || currentRow.getCell(3).value || currentRow.getCell(4).value || currentRow.getCell(5).value || currentRow.getCell(6).value || currentRow.getCell(7).value) {

      const studentData = {
        student_id: currentRow.getCell(1).value, // Cột A
        name: currentRow.getCell(2).value,       // Cột B
        class: currentRow.getCell(3).value,      // Cột C
        faculty: currentRow.getCell(4).value,    // Cột D
        birthday: currentRow.getCell(5).value,   // Cột E
        phone_number: currentRow.getCell(6).value, // Cột F
        gender: currentRow.getCell(7).value,     // Cột G
      };
      Temp_user.push(studentData);
      rowNumber++;
      currentRow = worksheet.getRow(rowNumber);
    }
  

    res.json({ Temp_user });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Đã có lỗi xảy ra khi xử lý file." });
  }
});

app.listen(port, () => {
  console.log(`Server đang lắng nghe tại cổng ${port}`);
});
