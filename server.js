const express = require('express');
const http = require('http');
const socketIO = require('socket.io');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();
const server = http.createServer(app);
const io = socketIO(server);

const port = 9999;

app.get('/', (req, res) => {
    res.send('Server is running.');
});

io.on('connection', (socket) => {
    console.log('Client connected');

    socket.on('qrcode', (data) => {
        console.log('Received QR code:', data);

        // Đọc tệp Excel hiện có và thêm dữ liệu mới
        updateExcel(data);

        // Gửi thông báo cho client rằng dữ liệu đã được lưu vào tệp Excel
        socket.emit('qrcode_saved', 'QR code data has been saved to Excel');
    });

    socket.on('image', (data) => {
        // ... (xử lý lưu hình ảnh)
    });

    socket.on('disconnect', () => {
        console.log('Client disconnected');
    });
});

server.listen(port, () => {
    console.log(`Server is running on port ${port}`);
});

// Hàm cập nhật dữ liệu vào tệp Excel
function updateExcel(data) {
    const excelFilePath = path.join(__dirname, 'data_cccd.xlsx');
    const workbook = new ExcelJS.Workbook();

    // Đọc tệp Excel hiện có
    workbook.xlsx.readFile(excelFilePath)
        .then(() => {
            const worksheet = workbook.getWorksheet('Sheet1');

            // Thêm dòng dữ liệu mới vào tệp Excel, bắt đầu từ hàng thứ 2
            const newRow = worksheet.addRow([
                data.CCCD,
                data.CMT,
                data.HOTEN,
                data.NGAYSINH,
                data.DIACHI,
                data.NGAYCAP,
            ]);

            // Ghi tệp Excel lại
            return workbook.xlsx.writeFile(excelFilePath);
        })
        .then(() => {
            console.log('QR code data updated in Excel');
        })
        .catch((err) => {
            console.error('Error updating QR code data in Excel:', err);
        });
}