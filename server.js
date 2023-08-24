const express = require('express');
const http = require('http');
const socketIO = require('socket.io');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { google } = require('googleapis');

const app = express();
const server = http.createServer(app);
const io = socketIO(server);

const port = 9999;

app.get('/', (req, res) => {
    res.send('Server is running.');
});


const auth = new google.auth.GoogleAuth({
    credentials: {
        "type": "service_account",
        "project_id": "data-396609",
        "private_key_id": "b7938d157123237d2cfc7638fb690806a6da00f2",
        "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQC8V5ct6XDKYkgL\n14OZdZaOw4OewO7JkgxdHVpPaG7MUM/TEHKcaQrOp0vy1A6o4+6oRsywXNSDGoRI\n/YAZ0ZzAA99WpT7ISCnrjEGZ8HgI+3vqmcfDVWmdIeg50h/EJPVMjCkZVtVOWjKu\nVDhWsrZFW0AHzHpgLZcvz97lZPdYaw77bfYS4PUq6ytZLcEtGv1GYPsTAMXrWfuf\nh90hbpbApPSME8amhRWIOWHUn0T35RkhR0xFD2noMqDNC94ERwbYch3vYqA14kUT\nMkyWd1cbgkXr8zApo0gc+aBwwUp6fGyRQwQgs+5KpxF0lS8/Txq9SfzHYq66g+XP\n9n3daLRPAgMBAAECggEAA9wLpg41mIAhF9UHxlpJho1tIhC0E/9hL3tNgXkTZC53\nw8x/fMMOTKPTA5vi3QUmu94PjcGpPTKK6XAAcwb/nlO1/PGP8sy/xyrtoTjzTMPS\ni0Exfcg/TNAIET9EFWHogPQSV4mW/28LwNTK/sPUxmETU9WV4Giuw3UQDUVfaDlo\nZcrbgW/XEBNh/gdYuksnhR+U69KQQcYZmaChBg+wY3gEkYb/ZmsKUa3Z3usb2p1i\n7IhZ8iaQfulJtb7hd2z9UzzytqopRY1SwtxNzV7phncBFgRfowILVuAvSs1Yh91T\nzQVZcufGb5xKMj2O1jjYSEZgOeO2/FiHUPKF7FemaQKBgQDpxOqW695STXZtSXHc\ndM6wJdH6AsmFYpXCYJ5F1tsiNfBJJ7lLT2+jMM9maaFptZqUskRazJ/dRwlooIhW\nhqVH3eG+9HsEz2VwGM5YU3KzBh8q2hs9AVr3NieWNlCTR4JmIPgBR9Q+EqCavjfh\nW4fBh1bX8cZ/dT4yNAjYvZwqvQKBgQDOQMIHOKwAF+QM5h1Yfl9mC5I8dYg/5ekh\n2ggoB4iBF6+EBVFYkYUCHvVo4GqNWU9e5bbLoeSn/OO3T0jHShP5Q7WM/itJ3thE\n1fhRugSpKInjBiDU94uJuJdML0NI+1fffHbDg28ona1KlBmewPTPXv9wRMBtBO+b\nT1SaealR+wKBgHm8y3HW2UtA/chB9CKbTbubpnKtGub0hQrZp/K0xh9VuZFPN4aJ\nkpiIZaluntle8mY3Q7OJVkM0qCitWPK+YbpASTxZMus5Whj7QhHrOxMRwA9fz8mA\nOC//KrRmCqX4Gmc3ChAYqOW+a5bKMm2Qbe0Rnt8MEJP1qXMZd/XvIDF9AoGBAMAT\nPp5LAKL1nMMGab3Hsj/t9rmnGsOm8H099uqQWWcfD6z65s58dkmLWy/YDmKkEW5m\nrtzkX3Sx5b8IbtZo/kDb9W7gJKAej3lLan1xpnWB8ycgxeKOxbvz07J3MUn+B89w\nsYlSFWrVrFQPp+xX9aRI68k5vZnJRvpz3m4dbrmRAoGACUQZAqgg4L5/yGdI4x1k\nBoBFZNQEk3wSAiCcloimeThUqcnd0hC51PNrsJJJA6U+bzd93qIkct1WjG3u0Usc\ngNdsgqJsfHxQ0KmxpfaSWM/0wroyxz50T6HAkbvZECClg6NnNG8/hxLJF0em6yZN\n0bDac4Oq1Srcx7N1ZS6t/wY=\n-----END PRIVATE KEY-----\n",
        "client_email": "datacccd@data-396609.iam.gserviceaccount.com",
        "client_id": "117314679268836226436",
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/datacccd%40data-396609.iam.gserviceaccount.com",
        "universe_domain": "googleapis.com"
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

io.on('connection', (socket) => {
    console.log('Client connected');

    socket.on('qrcode', (data) => {
        console.log('Received QR code:', data);

        updateExcel(data);


        updateGoogleSheet(data);


        socket.emit('qrcode_saved', 'QR code data has been saved.');
    });

    socket.on('image', (data) => {

    });

    socket.on('disconnect', () => {
        console.log('Client disconnected');
    });
});

server.listen(port, () => {
    console.log(`Server is running on port ${port}`);
});


function updateExcel(data) {
    const excelFilePath = path.join(__dirname, 'data_cccd.xlsx');
    const workbook = new ExcelJS.Workbook();

    workbook.xlsx.readFile(excelFilePath)
        .then(() => {
            const worksheet = workbook.getWorksheet('Sheet1');

            const newRow = worksheet.addRow([
                data.CCCD,
                data.CMT,
                data.HOTEN,
                data.NGAYSINH,
                data.DIACHI,
                data.NGAYCAP,
            ]);

            return workbook.xlsx.writeFile(excelFilePath);
        })
        .then(() => {
            console.log('QR code data updated in Excel');
        })
        .catch((err) => {
            console.error('Error updating QR code data in Excel:', err);
        });
}


async function updateGoogleSheet(data) {
    const sheets = google.sheets('v4');
    const spreadsheetId = '1m-Y0-nny4kbyVlGMiy8gtQzsN_POo7MH3aeSnTLhmF8';

    const values = [
        [
            data.CCCD,
            data.CMT,
            data.HOTEN,
            data.NGAYSINH,
            data.DIACHI,
            data.NGAYCAP,
        ],
    ];

    try {
        const response = await sheets.spreadsheets.values.append({
            auth,
            spreadsheetId,
            range: 'cccd',
            valueInputOption: 'RAW',
            insertDataOption: 'INSERT_ROWS',
            resource: {
                values,
            },
        });

        console.log('QR code data updated in Google Sheets');
    } catch (error) {
        console.error('Error updating QR code data in Google Sheets:', error);
    }
}
