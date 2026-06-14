            const express = require("express");
            const multer = require("multer");
            const XLSX = require("xlsx");
            const fs = require("fs");
            const path = require("path");

            const app = express();

            // Configure upload folder
            const storage = multer.diskStorage({
                destination: function (req, file, cb) {
                    cb(null, "uploads/");
                },
                filename: function (req, file, cb) {
                    cb(null, file.originalname);
                }
            });

            const upload = multer({ storage: storage });

            // Serve website files
            app.use(express.static("public"));

            // Upload route
            app.post("/upload", upload.single("excelFile"), (req, res) => {

                const workbook = XLSX.readFile(req.file.path);

                const sheetName = workbook.SheetNames[0];

                const worksheet = workbook.Sheets[sheetName];

                const data = XLSX.utils.sheet_to_json(worksheet, {
                defval: ""
            });

                let txtContent = "";

            data.forEach((row, rowIndex) => {

                const values = Object.values(row);

                // Keep only columns A to U
            const limitedValues = values.slice(0, 21);

                const hasData = limitedValues.some(value =>
                    value !== null &&
                    value !== undefined &&
                    value !== ""
                );

                if (hasData) {

            let filteredValues = [...limitedValues];

            // Keep R:U for the first data row (Excel row 2)
            // Remove R:U from Excel row 3 and below
            if (rowIndex >= 1) {
                filteredValues.splice(17, 4);
            }

                    const formattedValues = [];

            filteredValues.forEach((value, index) => {

                // Skip R,S,T,U for row 3 and below
                if (rowIndex >= 1 && index >= 17 && index <= 20) {
                    return;
                }

                // A and B
                if (index === 0 || index === 1) {
                    formattedValues.push(value);
                    return;
                }

            // U2 only - no 2 decimal formatting
            if (rowIndex === 0 && index === 20) {
                formattedValues.push(`"${value}"`);
                return;
            }

            // Row 2 - keep columns R,S,T,U exactly as they are
            if (rowIndex === 0 && index >= 17 && index <= 20) {
                formattedValues.push(`"${value}"`);
                return;
            }

            // Other numeric columns
            if (typeof value === "number") {
                formattedValues.push(`"${value.toFixed(2)}"`);
                return;
            }

                formattedValues.push(`"${value}"`);
            });

            txtContent += formattedValues.join(",") + "\n";
                }
            });

            const sheetData = XLSX.utils.sheet_to_json(worksheet, {
                header: 1
            });

            // Validate template
            if (
                sheetData[0][2] !== "SUPPLIER TIN" ||
                sheetData[0][3] !== "REGISTER NAME" ||
                sheetData[0][15] !== "COMPANY TIN" ||
                sheetData[0][19] !== "MONTH"
            ) {
                return res.send(`
                    <h2>Invalid Template</h2>
                    <p>Please use the official Purchases Template - Vat Relief.xlsx file.</p>
                    <a href="/">Go Back</a>
                `);
            }

            const c2 = sheetData[1][2];   // Row 2, Column C
            const t2 = sheetData[1][19];  // Row 2, Column T

            const date = new Date(t2);

            const month = String(date.getMonth() + 1).padStart(2, "0");
            const year = date.getFullYear();

            const datFileName = `${c2}P${month}${year}.dat`;

            fs.writeFileSync(datFileName, txtContent);

            const filePath = path.join(__dirname, datFileName);

        res.redirect(`/success.html?file=${datFileName}`);

            });

            app.get("/download/:filename", (req, res) => {

        const filePath = path.join(__dirname, req.params.filename);

        res.download(filePath);

             });

    const PORT = process.env.PORT || 3000;

    app.listen(PORT, () => {
        console.log(`Server running on port ${PORT}`);
    });