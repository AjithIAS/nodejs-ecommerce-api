const express = require("express");
const router = express.Router();
const mongoose = require("mongoose");
const multer = require("multer");
const path = require("path");

const Product = require("../models/product");

const excelToJson = require("convert-excel-to-json");
const fs = require("fs");

global._basedir = __dirname;
global.appRoot = path.resolve(__dirname);

// const storage = multer.diskStorage({
//   destination: (req, file, cb) => {
//     cb(null, path.join(appRoot, "uploads/public"));
//   },
//   filename: (req, file, cb) => {
//     cb(null, file.originalname);
//   },
// });
const FILE_TYPE_MAP = {
  "application/vnd.ms-excel": "xls",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
};

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    const isValid = FILE_TYPE_MAP[file.mimetype];
    let uploadError = new Error("Invalid Image Type");
    if (isValid) {
      uploadError = null;
    }
    cb(uploadError, "public/uploads");
  },
  filename: function (req, file, cb) {
    const fileName = file.originalname.split("").join("-");
    const extension = FILE_TYPE_MAP[file.mimetype];
    cb(null, `${fileName}-${Date.now()}.${extension}`);
  },
});

const upload = multer({ storage: storage });

// Express upload REST API
router.post("/", upload.single("uploadfile"), (req, res) => {
  try {
    const filePath = req.file;
    // console.log("filePath");
    // console.log(filePath);
    importExcelData2MongoDB(filePath);
    res.json({
      msg: "File Uploaded",
      file: req.file?.filename,
    });
  } catch (err) {
    console.log("Error uploading file:", err);
    res.status(500).json({ error: "Failed to upload file" });
  }
});

// Import to MongoDB
async function importExcelData2MongoDB(filePath) {
  try {
    console.log(filePath);
    const excelData = await excelToJson({
      sourceFile: filePath?.path,
      sheets: [
        {
          name: "Sheet1",
          header: {
            rows: 1, // no/s. of the row that has headers in your excel sheet
          },
          columnToKey: {
            A: "name", //these attributes should be similar to the ones in your excel file
            B: "description",
            C: "richDescription",
            D: "image",
            E: "brand",
            F: "price",
            G: "category",
            H: "countInStock",
            I: "isFeatured",
          },
        },
      ],
    });

    const sheetName = "Sheet1"; // Update the sheet name
    console.log(excelData[sheetName]);
    if (!excelData[sheetName] || excelData[sheetName].length === 0) {
      throw new Error(`No data found in the '${sheetName}' sheet.`);
    } else {
      const result = await Product.insertMany(excelData[sheetName]);
      console.log("===============================");
      console.log(result.length);
      //   const result = await collection.insertMany(excelData[sheetName]);
      console.log("Number of documents inserted:", result.length);

      fs.unlinkSync(filePath?.path);
    }
  } catch (err) {
    console.log("Error importing data to MongoDB:", err);
    throw err;
  }
}

module.exports = router;
