import path from "path";
import fs from "fs-extra";
import fss from "fs";
import express from "express";
import bodyParser from "body-parser";
import multer from "multer";
("use strict");
import excelToJson from "convert-excel-to-json";
import mongoXlsx from "mongo-xlsx";
import mongoose from "mongoose";
import pkg from "lodash";

const app = express();
// require('dotenv').config;
import "dotenv/config";
import connectDB from "./connectMongo.js";
import { log } from "console";
connectDB();

const assemblyMainSchema = new mongoose.Schema({
  DATE: String,
  SHIFT: String,
  PRODUCT: String,
  TYPE: String,
  ACTIVITY: String,
  OPERATOR: String,
  QTY: Number
});
const Main = mongoose.model("Main", assemblyMainSchema);

const assemblyOthersSchema = new mongoose.Schema({
  DATE: String,
  SHIFT: String,
  PRODUCT: String,
  TYPE: String,
  ACTIVITY: String,
  OPERATOR: String,
  QTY: Number
});
const Other = mongoose.model("Other", assemblyOthersSchema);

const smtSchema = new mongoose.Schema({
  DATE: String,
  SHIFT: String,
  OPERATOR: String,
  LINE: String,
  PRODUCT: String,
  CALIBER: String,
  QTY: Number
});
const SMT = mongoose.model("SMT", smtSchema);

const AoiSchema = new mongoose.Schema({
  DATE: String,
  SHIFT: String,
  OPERATOR: String,
  AOI_LINE: String,
  PRODUCT: String,
  CALIBER: String,
  QTY: Number
});
const AOI = mongoose.model("AOI", AoiSchema);

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    return cb(null, "./uploads");
  },
  filename: function (req, file, cb) {
    return cb(null, "excelsheet.xlsx");
  },
});
const upload = multer({ storage: storage });

app.use(express.static("public"));

app.set("view engine", "ejs");
app.set("views", path.resolve("./views"));

app.use(express.urlencoded({ extended: false }));
app.use(bodyParser.urlencoded({ extended: true }));

app.get("/", (req, res) => {
  res.render("home");
});

app.post("/", upload.single("excelsheet"), (req, res) => {
  // let doc = parser.parseXls2Json(req.file.path);
  console.log("Uploaded");
  res.redirect("/");
});

app.post("/sheetSelector",(req,res)=>{
  Main.deleteMany()
    .then((items) => {
      console.log("Deleted Main", items);
    })
    .catch((error) => {
      console.log(error);
    });

    Other.deleteMany()
    .then((items) => {
      console.log("Deleted Other", items);
    })
    .catch((error) => {
      console.log(error);
    });

    SMT.deleteMany()
    .then((items) => {
      console.log("Deleted SMT", items);
    })
    .catch((error) => {
      console.log(error);
    });
    AOI.deleteMany()
    .then((items) => {
      console.log("Deleted AOI", items);
    })
    .catch((error) => {
      console.log(error);
    });

let mainProducts=["APFC","Danlaw","Dellorto","Dura","Emerald","IJL230B","IJL660A","Kongsberg","LEDs","S3RP","Tyco"];
let otherProducts=["Danlaw","Dellorto","Dura","IJL230B","IJL660A","Kongsberg","S3RP","Tyco"];
let otherActivities = ["Cleaning","Routing","Singulation"];
let smtLines = ["Line 1","Line 2","Line 3","Line 4","Line 5"];
let smtProducts = ["APFC","Danlaw","Dellorto","Dura","Emerald","IJL660A","KA","LEDs","Neutral_4P","S3RP","Tyco"];

  const sheet = req.body.SHEET;
  let sheetFilterOptions;
  let sheetFilterName;
  let sheetNo;
  // console.log(sheet);
  const doc = excelToJson({
    sourceFile: "./uploads/excelsheet.xlsx",
    header: {
      rows: 1,
    },
    sheets: [sheet]
  });
  // console.log(Object.keys(doc["Assembly_Main"]).length);
  // console.log(doc[sheet][0]);
  // console.log(doc["Assembly_Main"][0].G);
  if (sheet == "Assembly_Main") {
    sheetFilterOptions = ["DATE","SHIFT","PRODUCT","TYPE","ACTIVITY","OPERATOR"];
    sheetFilterName = ["Date","Shift","Product","Type","Activity","Operator"];
    sheetNo = "1";
     Main.find()
    .then(function (foundItems) {
      for (let i = 0; i < Object.keys(doc[sheet]).length; i++) {
        let mains = new Main({
          DATE: doc[sheet][i].A,
          SHIFT: doc[sheet][i].B,
          PRODUCT: doc[sheet][i].C,
          TYPE: doc[sheet][i].D,
          ACTIVITY: doc[sheet][i].E,
          OPERATOR: doc[sheet][i].F,
          QTY: doc[sheet][i].G,
        });
        mains.save();
        //  console.log(doc[0][i].name);
      }
      console.log("Success");
    })
    .catch(function (error) {
      console.log(error);
    });
  } else if (sheet == "Assembly_Others") {
    sheetFilterOptions = ["DATE","SHIFT","PRODUCT","TYPE","ACTIVITY","OPERATOR"];
    sheetFilterName = ["Date","Shift","Product","Type","Activity","Operator"];
    sheetNo = "2";

    Other.find()
    .then(function (foundItems) {
      for (let i = 0; i < Object.keys(doc[sheet]).length; i++) {
        let others = new Other({
          DATE: doc[sheet][i].A,
          SHIFT: doc[sheet][i].B,
          PRODUCT: doc[sheet][i].C,
          TYPE: doc[sheet][i].D,
          ACTIVITY: doc[sheet][i].E,
          OPERATOR: doc[sheet][i].F,
          QTY: doc[sheet][i].G,
        });
        others.save();
        //  console.log(doc[0][i].name);
      }
      console.log("Success");
    })
    .catch(function (error) {
      console.log(error);
    });
  } else if (sheet == "SMT") {
    sheetFilterOptions = ["DATE","SHIFT","OPERATOR","LINE","PRODUCT","CALIBER"];
    sheetFilterName = ["Date","Shift","Operator","Line","Product","Caliber"];
    sheetNo = "3";

    SMT.find()
    .then(function (foundItems) {
      for (let i = 0; i < Object.keys(doc[sheet]).length; i++) {
        let smts = new SMT({
          DATE: doc[sheet][i].A,
          SHIFT: doc[sheet][i].B,
          OPERATOR: doc[sheet][i].C,
          LINE: doc[sheet][i].D,
          PRODUCT: doc[sheet][i].E,
          CALIBER: doc[sheet][i].F,
          QTY: doc[sheet][i].G,
        });
        smts.save();
        //  console.log(doc[0][i].name);
      }
      console.log("Success");
    })
    .catch(function (error) {
      console.log(error);
    });
  } else if (sheet == "AOI") {
    sheetFilterOptions = ["DATE","SHIFT","OPERATOR","AOI_LINE","PRODUCT","CALIBER"];
    sheetFilterName = ["Date","Shift","Operator","AOI_Line","Product","Caliber"];
    sheetNo = "4";

    AOI.find()
    .then(function (foundItems) {
      for (let i = 0; i < Object.keys(doc[sheet]).length; i++) {
        let aois = new AOI({
          DATE: doc[sheet][i].A,
          SHIFT: doc[sheet][i].B,
          OPERATOR: doc[sheet][i].C,
          AOI_LINE: doc[sheet][i].D,
          PRODUCT: doc[sheet][i].E,
          CALIBER: doc[sheet][i].F,
          QTY: doc[sheet][i].G,
        });
        aois.save();
        //  console.log(doc[0][i].name);
      }
      console.log("Success");
    })
    .catch(function (error) {
      console.log(error);
    });
  }
  res.render("filterpage",{
    sheetFilterName:sheetFilterName,
    sheetFilterOptions:sheetFilterOptions,
    sheetNo:sheetNo,
    sheet:sheet,
    mainProducts:mainProducts,
    otherProducts:otherProducts,
    otherActivities:otherActivities,
    smtLines:smtLines,
    smtProducts:smtProducts
  })
});

app.post("/filterby", (req, res) => {
  const d1 = req.body.one;
  const d2 = req.body.two;
  const d3 = req.body.three;
  const d4 = req.body.four;
  const d5 = req.body.five;
  const d6 = req.body.six;
 
  const v1 = req.body.hid1;
  const v2 = req.body.hid2;
  const v3 = req.body.hid3;
  const v4 = req.body.hid4;
  const v5 = req.body.hid5;
  const v6 = req.body.hid6;

  const sheetNumber = req.body.sheetNumber;
  let sheetFilterName;
  let sheetFilterOptions;

  const pathToFile2 = "./downloads/file.xlsx";
  if (fss.existsSync(pathToFile2)) {
    fs.unlink("./downloads/file.xlsx", function (err) {
      if (err) console.log(err);
      // if no error, file has been deleted successfully
      console.log("File deleted!");
    });
  }
  const obj = {};
  obj[v1]= d1;
  obj[v2]= d2;
  obj[v3]= d3;
  obj[v4]= d4;
  obj[v5]= d5;
  obj[v6]= d6;

  // console.log(obj);
  const parameters = {};
  for (let i in obj) {
    if (obj[i] != undefined) {
      const key = i;
        const value = obj[i];
        // console.log(key, value);
        parameters[key] = value;
      
    }
  }
  // console.log(parameters,sheetNumber);

  // console.log(obj);
  if (sheetNumber=="1") {
    sheetFilterOptions = ["DATE","SHIFT","PRODUCT","TYPE","ACTIVITY","OPERATOR"];
    sheetFilterName = ["Date","Shift","Product","Type","Activity","Operator"];

    async function sheet1() {
      try {
        const totems = await Main.find(parameters).select("-_id -__v");
        const number = await Main.find(parameters).exec();
        // console.log(number);
        const model = mongoXlsx.buildDynamicModel(totems);
        mongoXlsx.mongoData2Xlsx(totems, model, function (err, totems) {
          console.log("File saved at:", totems.fullPath);
          fs.move(totems.fullPath, "./downloads/file.xlsx")
            .then(() => {
              console.log("successfully shifted to downloads!");
            })
            .catch((err) => {
              console.error(err);
            });
        });
        res.render("download", {
          number: number,
          sheetNumber:sheetNumber,
          sheetFilterName:sheetFilterName,
          title:"Assembly_Main"
        });
      } catch (error) {
        console.log(error);
      }
    }
    const user = sheet1();
    
  } else if (sheetNumber=="2") {
    sheetFilterOptions = ["DATE","SHIFT","PRODUCT","TYPE","ACTIVITY","OPERATOR"];
    sheetFilterName = ["Date","Shift","Product","Type","Activity","Operator"];
    async function sheet2() {
      try {
        const totems = await Other.find(parameters).select("-_id -__v");
        const number = await Other.find(parameters).exec();
        const model = mongoXlsx.buildDynamicModel(totems);
        mongoXlsx.mongoData2Xlsx(totems, model, function (err, totems) {
          console.log("File saved at:", totems.fullPath);
          fs.move(totems.fullPath, "./downloads/file.xlsx")
            .then(() => {
              console.log("successfully shifted to downloads!");
            })
            .catch((err) => {
              console.error(err);
            });
        });
        res.render("download", {
          number: number,
          sheetNumber:sheetNumber,
          sheetFilterName:sheetFilterName,
          title:"Assembly_Others"
        });
      } catch (error) {
        console.log(error);
      }
    }
    const user = sheet2();
    
  } else if (sheetNumber=="3") {
    sheetFilterOptions = ["DATE","SHIFT","OPERATOR","LINE","PRODUCT","CALIBER"];
    sheetFilterName = ["Date","Shift","Operator","Line","Product","Caliber"];
    async function sheet3() {
      try {
        const totems = await SMT.find(parameters).select("-_id -__v");
        const number = await SMT.find(parameters).exec();
        const model = mongoXlsx.buildDynamicModel(totems);
        mongoXlsx.mongoData2Xlsx(totems, model, function (err, totems) {
          console.log("File saved at:", totems.fullPath);
          fs.move(totems.fullPath, "./downloads/file.xlsx")
            .then(() => {
              console.log("successfully shifted to downloads!");
            })
            .catch((err) => {
              console.error(err);
            });
        });
        res.render("download", {
          number: number,
          sheetNumber:sheetNumber,
          sheetFilterName:sheetFilterName,
          title:"SMT"
        });
      } catch (error) {
        console.log(error);
      }
    }
    const user = sheet3();
  } else if (sheetNumber=="4") {
    sheetFilterOptions = ["DATE","SHIFT","OPERATOR","AOI_LINE","PRODUCT","CALIBER"];
    sheetFilterName = ["Date","Shift","Operator","AOI_Line","Product","Caliber"];
    async function sheet4() {
      try {
        const totems = await AOI.find(parameters).select("-_id -__v");
        const number = await AOI.find(parameters).exec();
        const model = mongoXlsx.buildDynamicModel(totems);
        mongoXlsx.mongoData2Xlsx(totems, model, function (err, totems) {
          console.log("File saved at:", totems.fullPath);
          fs.move(totems.fullPath, "./downloads/file.xlsx")
            .then(() => {
              console.log("successfully shifted to downloads!");
            })
            .catch((err) => {
              console.error(err);
            });
        });
        res.render("download", {
          number: number,
          sheetNumber:sheetNumber,
          sheetFilterName:sheetFilterName,
          title:"AOI"
        });
      } catch (error) {
        console.log(error);
      }
    }
    const user = sheet4();
  }
});

app.get("/downloadFile", function (req, res) {
  console.log("Downloaded");
  res.download("./downloads/file.xlsx", function (err) {
    if (err) {
      console.log(err);
    }
  });
});

const PORT = process.env.PORT;

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});

