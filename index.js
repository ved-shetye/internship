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

let mainProducts=["APFC","Danlaw","Dellorto","Dura","Emerald","IJL230B","IJL660A","Kongsberg","LEDs","Molbio","Neutral_4P","S3RP","Tyco"];
let otherProducts=["APFC","Danlaw","Dellorto","Dura","IJL230B","IJL660A","Kongsberg","Molbio","S3RP","Tyco"];
let mainActivities = ["Antenna assembly","Base Assy","Battery Assembly","Battery cover glueing","Bottom cover Gasket","Bottom cover glueing","Box build assembly","Cable Insertion","Cleaning","Conn Assembly","Conn soldering","CS1","CS2","DIP swtich sold","Display Sold","EOL Testing","Final VI","Heat stacking 1","Heat stacking 2","Heat stacking 3","IC Soldering","ICT 1","ICT 2","Lead Cutting","Metal Plate Assembly","NTC Forming","NTC Soldering","Packing","PCBA testing","Potting VI","Robotic soldering","Screwing","Singulation 2","Switch Sold","Testing","Testing 1","Testing 2","TH Soldering","TH1","TH2","Top cover Gasket","Touch up","VI"];
let otherActivities = ["Cleaning","Routing","Singulation","Singulation 1"];
let smtLines = ["Line 1","Line 2","Line 3","Line 4","Line 5"];
let smtProducts = ["APFC","Danlaw","Dellorto","Dura","Emerald","EMERALD","IJL660A","KA","LEDs","Neutral_4P","S3RP","Tyco"];
let otherTypes = ["AP30_11","AP30_13","AP30_25","AQ30_11","AW_30","Bajaj_ETB","BMW","BW_30","BMW K03","BMW K2X DX","BMW K2X SX","Controller series-2","Controller_144","Controller_96","D1_100","D1_32","D1_80","D3_100","D3_50","D3_62.5","D3_80","D4 Main Board","DAQ Series-2","Daughter Board","Diesel","Gasoline","K08 BMW","Illumination","Lamp_LED","LH_230B","LH_660A","Logger_4G","NM20_74","NM20_76","NP20_76","NP30_25","NP30_74","NP30_76","NQ30_76","P_202","P_33","Power Supply series-2","PS_144","PS_96","Relay_12step","Relay_14step","Relay_16step","Relay_4step","Relay_6step","Relay_8step","Relay_Card 96","Renault","Renault-melexis","RH_230B","RH660A","S106 Main","S201 bezel","Sensor","Sensor_2_in_1","Sensor_3_in_1","SG_1","Truenat","Testing","TVS D33","VI","Volvo"];
let mainTypes = ["AP30_11","AP30_13","AP30_25","AQ30_11","AW_30","Bajaj_ETB","BMW","BW_30","BMW K03","BMW K2X DX","BMW K2X SX","Controller series-2","Controller_144","Controller_96","D1_100","D1_32","D1_80","D3_100","D3_50","D3_62.5","D3_80","D4 Main Board","DAQ Series-2","Daughter Board","Diesel","Gasoline","K08 BMW","Illumination","Lamp_LED","LH_230B","LH_660A","Logger_4G","NM20_74","NM20_76","NP20_76","NP30_25","NP30_74","NP30_76","NQ30_76","P_202","P_33","Power Supply series-2","PS_144","PS_96","Relay_12step","Relay_14step","Relay_16step","Relay_4step","Relay_6step","Relay_8step","Relay_Card 96","Renault","Renault-melexis","RH_230B","RH660A","S106 Main","S201 bezel","Sensor","Sensor_2_in_1","Sensor_3_in_1","SG_1","Truenat","Testing","TVS D33","Thyristor","VI","Volvo"];
let smtCaliber = ["1.1 bot","1.1 top","230B LH Bot","230B LH Top","230B RH Bot","230B RH Top","660A LH Top","660A LH Bot","660A RH Top","660A RH Bot","AP30","AP30_11","AP30_13","AP30_25","AP30_33","Ap30_31","AQ30_11","AW_30","Bajaj_ETB","BMW","BW_30","Controller 144x144 Bot","Controller 144x144 Top","Controller 96x96 Bot","Controller 96x96 Top","Controller Series-2 Bot","Controller Series-2 Top","D1 100% Top","D1 32% Top","D1 62.5% Top","D1 Bot","D1_Bot","D3 100% Bot","D3 100% Top","D3 50% Bot","D3 50% Top","D3 62.5% Bot","D3 62.55 Top","D3 80% Bot","D3 80% Top","DAQ Series-2","D5 Daughter Board","Diesel","Gasoline","Illumination Board Bot","Illumination Board Top","Lamp_LED_Top","Lamp_LED_Bot","Logger4G_bot","Logger4G_top","LTCC_placement","Neutral_4P","NM20","NM20_76","NP20_76","NP30_76","P_202","P_33","PCB assy","Power Supply series-2","Power Supply series-2 Bot","Power Supply series-2 Top","Power Supply 96x96 Bot","Power Supply 96x96 Top","PS_144","Renault","Renault-melexis","RTD_placement","S106 Main","s201 bezel_Bot","s201 bezel_top","Sensor","Sensor Board","Sensor_3_in_1","SG_1","Thyristor","Volvo"];
let AoiAOI_line = ["AOI 1","AOI 2","Yestech"];
let AoiProducts = ["APFC","Danlaw","Dellorto","Dura","Emerald","IJL230B","IJL660A","Kongsberg","LEDs","Molbio","Neutral_4P","S3RP","Tyco"];
let AoiCaliber = ["1.1 bot","1.1 top","230B LH Bot","230B LH Top","230B RH Bot","230B RH Top","660A LH Top","660A LH Bot","660A RH Top","660A RH Bot","AP30","AP30_11","AP30_13","AP30_25","AP30_33","Ap30_31","AQ30_11","AW_30","Bajaj_ETB","BMW","BW_30","Controller 144x144 Bot","Controller 144x144 Top","Controller 96x96 Bot","Controller 96x96 Top","Controller Series-2 Bot","Controller Series-2 Top","D1 100% Top","D1 32% Top","D1 62.5% Top","D1 Bot","D1_Bot","D3 100% Bot","D3 100% Top","D3 50% Bot","D3 50% Top","D3 62.5% Bot","D3 62.55 Top","D3 80% Bot","D3 80% Top","DAQ Series-2","D5 Daughter Board","Diesel","Gasoline","Illumination Board Bot","Illumination Board Top","Lamp_LED_Top","Lamp_LED_Bot","Logger4G_bot","Logger4G_top","LTCC_placement","Neutral_4P","NM20","NM20_76","NP20_76","NP30_76","P_202","P_33","PCB assy","Power Supply series-2","Power Supply series-2 Bot","Power Supply series-2 Top","Power Supply 96x96 Bot","Power Supply 96x96 Top","PS_144","Renault","Renault-melexis","RTD_placement","S106 Main","s201 bezel_Bot","s201 bezel_top","Sensor","Sensor Board","Sensor_3_in_1","SG_1","Thyristor","Volvo"];

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
    mainTypes:mainTypes,
    mainActivities:mainActivities,
    otherProducts:otherProducts,
    otherActivities:otherActivities,
    otherTypes:otherTypes,
    smtLines:smtLines,
    smtProducts:smtProducts,
    smtCaliber:smtCaliber,
    AoiAOI_line:AoiAOI_line,
    AoiProducts:AoiProducts,
    AoiCaliber:AoiCaliber
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

