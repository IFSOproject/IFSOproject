

import express from "express";
import bodyParser from "body-parser";
import { dirname } from "path";
import { fileURLToPath } from "url";
import * as fs from 'fs';
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import path from "path";

const __dirname = dirname(fileURLToPath(import.meta.url));

const app = express();
const port = 3000;

app.use(bodyParser.urlencoded({ extended: true }));

app.get("/", (req, res) => {
  res.sendFile(__dirname + "/public/index.html");
});

app.post("/submit", (req, res) => {
  const sname = req.body.rword;
  const dated = req.body.datee;
  const amr = req.body.dated;
  const amer = req.body.us;
  const aliasName = req.body.Alias_namee; 
  const a = req.body.Parentage_Arst;
  const b = req.body.Location_of_arrest;
  const c = req.body.permanentaddr;
  const d = req.body.Convey_Info;
  const e = req.body.Arrest_Info;
  const f = req.body.Case_Name;
  const g = req.body.List_Injury;
  const h = req.body.List_Witn;
  const i = req.body.Ord_Court;
  const j = req.body.ALIAS_FATHER_NAME;
  const k = req.body.COMPANY_NAME;
  const l = req.body.COMPANY_ADDRESS;
  const m = req.body.RENT_OR_YOUR;
  const n = req.body.NAME_OF_OWNER;
  const o = req.body.DATE_JOINED;
  const p = req.body.BANK_INFO;
  const q = req.body.COMPANY_INFO;
  const r = req.body.ASSOCIATE_LIST;
  const s = req.body.LIST_BANKS;
  const t = req.body.COMPANY_OUTSIDE_ACCOUNT;
  const u = req.body.COMPANY_TRANSFER;
  const v = req.body.COMPANY_TOLL_NO;
  const w = req.body.COMPANY_WEBSITE;
  const x = req.body.ALIAS_PAN;
  const y = req.body.COMPANY_SALARY_INFO;
  const z = req.body.COMPANY_SALARY_OUTSIDE_INFO;
  const aa = req.body.COMPANY_ISP_INFO;
  const bb = req.body.COMPANY_CLIENT_INFO;
  const cc = req.body.HOSP_NAME;
  const dd = req.body.CONST_NAME;
  const ee = req.body.AGE;
  const ff = req.body.ALIAS_POSSESION;
  const gg = req.body.OFFICER_1;
  const hh = req.body.OFFICER_2;
  const ii = req.body.VISITING_TO;
  const jj = req.body.DATE_OF_SEIZURE;
  const kk = req.body.TIME_OF_SEIZURE;
  const ll = req.body.LOCATION_SEIZE;
  const mm = req.body.PERSON_SUBMIT_DATA;
  const nn = req.body.PERSON_RECEIVING_DATA;
  const oo = req.body.LIST_OF_SEIZED_ITEMS;


  const fileappl= req.body.language;
  

  const content = fs.readFileSync(
    path.resolve(__dirname, fileappl),
    "binary"
  );

  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  doc.render({
    FIR_NO: sname,
    DATED_FIR: dated,
    DATED:amr,
    US:amer,
    ALIAS_NAME: aliasName,
    PARENTAGE_ARST:a,
    LOCATION_OF_ARREST: b,
    PERMANENT_ADDR:c,
    CONVEY_INFO:d ,   
    ARREST_INFO:e,
    CASE_NAME:f,
    LIST_INJURY:g,
    LIST_WITN:h,
    ORD_COURT:i,
    ALIAS_FATHER_NAME:j,
    COMPANY_NAME:k,
    COMPANY_ADDRESS:l,
    RENT_OR_YOUR:m,
    NAME_OF_OWNER:n,
    DATE_JOINED:o,
    BANK_INFO:p,
    COMPANY_INFO:q,
    ASSOCIATE_LIST:r,
    LIST_BANKS:s,
    COMPANY_OUTSIDE_ACCOUNT:t,
    COMPANY_TRANSFER:u,
    COMPANY_TOLL_NO:v,
    COMPANY_WEBSITE:w,
    ALIAS_PAN:x,
    COMPANY_SALARY_INFO:y,
    COMPANY_SALARY_OUTSIDE_INFO:z,
    COMPANY_ISP_INFO:aa,
    COMPANY_CLIENT_INFO:bb,
    HOSP_NAME:cc,
    CONST_NAME:dd,
    AGE:ee,
    ALIAS_POSSESION:ff,
    OFFICER_1:gg,
    OFFICER_2:hh,
    VISITING_TO:ii,
    DATE_OF_SEIZURE:jj,
    TIME_OF_SEIZURE:kk,
    LOCATION_SEIZE:ll,
    PERSON_SUBMIT_DATA:mm,
    PERSON_RECEIVING_DATA:nn,
    LIST_OF_SEIZED_ITEMS:oo
    
    // Add additional keys as needed based on your template
  });

  const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });

  const outputFileName = "output.docx";
  fs.writeFileSync(path.resolve(__dirname, outputFileName), buf);
  res.download(path.resolve(__dirname, outputFileName));
});

app.listen(port, () => {
  console.log(`Listening on port ${port}`);
});
