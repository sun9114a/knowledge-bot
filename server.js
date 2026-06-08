const express = require("express");
const xlsx = require("xlsx");
const path = require("path");

const app = express();
app.use(express.json());

const EXCEL_PATH = path.join(__dirname, "data", "검사DB.xlsx");

function loadData() {
  const workbook = xlsx.readFile(EXCEL_PATH);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  return xlsx.utils.sheet_to_json(sheet);
}

function findExam(userText) {
  const data = loadData();
  const text = userText.toLowerCase().replace(/\s/g, "");

  return data.find((row) => {
    const name = String(row["검사, 시술, 수술명"] || "").toLowerCase().replace(/\s/g, "");
    const aliases = String(row["별칭"] || "").toLowerCase().replace(/\s/g, "");

    return text.includes(name) || aliases.split(",").some(alias => {
      return alias && text.includes(alias.trim().replace(/\s/g, ""));
    });
  });
}

function makeAnswer(row) {
  if (!row) {
    return "해당 검사 정보를 찾지 못했습니다. 검사명이나 별칭을 다시 입력해주세요.";
  }

  return [
    `📌 ${row["검사, 시술, 수술명"]}`,
    row["금식"] ? `\n🍽 금식: ${row["금식"]}` : "",
    row["IV준비"] ? `\n💉 IV준비: ${row["IV준비"]}` : "",
    row["준비사항"] ? `\n📝 준비사항:\n${row["준비사항"]}` : "",
    row["주의사항"] ? `\n⚠️ 주의사항:\n${row["주의사항"]}` : "",
    row["검사 후 간호"] ? `\n🏥 검사 후 간호:\n${row["검사 후 간호"]}` : "",
    row["위치"] ? `\n📍 위치: ${row["위치"]}` : "",
    row["내선번호"] ? `\n☎️ 내선번호: ${row["내선번호"]}` : "",
  ].join("");
}

app.get("/", (req, res) => {
  res.send("knowledge bot server running");
});

app.post("/skill", (req, res) => {
  const userText =
    req.body?.userRequest?.utterance ||
    req.body?.utterance ||
    "";

  const row = findExam(userText);
  const answer = makeAnswer(row);

  res.json({
    version: "2.0",
    template: {
      outputs: [
        {
          simpleText: {
            text: answer,
          },
        },
      ],
    },
  });
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log("knowledge bot server start");
});