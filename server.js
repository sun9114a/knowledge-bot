const express = require("express");
const xlsx = require("xlsx");
const path = require("path");

const app = express();
app.use(express.json());

const EXCEL_PATH = path.join(__dirname, "data", "검사DB.xlsx");

function normalize(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/\s/g, "")
    .replace(/[?？!！.,]/g, "");
}

function loadData() {
  const workbook = xlsx.readFile(EXCEL_PATH);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  return xlsx.utils.sheet_to_json(sheet);
}

function detectIntent(userText) {
  const text = normalize(userText);

  if (text.includes("금식") || text.includes("밥") || text.includes("식사") || text.includes("npo")) {
    return "금식";
  }

if (
  text.includes("iv") ||
  text.includes("아이브이") ||
  text.includes("라인") ||
  text.includes("주사") ||
  text.includes("혈관") ||
  text.includes("정맥") ||
  text.includes("루트") ||
  text.includes("route") ||
  text.includes("확보")
) {
  return "IV준비";
}

  if (text.includes("준비") || text.includes("준비물") || text.includes("전처치")) {
    return "준비사항";
  }

  if (text.includes("주의") || text.includes("조심")) {
    return "주의사항";
  }

  if (text.includes("간호") || text.includes("후간호") || text.includes("검사후") || text.includes("시술후")) {
    return "검사 후 간호";
  }

  if (text.includes("어디") || text.includes("위치") || text.includes("장소") || text.includes("몇층")) {
    return "위치";
  }

  if (text.includes("전화") || text.includes("내선") || text.includes("번호")) {
    return "내선번호";
  }

  return "전체";
}

function findExam(userText) {
  const data = loadData();
  const text = normalize(userText);

  let bestMatch = null;
  let bestScore = 0;

  for (const row of data) {
    const name = normalize(row["검사명"]);
    const aliases = String(row["별칭"] || "")
      .split(/[,/|]/)
      .map(a => normalize(a))
      .filter(Boolean);

    const candidates = [name, ...aliases].filter(Boolean);

    for (const candidate of candidates) {
      let score = 0;

      if (text === candidate) score = 100;
      else if (text.includes(candidate)) score = candidate.length;
      else if (candidate.includes(text)) score = text.length - 1;

      if (score > bestScore) {
        bestScore = score;
        bestMatch = row;
      }
    }
  }

  return bestScore > 0 ? bestMatch : null;
}

function value(row, key) {
  return row[key] ? String(row[key]).trim() : "";
}

function makeAnswer(row, intent) {
  if (!row) {
    return "해당 검사 정보를 찾지 못했습니다. 검사명이나 별칭을 다시 입력해주세요.";
  }

  const examName = value(row, "검사명");

  if (intent !== "전체") {
    const result = value(row, intent);

    if (!result) {
      return `📌 ${examName}\n\n${intent} 정보가 등록되어 있지 않습니다.`;
    }

    return `📌 ${examName}\n${intent}: ${result}`;
  }

  return [
    `📌 ${examName}`,
    value(row, "금식") ? `\n🍽 금식: ${value(row, "금식")}` : "",
    value(row, "IV준비") ? `\n💉 IV준비: ${value(row, "IV준비")}` : "",
    value(row, "준비사항") ? `\n📝 준비사항:\n${value(row, "준비사항")}` : "",
    value(row, "주의사항") ? `\n⚠️ 주의사항:\n${value(row, "주의사항")}` : "",
    value(row, "검사 후 간호") ? `\n🏥 검사 후 간호:\n${value(row, "검사 후 간호")}` : "",
    value(row, "위치") ? `\n📍 위치: ${value(row, "위치")}` : "",
    value(row, "내선번호") ? `\n☎️ 내선번호: ${value(row, "내선번호")}` : "",
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
  const intent = detectIntent(userText);
  const answer = makeAnswer(row, intent);

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