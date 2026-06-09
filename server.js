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
    .replace(/[?？!！.,~]/g, "");
}
let cachedData = [];
function loadData() {
  if (cachedData.length > 0) {
    return cachedData;
  }

  const workbook = xlsx.readFile(EXCEL_PATH);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  cachedData = xlsx.utils.sheet_to_json(sheet);

  return cachedData;
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

  if (text.includes("주의") || text.includes("조심") || text.includes("후간호") || text.includes("검사후") || text.includes("시술후") || text.includes("간호")) {
    return "주의사항";
  }

  if (text.includes("어디") || text.includes("위치") || text.includes("장소") || text.includes("몇층")) {
    return "위치";
  }

  if (text.includes("전화") || text.includes("내선") || text.includes("번호")) {
    return "내선번호";
  }
if (
  text.includes("알려줘") ||
  text.includes("정보") ||
  text.includes("설명") ||
  text.includes("뭐야") ||
  text.includes("대해")
) {
  return "전체";
}
  return "전체";
}

function levenshtein(a, b) {
  const dp = Array.from({ length: a.length + 1 }, () =>
    Array(b.length + 1).fill(0)
  );

  for (let i = 0; i <= a.length; i++) dp[i][0] = i;
  for (let j = 0; j <= b.length; j++) dp[0][j] = j;

  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,
        dp[i][j - 1] + 1,
        dp[i - 1][j - 1] + cost
      );
    }
  }

  return dp[a.length][b.length];
}

function similarity(a, b) {
  if (!a || !b) return 0;

  const longer = a.length >= b.length ? a : b;
  const shorter = a.length >= b.length ? b : a;
  const distance = levenshtein(longer, shorter);

  return (longer.length - distance) / longer.length;
}

function removeIntentWords(text) {
  return normalize(text)
    .replace(/금식|밥|식사|npo/g, "")
    .replace(/iv|아이브이|라인|주사|혈관|정맥|루트|route|확보/g, "")
    .replace(/준비사항|준비물|준비|전처치/g, "")
    .replace(/주의사항|주의|조심|후간호|검사후|시술후|간호/g, "")
    .replace(/어디서|어디|위치|장소|몇층|해|하나요|하니|해야돼|해야되|돼|되/g, "")
    .replace(/전화|내선|번호/g, "")
.replace(/알려줘|정보|설명|뭐야|에대해|대해/g, "");
}

function findExam(userText) {
  const data = loadData();

  const fullText = normalize(userText);
  const searchText = removeIntentWords(userText);

  let bestMatch = null;
  let bestScore = 0;

  for (const row of data) {
    const name = normalize(row["검사명"]);
    const aliases = String(row["별칭"] || "")
      .split(/[,/|]/)
      .map((a) => normalize(a))
      .filter(Boolean);

    const candidates = [name, ...aliases].filter(Boolean);

    for (const candidate of candidates) {
      let score = 0;

      if (searchText === candidate || fullText === candidate) {
        score = 100;
      } else if (fullText.includes(candidate)) {
        score = 95 + candidate.length;
      } else if (searchText.includes(candidate)) {
        score = 90 + candidate.length;
      } else if (candidate.includes(searchText) && searchText.length >= 2) {
        score = 70 + searchText.length;
      } else {
        const sim = similarity(searchText, candidate);

        if (candidate.length <= 3 && sim >= 0.9) {
          score = 60 + sim * 10;
        } else if (candidate.length >= 4 && sim >= 0.72) {
          score = 50 + sim * 10;
        }
      }

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