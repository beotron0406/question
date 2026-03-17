/**
 * Convert "BỘ CÂU HỎI ÔN THI.docx" → questions.json + questions_data.js
 *
 * Cấu trúc trong file DOCX:
 *   - Mỗi câu hỏi, đáp án (A/B/C/D) là một đoạn (paragraph) riêng.
 *   - Với câu 1–~340: đáp án đúng được tô màu xanh lá (#28A745).
 *   - Với câu ~341+: có dòng "Đáp án: X" hoặc "Đáp án X" (chữ rõ ràng).
 */

const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");

const DOCX_PATH = path.join(__dirname, "BỘ CÂU HỎI ÔN THI.docx");
const GREEN_COLOR = "28A745"; // màu xanh đánh dấu đáp án đúng

// ─── Bước 1: Đọc XML từ DOCX ────────────────────────────────────────────────

function readDocxXml(filePath) {
  const zip = new AdmZip(filePath);
  const entry = zip.getEntry("word/document.xml");
  if (!entry) throw new Error("Không tìm thấy word/document.xml trong file DOCX");
  return entry.getData().toString("utf8");
}

// ─── Bước 2: Trích xuất danh sách paragraph ─────────────────────────────────

function extractParagraphs(xml) {
  const paragraphs = [];
  let i = 0;

  while (i < xml.length) {
    // Tìm thẻ mở <w:p> (chỉ là paragraph, không phải <w:pPr>, <w:pStyle>, ...)
    const pStart = xml.indexOf("<w:p", i);
    if (pStart === -1) break;

    // Kiểm tra ký tự thứ 4 sau "<w:p" phải là ' ' hoặc '>'
    const ch = xml[pStart + 4];
    if (ch !== " " && ch !== ">") {
      i = pStart + 4;
      continue;
    }

    // Tìm thẻ đóng </w:p>
    const pEnd = xml.indexOf("</w:p>", pStart);
    if (pEnd === -1) break;

    const paraXml = xml.substring(pStart, pEnd + 6);

    // Kiểm tra có màu xanh đánh dấu đáp án đúng không
    const isGreen =
      paraXml.includes(GREEN_COLOR) ||
      paraXml.includes(GREEN_COLOR.toLowerCase());

    // Ghép tất cả <w:t> thành text
    let text = "";
    const tRe = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
    let m;
    while ((m = tRe.exec(paraXml)) !== null) {
      text += m[1];
    }

    text = decodeXmlEntities(text).trim();
    if (text) {
      paragraphs.push({ text, isGreen });
    }

    i = pEnd + 6;
  }

  return paragraphs;
}

function decodeXmlEntities(str) {
  return str
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(parseInt(n, 10)));
}

// ─── Bước 3: Parse câu hỏi từ danh sách paragraph ───────────────────────────

function parseQuestions(paragraphs) {
  const questions = [];
  let cur = null; // câu hỏi đang xử lý
  let lastOptIdx = -1; // index option vừa đọc (0=A,1=B,2=C,3=D)

  const pushCurrent = () => {
    if (!cur) return;
    if (cur.options.filter((o) => o.text.length > 0).length < 4) return;
    questions.push(finalizeQuestion(cur));
    cur = null;
    lastOptIdx = -1;
  };

  for (const { text, isGreen } of paragraphs) {
    // ── Câu mới: "Câu N:" ────────────────────────────────────────────────────
    const qMatch = text.match(/^Câu\s+(\d+)\s*[:\.]?\s*([\s\S]*)/);
    if (qMatch && /^\d+$/.test(qMatch[1])) {
      pushCurrent();
      cur = {
        id: parseInt(qMatch[1], 10),
        question: cleanText(qMatch[2]),
        options: [
          { text: "", isGreen: false },
          { text: "", isGreen: false },
          { text: "", isGreen: false },
          { text: "", isGreen: false },
        ],
        explicitAnswer: -1,
      };
      lastOptIdx = -1;
      continue;
    }

    if (!cur) continue;

    // ── Option A/B/C/D ────────────────────────────────────────────────────────
    const optMatch = text.match(/^([A-D])\s*:\s*([\s\S]*)/);
    if (optMatch) {
      const idx = "ABCD".indexOf(optMatch[1]);
      let optText = optMatch[2];

      // Trong text option có thể nhúng "Đáp án: X" (thường ở option D)
      const embeddedAns = extractEmbeddedAnswer(optText);
      if (embeddedAns !== null) {
        cur.explicitAnswer = embeddedAns.answerIdx;
        optText = embeddedAns.cleanText;
      }

      cur.options[idx] = { text: cleanText(optText), isGreen };
      lastOptIdx = idx;
      continue;
    }

    // ── Dòng "Đáp án: X" standalone ─────────────────────────────────────────
    const ansMatch = text.match(/^Đáp\s*án\s*[:\s]\s*([A-D])\b/i);
    if (ansMatch) {
      cur.explicitAnswer = "ABCD".indexOf(ansMatch[1].toUpperCase());
      lastOptIdx = -1;
      continue;
    }

    // ── Continuation: nối vào option cuối hoặc question ──────────────────────
    // Bỏ qua các dòng giải thích / tác giả
    if (isNoiseLine(text)) continue;

    if (lastOptIdx >= 0) {
      // Có thể có "Đáp án" nhúng trong dòng continuation
      const emb = extractEmbeddedAnswer(text);
      if (emb !== null) {
        cur.explicitAnswer = emb.answerIdx;
        const extra = emb.cleanText;
        if (extra) {
          cur.options[lastOptIdx].text =
            cleanText(cur.options[lastOptIdx].text + " " + extra);
        }
      } else {
        cur.options[lastOptIdx].text = cleanText(
          cur.options[lastOptIdx].text + " " + text
        );
      }
    } else if (cur.options[0].text === "") {
      // Chưa có option nào → continuation của question
      cur.question = cleanText(cur.question + " " + text);
    }
  }

  pushCurrent();
  return questions;
}

/** Tách "Đáp án: X" nhúng trong text option, trả về null nếu không có */
function extractEmbeddedAnswer(text) {
  const re = /Đáp\s*án\s*[:\s]\s*([A-D])\b/i;
  const m = text.match(re);
  if (!m) return null;
  const answerIdx = "ABCD".indexOf(m[1].toUpperCase());
  const cutPos = text.search(re);
  const cleanText = text.substring(0, cutPos).trim();
  return { answerIdx, cleanText };
}

/** Xác định đây có phải dòng "rác" (giải thích, tác giả, ...) không */
function isNoiseLine(text) {
  const noisePatterns = [
    /^Lý giải\s*:/i,
    /^Giải thích\s*:/i,
    /^Phạm Văn Bình/i,
    /^Theo dõi\s/i,
    /^Ủng hộ tác giả/i,
    /^Căn cứ (vào|theo)/i,
    /^Phân tích/i,
    /^Do đó/i,
    /^Vì (vậy|thế)/i,
    /^Kết luận/i,
  ];
  return noisePatterns.some((re) => re.test(text));
}

function finalizeQuestion(cur) {
  let answer = cur.explicitAnswer;

  // Nếu chưa có answer rõ ràng → tìm option màu xanh
  if (answer === -1) {
    for (let i = 0; i < 4; i++) {
      if (cur.options[i].isGreen) {
        answer = i;
        break;
      }
    }
  }

  return {
    id: cur.id,
    question: cur.question,
    options: cur.options.map((o) => o.text),
    answer,
  };
}

function cleanText(str) {
  return str.replace(/\s+/g, " ").trim();
}

// ─── Main ────────────────────────────────────────────────────────────────────

function main() {
  console.log("📖 Đọc file DOCX...");
  const xml = readDocxXml(DOCX_PATH);

  console.log("🔍 Trích xuất paragraphs...");
  const paragraphs = extractParagraphs(xml);
  console.log(`   Tổng số đoạn văn: ${paragraphs.length}`);

  console.log("📝 Parse câu hỏi...");
  const questions = parseQuestions(paragraphs);
  console.log(`   Tổng câu đã parse: ${questions.length}`);

  // Lọc câu hợp lệ (đủ 4 đáp án + có answer)
  const valid = questions.filter(
    (q) =>
      q.options.every((o) => o && o.length > 0) &&
      q.answer >= 0 &&
      q.answer <= 3
  );
  console.log(`   Câu hợp lệ: ${valid.length}`);

  // Thống kê câu bị lỗi
  const invalid = questions.filter(
    (q) => !q.options.every((o) => o && o.length > 0) || q.answer < 0
  );
  if (invalid.length > 0) {
    console.warn(`⚠️  ${invalid.length} câu bị lỗi (thiếu đáp án / option):`);
    invalid.slice(0, 10).forEach((q) => {
      console.warn(
        `   Câu ${q.id}: answer=${q.answer}, options=${JSON.stringify(
          q.options
        )}`
      );
    });
  }

  // Ghi questions.json (dạng đẹp)
  fs.writeFileSync(
    path.join(__dirname, "questions.json"),
    JSON.stringify(valid, null, 2),
    "utf8"
  );
  console.log("✅ Đã tạo questions.json");

  // Ghi questions_data.js (để nhúng trực tiếp vào HTML)
  const jsContent = `// Auto-generated by convert.js - ${new Date().toLocaleString("vi-VN")}
const QUESTIONS_DATA = ${JSON.stringify(valid)};
`;
  fs.writeFileSync(
    path.join(__dirname, "questions_data.js"),
    jsContent,
    "utf8"
  );
  console.log("✅ Đã tạo questions_data.js");

  // Tổng kết
  const sets = Math.ceil(valid.length / 25);
  console.log(`\n🎉 Hoàn thành! ${valid.length} câu hỏi → ${sets} bộ đề (25 câu/bộ)`);
}

main();
