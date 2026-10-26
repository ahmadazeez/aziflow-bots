// يحوّل ملف (صورة، قالب Excel...) إلى ملف .js فيه الملف بصيغة base64
// ليشتغل البوت حتى لو انفتح بدبل كليك (file://) بدون سيرفر.
//
// الاستخدام:
//   node tools/embed.js <الملف> <اسم المفتاح>
// مثال:
//   node tools/embed.js bots/transit-invoice/template.xlsx transitInvoiceXlsx
//   → ينشئ bots/transit-invoice/template.xlsx.js
//   → وبالصفحة: AZI_ASSETS.transitInvoiceXlsx

const fs = require("fs");
const path = require("path");

const [, , file, key] = process.argv;
if (!file || !key) {
  console.error("الاستخدام: node tools/embed.js <الملف> <اسم المفتاح>");
  process.exit(1);
}

const b64 = fs.readFileSync(file).toString("base64");
const out = file + ".js";
fs.writeFileSync(
  out,
  "// ملف مولّد تلقائياً من " + path.basename(file) + " — لا تعدّله يدوياً، عدّل الأصل وشغّل tools/embed.js\n" +
  "window.AZI_ASSETS = window.AZI_ASSETS || {};\n" +
  "AZI_ASSETS[" + JSON.stringify(key) + "] = \"" + b64 + "\";\n"
);
console.log("✓ " + out + " (" + Math.round(b64.length / 1024) + " KB)");
