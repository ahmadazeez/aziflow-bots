// ===== بوت فاتورة ترانزيت =====
// يعبّي قالب Excel (template.xlsx) ويرسم نفس الفاتورة كـ PDF.

const FIXED_YEAR = "2024";
const MAX_DRIVERS = 10;
const DRIVER_ROWS_XLSX = [20, 21, 22, 23, 24, 25, 26, 27, 28, 29]; // صفوف السائقين بالقالب (أعمدة G و H)
const ITEMS = ["رسوم جمركية", "اتعاب التخليص", "لوحات", "نهايه حبل", "حبل", "بدل عمال", "رسوم موعد+سداد", "سداد بطحاء", "سداد حديثه"];

const $ = id => document.getElementById(id);
const msg = AZI.messages($("errorBox"), $("successBox"));
let drivers = [{ carNo: "", name: "" }];

AZI.folder.mount($("folder"));
AZI.credit.mount($("credit"));
$("date").value = AZI.today();

// ---------- السائقون ----------
function renderDrivers() {
  const list = $("driversList"); list.innerHTML = "";
  drivers.forEach((d, i) => {
    const row = document.createElement("div"); row.className = "driverRow";
    row.innerHTML = '<div class="driverNum">' + (i + 1) + "</div>" +
      '<input type="text" placeholder="رقم السيارة" value="' + AZI.htmlEsc(d.carNo) + '" data-i="' + i + '" data-f="carNo">' +
      '<input type="text" placeholder="اسم السائق" value="' + AZI.htmlEsc(d.name) + '" data-i="' + i + '" data-f="name">' +
      '<button class="removeBtn" type="button" data-rem="' + i + '"' + (drivers.length === 1 ? " disabled" : "") + ">✕</button>";
    list.appendChild(row);
  });
  list.querySelectorAll("input").forEach(inp => inp.addEventListener("input", e => { drivers[+e.target.dataset.i][e.target.dataset.f] = e.target.value; }));
  list.querySelectorAll("[data-rem]").forEach(btn => btn.addEventListener("click", e => {
    const i = +e.target.dataset.rem; if (drivers.length > 1) { drivers.splice(i, 1); renderDrivers(); }
  }));
  $("addDriverBtn").disabled = drivers.length >= MAX_DRIVERS;
}
renderDrivers();

$("addDriverBtn").addEventListener("click", () => {
  if (drivers.length < MAX_DRIVERS) { drivers.push({ carNo: "", name: "" }); renderDrivers(); }
});

// ---------- المبلغ ----------
$("amount").addEventListener("input", e => {
  const n = parseFloat(e.target.value) || 0, box = $("calcBox");
  if (n > 0) {
    box.classList.remove("hidden");
    const v = n.toFixed(2) + " " + AZI.CURRENCY;
    $("c1").textContent = v; $("c2").textContent = v; $("c3").textContent = v;
  } else box.classList.add("hidden");
});

// ---------- البيانات ----------
function getData() {
  return {
    docNo: $("docNo").value.trim(),
    date: $("date").value,
    amount: parseFloat($("amount").value) || 0,
    drivers: drivers.filter(d => d.carNo || d.name),
  };
}
function validate(d) {
  if (!d.docNo) { msg.error("الرجاء إدخال رقم البيان"); return false; }
  if (!d.date) { msg.error("الرجاء إدخال التاريخ"); return false; }
  if (d.amount <= 0) { msg.error("الرجاء إدخال مبلغ صحيح أكبر من صفر"); return false; }
  return true;
}
const fileName = (d, ext) => "فاتورة_ترانزيت_" + d.docNo + "_" + d.date + "." + ext;

// يغلّف أي زر تصدير: تحقق، تعطيل الزر، حفظ، رسالة
function exportButton(btn, label, kind, build, opts = {}) {
  btn.addEventListener("click", async () => {
    msg.clear();
    const d = getData(); if (!validate(d)) return;
    const hasFolder = await AZI.folder.ready();
    if (opts.requireFolder && AZI.folder.supported && !hasFolder) { msg.error("الرجاء اختيار مجلد الحفظ أولاً (زر اختيار المجلد)"); return; }
    btn.disabled = true; btn.textContent = "⏳ جاري...";
    try {
      const { blob, ext } = await build(d);
      const where = await AZI.folder.save(blob, fileName(d, ext), d.date);
      msg.success(where === "folder" ? "تم حفظ " + kind + " في مجلد " + d.date : "تم تنزيل " + kind);
    } catch (e) {
      console.error(e); msg.error("خطأ في إنشاء " + kind + ": " + e.message);
    } finally {
      btn.disabled = false; btn.textContent = label;
    }
  });
}

// ---------- Excel ----------
exportButton($("xlsBtn"), "📊 تحميل Excel", "Excel", async d => {
  const { setStr, setNum, clear } = AZI.xlsx;
  const blob = await AZI.xlsx.fill(AZI.asset("transitInvoiceXlsx"), sheet => {
    sheet = setStr(sheet, "B16", d.docNo);
    sheet = setNum(sheet, "D16", AZI.excelSerial(d.date));
    sheet = setNum(sheet, "C21", d.amount);
    sheet = setNum(sheet, "B30", d.amount);
    sheet = setNum(sheet, "B33", d.amount);
    DRIVER_ROWS_XLSX.forEach((r, idx) => {
      const drv = d.drivers[idx];
      if (drv) { sheet = setStr(sheet, "G" + r, drv.carNo || ""); sheet = setStr(sheet, "H" + r, drv.name || ""); }
      else { sheet = clear(sheet, "G" + r); sheet = clear(sheet, "H" + r); }
    });
    return sheet;
  });
  return { blob, ext: "xlsx" };
});

// ---------- PDF ----------
function buildInvoiceHTML(d) {
  const esc = AZI.htmlEsc, date = AZI.fmtDate(d.date);
  const amt = d.amount > 0 ? AZI.money(d.amount) : "";
  const n = Math.max(ITEMS.length, d.drivers.length);
  let rows = "";
  for (let i = 0; i < n; i++) {
    const drv = d.drivers[i] || { carNo: "", name: "" };
    rows += "<tr>" +
      (i === 0 ? '<td rowspan="' + n + '" class="ctr"></td>' : "") +
      '<td class="amt">' + (i === 1 ? amt : "") + "</td>" +
      '<td class="rgt">' + esc(ITEMS[i] || "") + "</td>" +
      '<td class="ctr">' + esc(drv.carNo || "") + "</td>" +
      '<td class="ctr">' + esc(drv.name || "") + "</td>" +
    "</tr>";
  }
  return '<div class="inv">' +
    '<div class="inv-hdr"><img src="data:image/jpeg;base64,' + AZI.asset("headerJpg") + '"></div>' +
    '<div class="inv-contact"><span>المهندس موسى أبومسامح: 00962780994109</span><span>المهندس عيسى أبومسامح: 00962796165695</span></div>' +
    '<div class="inv-body">' +
      '<div class="inv-titlebar"><span>فاتورة ضريبة رقم :</span><span class="v">' + esc(d.docNo) + "</span><span>" + FIXED_YEAR + '</span><span>الموافق :</span><span class="v">' + date + "</span></div>" +
      '<table class="t"><tbody>' +
        '<tr><td class="lblc" style="width:18%">جهه الوصول</td><td class="lblc" style="width:24%">نوع الفاتورة</td><td style="width:24%"></td><td class="lblc rgt" style="width:34%">المدينة / البلد<br><span class="en">City / Country</span></td></tr>' +
        '<tr><td class="ctr">الأردن</td><td class="rgt">نوع البضاعه//</td><td></td><td></td></tr>' +
      "</tbody></table>" +
      '<div class="gap"></div>' +
      '<table class="t"><tbody>' +
        '<tr><td class="lblc">رقم التواصل<br><span class="en">CONTACT NO</span></td><td class="lblc">نوع البيان<br><span class="en">TYPE</span></td><td class="lblc">تاريخ البيان<br><span class="en">DATE</span></td><td class="lblc">بيان رقم<br><span class="en">DOCUMENT NO</span></td></tr>' +
        '<tr><td class="ctr"></td><td class="ctr">ترانزيت</td><td class="ctr">' + date + '</td><td class="ctr">' + esc(d.docNo) + "</td></tr>" +
      "</tbody></table>" +
      '<div class="gap"></div>' +
      '<table class="t"><tbody>' +
        '<tr><td class="lblc" style="width:9%">الضريبة<br><span class="en">VAT</span></td><td class="lblc" style="width:15%">المبلغ<br><span class="en">Amount</span></td><td class="lblc" style="width:30%">البيان</td><td class="lblc" style="width:23%">رقم السيارة</td><td class="lblc" style="width:23%">اسم السائق</td></tr>' +
        rows +
      "</tbody></table>" +
      '<div class="gap"></div>' +
      '<table class="t"><tbody>' +
        '<tr><td class="amt" style="width:18%">' + amt + '</td><td class="rgt" style="width:46%">إجمالي مبلغ الفاتورة <span class="en">Total Amount Without VAT</span></td><td style="width:36%"></td></tr>' +
        '<tr><td class="amt"></td><td class="rgt">إجمالي المبلغ الخاضع للضريبة <span class="en">Total Taxable Amount</span></td><td></td></tr>' +
        '<tr><td class="amt"></td><td class="rgt">ضريبة القيمة المضافة 15% <span class="en">VAT Amount 15%</span></td><td></td></tr>' +
        '<tr><td class="amt" style="background:#e8f5e9;font-weight:700">' + amt + '</td><td class="rgt" style="background:#e8f5e9;font-weight:700">إجمالي المبلغ شامل الضريبة <span class="en">Total Amount With VAT</span></td><td style="background:#e8f5e9"></td></tr>' +
      "</tbody></table>" +
    "</div>" +
  "</div>";
}

exportButton($("pdfBtn"), "📄 حفظ PDF", "PDF", async d => ({ blob: await AZI.pdf.fromHTML(buildInvoiceHTML(d)), ext: "pdf" }), { requireFolder: true });

// ---------- تفريغ ----------
$("resetBtn").addEventListener("click", () => {
  $("docNo").value = ""; $("date").value = AZI.today(); $("amount").value = "";
  drivers = [{ carNo: "", name: "" }]; renderDrivers();
  $("calcBox").classList.add("hidden");
  msg.clear();
});
