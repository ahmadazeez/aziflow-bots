# AZIflow Bots 🤖

بوتات بتسهّل شغل AZIflow للتخليص الجمركي: تعبئة الفواتير والمستندات وتصديرها PDF و Excel.

## التشغيل

افتح `index.html` بـ **Chrome أو Edge** (دبل كليك بيكفي، ما بدها سيرفر). بتطلعلك الصفحة الرئيسية وفيها كل البوتات.

> المجلد الرئيسي اللي بتختاره للحفظ بينحفظ ومشترك بين كل البوتات. كل ملف بيتحفظ بمجلد فرعي باسم تاريخه.

## الهيكل

```
index.html                    ← الصفحة الرئيسية (قائمة البوتات)
shared/
  azi.css                     ← الثيم المشترك (ألوان، حقول، أزرار، شاشة الكريدت)
  azi.js                      ← أدوات مشتركة: مجلد الحفظ، Excel، PDF، رسائل، تواريخ
  brand/header.jpg(.js)       ← ترويسة AZIflow للفواتير
bots/
  registry.js                 ← قائمة البوتات اللي بتظهر بالرئيسية
  transit-invoice/            ← بوت فاتورة ترانزيت
    index.html                ← الواجهة + شكل الفاتورة
    bot.js                    ← المنطق
    template.xlsx(.js)        ← قالب Excel
tools/
  embed.js                    ← يحوّل صورة/قالب إلى .js
```

## إضافة بوت جديد

1. انسخ مجلد `bots/transit-invoice` باسم جديد، مثلاً `bots/export-invoice`.
2. عدّل `index.html` و `bot.js` حسب البوت.
3. إذا عندك قالب Excel جديد، حطّه بالمجلد وحوّله:
   ```
   node tools/embed.js bots/export-invoice/template.xlsx exportInvoiceXlsx
   ```
   وبالصفحة: `<script src="template.xlsx.js"></script>` وبالكود: `AZI.asset("exportInvoiceXlsx")`.
4. أضف سطر بـ `bots/registry.js`، وبيظهر بالصفحة الرئيسية.

### ليش القوالب والصور كـ `.js`؟
عشان البوت يشتغل بدبل كليك (`file://`). المتصفح بيمنع قراءة الملفات المحلية بـ `fetch`، وبيمنع رسم صورة محلية على canvas للـ PDF. لو عدّلت `template.xlsx` أو `header.jpg`، **لازم** تعيد تشغيل `tools/embed.js` عليه.

## أدوات `AZI` المشتركة

| الأداة | شو بتعمل |
|---|---|
| `AZI.folder.mount(el)` | صندوق اختيار مجلد الحفظ |
| `AZI.folder.ready()` | تتأكد من إذن المجلد (استدعِها أول الكبسة) |
| `AZI.folder.save(blob, name, sub)` | تحفظ بالمجلد/`sub`، أو تنزّل الملف إذا ما في مجلد |
| `AZI.xlsx.fill(b64, sheet => ...)` | تعبّي قالب Excel (`setStr` / `setNum` / `clear`) |
| `AZI.pdf.fromHTML(html)` | ترسم HTML إلى PDF بحجم A4 |
| `AZI.messages(errEl, okEl)` | رسائل الخطأ والنجاح |
| `AZI.credit.mount(el)` | زر «اكبس هون ضروري» ⚡ |
| `AZI.today / fmtDate / excelSerial / money / htmlEsc` | أدوات صغيرة |
