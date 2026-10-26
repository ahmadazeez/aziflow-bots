// ===== AZIflow Bots — أدوات مشتركة لكل البوتات =====
// كل بوت يحمّل هذا الملف ويستخدم AZI.* بدل ما يكرر نفس الكود.
// يحتاج (للـ PDF و Excel): jszip, html2canvas, jspdf من cdnjs.

window.AZI = (function () {
  const CURRENCY = "ر.س";

  // تطبيق (PWA): sw.js بجذر الموقع، مكانه محسوب من مكان هذا الملف
  if ("serviceWorker" in navigator && location.protocol.startsWith("http") && document.currentScript) {
    const sw = new URL("../sw.js", document.currentScript.src);
    navigator.serviceWorker.register(sw, { scope: new URL("./", sw).pathname }).catch(() => { });
  }

  // ---------- أدوات عامة ----------
  const today = () => {
    const d = new Date();
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
  };
  function fmtDate(s) { if (!s) return ""; const [y, m, d] = s.split("-"); return d + "/" + m + "/" + y; }
  function excelSerial(s) { const ep = Date.UTC(1899, 11, 30); const [y, m, d] = s.split("-").map(Number); return Math.round((Date.UTC(y, m - 1, d) - ep) / 86400000); }
  function xmlEsc(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
  function htmlEsc(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  function money(n) { return n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function b64ToU8(b) { const bin = atob(b); const a = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) a[i] = bin.charCodeAt(i); return a; }
  function asset(key) {
    const v = window.AZI_ASSETS && window.AZI_ASSETS[key];
    if (!v) throw new Error("الملف المضمّن غير موجود: " + key);
    return v;
  }

  function downloadBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  // ---------- رسائل الخطأ والنجاح ----------
  function messages(errorEl, successEl) {
    let t;
    return {
      error(m) { errorEl.textContent = "⚠️ " + m; errorEl.classList.remove("hidden"); },
      clear() { errorEl.classList.add("hidden"); successEl.classList.add("hidden"); },
      success(m) {
        successEl.textContent = "✅ " + (m || "تم الحفظ بنجاح!"); successEl.classList.remove("hidden");
        clearTimeout(t); t = setTimeout(() => successEl.classList.add("hidden"), 6000);
      },
    };
  }

  // ---------- Excel: تعبئة قالب xlsx ----------
  function cellRe(ref) { return new RegExp('(<c r="' + ref + '"[^>]*?)(\\s+t="[^"]*")?(\\s*/>|>[\\s\\S]*?</c>)'); }
  const xlsx = {
    setNum(xml, ref, val) { return xml.replace(cellRe(ref), (_, a) => a + "><v>" + val + "</v></c>"); },
    setStr(xml, ref, val) { const v = xmlEsc(val); return xml.replace(cellRe(ref), (_, a) => a + ' t="inlineStr"><is><t xml:space="preserve">' + v + "</t></is></c>"); },
    clear(xml, ref) { return xml.replace(cellRe(ref), "$1/>"); },
    // يفتح القالب، ينفّذ edit(sheetXml) ويرجّع Blob جاهز
    async fill(templateB64, edit, sheetPath = "xl/worksheets/sheet1.xml") {
      const zip = await JSZip.loadAsync(b64ToU8(templateB64));
      const sheet = await zip.file(sheetPath).async("string");
      zip.file(sheetPath, edit(sheet));
      return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    },
  };

  // ---------- PDF: تصوير عنصر HTML إلى صفحة A4 ----------
  const pdf = {
    async fromHTML(html) {
      let area = document.getElementById("aziRenderArea");
      if (!area) { area = document.createElement("div"); area.id = "aziRenderArea"; area.className = "renderArea"; document.body.appendChild(area); }
      area.innerHTML = html;
      try {
        if (document.fonts && document.fonts.ready) { try { await document.fonts.load("700 16px Rubik"); await document.fonts.ready; } catch (e) { } }
        await new Promise(r => setTimeout(r, 200));
        const canvas = await html2canvas(area.firstElementChild, { scale: 2, useCORS: true, backgroundColor: "#ffffff" });
        const img = canvas.toDataURL("image/jpeg", 0.95);
        const doc = new window.jspdf.jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
        const pw = doc.internal.pageSize.getWidth(), ph = doc.internal.pageSize.getHeight();
        let w = pw, h = (canvas.height * w) / canvas.width;
        if (h > ph) { h = ph; w = (canvas.width * h) / canvas.height; }
        doc.addImage(img, "JPEG", (pw - w) / 2, 5, w, h);
        return doc.output("blob");
      } finally {
        area.innerHTML = "";
      }
    },
  };

  // ---------- الملف الناتج: معاينة + مشاركة واتساب + تحميل ----------
  const PDFJS = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/";
  const PDFJS_FONTS = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/standard_fonts/"; // خط Helvetica للأرقام (cdnjs ما فيه هالملفات)
  let pdfjsP;
  const loadScript = src => new Promise((res, rej) => {
    const s = document.createElement("script");
    s.src = src; s.onload = res; s.onerror = () => rej(new Error("تعذّر تحميل عارض PDF"));
    document.head.appendChild(s);
  });
  // الـ worker بينحمّل كسكربت عادي (pdfjsWorker) فبيشتغل بنفس الصفحة —
  // الـ Worker المنفصل من CDN ما بيشتغل على file:// وبيعلق أحياناً
  function loadPdfJs() {
    if (!pdfjsP) pdfjsP = loadScript(PDFJS + "pdf.min.js")
      .then(() => loadScript(PDFJS + "pdf.worker.min.js"))
      .then(() => { window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS + "pdf.worker.min.js"; return window.pdfjsLib; })
      .catch(e => { pdfjsP = null; throw e; });
    return pdfjsP;
  }

  // يعرض صفحات الـ PDF كصور (يشتغل على الموبايل كمان)، وإذا ما زبط بيعرضه بإطار المتصفح
  function preview(blob, name) {
    const url = URL.createObjectURL(blob);
    const ov = document.createElement("div");
    ov.className = "previewOverlay";
    ov.innerHTML =
      '<div class="previewBar"><span class="t"></span>' +
      '<a class="previewOpen" target="_blank" rel="noopener">فتح بتبويب ↗</a>' +
      '<button class="previewClose" type="button" aria-label="إغلاق">✕</button></div>' +
      '<div class="previewBody"><div class="previewLoading">⏳ جاري تجهيز المعاينة...</div></div>';
    ov.querySelector(".t").textContent = name;
    ov.querySelector(".previewOpen").href = url;
    const body = ov.querySelector(".previewBody");
    const close = () => { ov.remove(); document.body.style.overflow = ""; document.removeEventListener("keydown", onKey); URL.revokeObjectURL(url); };
    const onKey = e => { if (e.key === "Escape") close(); };
    ov.querySelector(".previewClose").addEventListener("click", close);
    document.addEventListener("keydown", onKey);
    document.body.style.overflow = "hidden";
    document.body.appendChild(ov);

    (async () => {
      try {
        const lib = await loadPdfJs();
        const doc = await lib.getDocument({ data: new Uint8Array(await blob.arrayBuffer()), disableFontFace: true, standardFontDataUrl: PDFJS_FONTS }).promise;
        body.innerHTML = "";
        const cssW = Math.min(body.clientWidth - 32, 820);
        for (let i = 1; i <= doc.numPages; i++) {
          const page = await doc.getPage(i);
          const base = page.getViewport({ scale: 1 });
          const vp = page.getViewport({ scale: (cssW / base.width) * (window.devicePixelRatio || 1) });
          const canvas = document.createElement("canvas");
          canvas.width = vp.width; canvas.height = vp.height;
          body.appendChild(canvas);
          await page.render({ canvasContext: canvas.getContext("2d"), viewport: vp }).promise;
        }
      } catch (e) {
        console.error(e);
        if (ov.isConnected) body.innerHTML = '<iframe title="معاينة" src="' + url + '"></iframe>';
      }
    })();
  }

  async function shareWhatsApp(blob, name, hintEl) {
    const file = new File([blob], name, { type: blob.type });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try { await navigator.share({ files: [file], title: name }); return; }
      catch (e) { if (e.name === "AbortError") return; }
    }
    // المتصفح ما بيدعم مشاركة الملفات (غالباً كمبيوتر): نزّل الملف وافتح واتساب، والمستخدم بيرفقه
    downloadBlob(blob, name);
    window.open("https://wa.me/?text=" + encodeURIComponent(name), "_blank");
    hintEl.textContent = "نزّلنا الملف وفتحنا واتساب — ارفق الملف بالمحادثة 📎";
    hintEl.classList.remove("hidden");
  }

  const result = {
    // يعرض صندوق الملف الناتج داخل el
    show(el, blob, name) {
      const isPdf = blob.type === "application/pdf";
      el.innerHTML =
        '<div class="resultBox"><div class="resultName"></div><div class="resultBtns">' +
        (isPdf ? '<button class="previewBtn" type="button">👁️ معاينة</button>' : "") +
        '<button class="waBtn" type="button">مشاركة واتساب</button>' +
        '<button class="dlBtn" type="button">⬇️ تحميل</button>' +
        '</div><div class="resultHint hidden"></div></div>';
      el.querySelector(".resultName").textContent = "📎 " + name;
      if (isPdf) el.querySelector(".previewBtn").addEventListener("click", () => preview(blob, name));
      el.querySelector(".waBtn").addEventListener("click", () => shareWhatsApp(blob, name, el.querySelector(".resultHint")));
      el.querySelector(".dlBtn").addEventListener("click", () => downloadBlob(blob, name));
    },
    clear(el) { el.innerHTML = ""; },
  };

  // ---------- زر الكريدت + الألعاب النارية ----------
  const credit = {
    mount(el) {
      el.innerHTML = '<div class="creditWrap"><button class="creditTrigger" type="button">⚡ اكبس هون ضروري ⚡</button></div>';
      el.querySelector("button").addEventListener("click", credit.open);
    },
    open() {
      const overlay = document.createElement("div");
      overlay.className = "overlay";
      overlay.innerHTML =
        '<button class="closeCredit" type="button">✕</button>' +
        "<canvas></canvas>" +
        '<div class="creditContent">' +
          '<div class="creditText">تمت البرمجة بواسطة القوي الخارق العبقري،<br>سيد الأكواد وحاكم المتغيرات بلا منازع،<br>صاحب الحلول المستحيلة والأفكار الأسطورية،<br>الذي يبرمج بسرعة الضوء ودقة الذرة</div>' +
          '<div class="nameLine"><span class="heart">❤️</span><span class="legendName">أحمد عبد العزيز</span><span class="heart h2">❤️</span></div>' +
        "</div>";
      document.body.appendChild(overlay);

      const canvas = overlay.querySelector("canvas");
      const ctx = canvas.getContext("2d");
      let particles = [], stop = false;
      const colors = ["#fbbf24", "#f472b6", "#a78bfa", "#60a5fa", "#34d399", "#f87171", "#fde68a"];
      function resize() { canvas.width = window.innerWidth; canvas.height = window.innerHeight; }
      resize(); window.addEventListener("resize", resize);

      function launch() {
        const x = Math.random() * canvas.width;
        const y = Math.random() * canvas.height * 0.5 + canvas.height * 0.1;
        const color = colors[Math.floor(Math.random() * colors.length)];
        const count = 40 + Math.floor(Math.random() * 30);
        for (let i = 0; i < count; i++) {
          const angle = (Math.PI * 2 * i) / count, speed = Math.random() * 4 + 2;
          particles.push({ x, y, vx: Math.cos(angle) * speed, vy: Math.sin(angle) * speed, life: 1, color, size: Math.random() * 3 + 1.5 });
        }
      }
      function tick() {
        if (stop) return;
        ctx.fillStyle = "rgba(7,1,15,0.25)";
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        particles.forEach(p => {
          p.x += p.vx; p.y += p.vy; p.vy += 0.04; p.life -= 0.012;
          ctx.globalAlpha = Math.max(p.life, 0);
          ctx.fillStyle = p.color;
          ctx.beginPath(); ctx.arc(p.x, p.y, p.size, 0, Math.PI * 2); ctx.fill();
        });
        ctx.globalAlpha = 1;
        particles = particles.filter(p => p.life > 0);
        requestAnimationFrame(tick);
      }
      tick(); launch();
      const iv = setInterval(launch, 650);

      overlay.querySelector(".closeCredit").addEventListener("click", () => {
        stop = true; clearInterval(iv); window.removeEventListener("resize", resize); overlay.remove();
      });
    },
  };

  return { CURRENCY, today, fmtDate, excelSerial, xmlEsc, htmlEsc, money, b64ToU8, asset, downloadBlob, messages, xlsx, pdf, result, credit };
})();
