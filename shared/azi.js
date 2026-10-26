// ===== AZIflow Bots — أدوات مشتركة لكل البوتات =====
// كل بوت يحمّل هذا الملف ويستخدم AZI.* بدل ما يكرر نفس الكود.
// يحتاج (للـ PDF و Excel): jszip, html2canvas, jspdf من cdnjs.

window.AZI = (function () {
  const CURRENCY = "ر.س";

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

  // ---------- مجلد الحفظ (مشترك بين كل البوتات ويُحفظ بين الجلسات) ----------
  const folder = (function () {
    const supported = "showDirectoryPicker" in window;
    const DB = "azi-bots", STORE = "kv", KEY = "rootDir";
    let handle = null, infoEl = null, btnEl = null;

    function idb(mode, op) {
      return new Promise((res, rej) => {
        const r = indexedDB.open(DB, 1);
        r.onupgradeneeded = () => r.result.createObjectStore(STORE);
        r.onerror = () => rej(r.error);
        r.onsuccess = () => {
          const tx = r.result.transaction(STORE, mode);
          const q = op(tx.objectStore(STORE));
          tx.oncomplete = () => res(q && q.result);
          tx.onerror = () => rej(tx.error);
        };
      });
    }
    const load = () => idb("readonly", s => s.get(KEY)).catch(() => null);
    const store = h => idb("readwrite", s => s.put(h, KEY)).catch(() => { });

    function render() {
      if (!infoEl) return;
      if (!supported) {
        infoEl.textContent = "⚠️ متصفحك لا يدعم اختيار المجلد. استخدم Chrome أو Edge على الكمبيوتر. سيتم تنزيل الملفات في مجلد التنزيلات بدلاً من ذلك.";
        btnEl.style.display = "none";
        return;
      }
      if (handle) {
        infoEl.innerHTML = "✅ المجلد المختار: <b></b><br>سيُنشأ مجلد فرعي باسم تاريخ كل ملف داخله.";
        infoEl.querySelector("b").textContent = handle.name;
        btnEl.textContent = "تغيير المجلد";
      } else {
        infoEl.textContent = "📁 لم يتم اختيار مجلد الحفظ بعد. اختر مجلداً رئيسياً وسيُنشئ البوت مجلداً لكل تاريخ تلقائياً.";
        btnEl.textContent = "اختيار المجلد";
      }
    }

    async function pick() {
      try {
        handle = await window.showDirectoryPicker({ mode: "readwrite" });
        await store(handle);
        render();
      } catch (e) { }
    }

    return {
      supported,
      // يرسم صندوق المجلد داخل العنصر المعطى
      async mount(el) {
        el.innerHTML = '<div class="folderBox"><div class="folderInfo"></div><button class="folderBtn" type="button"></button></div>';
        infoEl = el.querySelector(".folderInfo"); btnEl = el.querySelector(".folderBtn");
        btnEl.addEventListener("click", pick);
        render();
        if (supported) { handle = (await load()) || null; render(); }
      },
      // استدعِها أول شي داخل الكبسة (قبل أي انتظار طويل): بترجّع true إذا المجلد جاهز للكتابة
      async ready() {
        if (!handle) return false;
        try {
          const opts = { mode: "readwrite" };
          if ((await handle.queryPermission(opts)) === "granted") return true;
          return (await handle.requestPermission(opts)) === "granted";
        } catch (e) { return false; }
      },
      // يحفظ داخل <المجلد>/<sub>/ إذا متاح، وإلا ينزّل الملف. يرجّع "folder" أو "download"
      async save(blob, filename, sub) {
        if (handle && (await this.ready())) {
          let dir = handle;
          if (sub) dir = await handle.getDirectoryHandle(sub, { create: true });
          const fh = await dir.getFileHandle(filename, { create: true });
          const w = await fh.createWritable();
          await w.write(blob);
          await w.close();
          return "folder";
        }
        downloadBlob(blob, filename);
        return "download";
      },
    };
  })();

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

  return { CURRENCY, today, fmtDate, excelSerial, xmlEsc, htmlEsc, money, b64ToU8, asset, downloadBlob, messages, xlsx, pdf, folder, credit };
})();
