# -*- coding: utf-8 -*-
import base64

pptx_path = r"C:\Users\Moham\.gemini\antigravity\scratch\informatique_presentation\Informatique_Cours_Complet.pptx"
html_path = r"C:\Users\Moham\.gemini\antigravity\scratch\informatique_presentation\index.html"

with open(pptx_path, "rb") as f:
    b64_pptx = base64.b64encode(f.read()).decode("utf-8")

with open(html_path, "rb") as f:
    b64_html = base64.b64encode(f.read()).decode("utf-8")

new_presentation_html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>مركز تنزيل ملفات مقياس الإعلام الآلي</title>
  <script src="https://www.gstatic.com/antigravity/web/dev/tailwindcss.min.js"></script>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800;900&display=swap');
    body {{
      font-family: 'Cairo', system-ui, sans-serif;
    }}
  </style>
</head>
<body class="bg-[#090d16] text-slate-100 p-4 md:p-8 min-h-screen flex flex-col items-center justify-center">
  <div class="max-w-3xl w-full bg-slate-900/95 border border-cyan-500/30 rounded-3xl p-6 md:p-8 shadow-2xl">
    
    <!-- Header -->
    <div class="text-center mb-6">
      <div class="w-16 h-16 rounded-2xl bg-gradient-to-tr from-cyan-500 to-blue-600 flex items-center justify-center text-3xl mx-auto mb-3 shadow-lg shadow-cyan-500/25">
        💻
      </div>
      <span class="px-3 py-1 rounded-full bg-cyan-950 text-cyan-400 border border-cyan-500/40 text-xs font-bold mb-2 inline-block">
        مقياس الإعلام الآلي | مركز التنزيل المباشر
      </span>
      <h1 class="text-2xl md:text-3xl font-black text-white">
        تحميل وتشغيل ملفات العرض التقديمي
      </h1>
      <p class="text-xs md:text-sm text-slate-300 mt-2 max-w-xl mx-auto">
        نظراً لأن متصفحات الإنترنت تحظر فتح الروابط المحلية المباشرة (file:///) لأسباب أمنية، يمكنك الآن تنزيل الملفات بنقرة زر واحدة أو فتح مجلدها في جهازك مباشرة:
      </p>
    </div>

    <!-- Quick Download Action Cards -->
    <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
      <!-- PPTX Card -->
      <div class="p-5 rounded-2xl bg-slate-800/80 border border-slate-700/70 hover:border-emerald-500/50 transition flex flex-col justify-between">
        <div>
          <div class="flex items-center justify-between mb-2">
            <span class="text-2xl">📊</span>
            <span class="text-[10px] bg-emerald-950 text-emerald-400 border border-emerald-500/40 px-2 py-0.5 rounded-full font-bold">ملف رسمي جاهز</span>
          </div>
          <h3 class="text-sm font-bold text-white mb-1">عرض PowerPoint (.pptx)</h3>
          <p class="text-xs text-slate-300 leading-relaxed mb-4">
            26 شريحة كاملة ومفصلة بتصميم احترافي (16:9)، نصوص إلقاء أكاديمية كاملة للأستاذ السيد نهمار محمد مع نماذج تطبيقية موجهة.
          </p>
        </div>
        <button onclick="downloadPPTX()" class="w-full py-2.5 px-4 rounded-xl bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-bold text-xs flex items-center justify-center gap-2 shadow-lg shadow-emerald-600/20 transition active:scale-95">
          <span>📥 تحميل ملف PowerPoint (.pptx)</span>
        </button>
      </div>

      <!-- HTML Presentation Card -->
      <div class="p-5 rounded-2xl bg-slate-800/80 border border-slate-700/70 hover:border-cyan-500/50 transition flex flex-col justify-between">
        <div>
          <div class="flex items-center justify-between mb-2">
            <span class="text-2xl">🎬</span>
            <span class="text-[10px] bg-cyan-950 text-cyan-400 border border-cyan-500/40 px-2 py-0.5 rounded-full font-bold">سينمائي وتفاعلي (26 شريحة)</span>
          </div>
          <h3 class="text-sm font-bold text-white mb-1">العرض السينمائي (HTML5)</h3>
          <p class="text-xs text-slate-300 leading-relaxed mb-4">
            شريط علوي متحرك يختفي تلقائياً عند النقر، محاكي تفاعلي لسطح المكتب Windows 11، شرح مفصل للأيقونات وأمثلة تطبيقية.
          </p>
        </div>
        <button onclick="downloadHTML()" class="w-full py-2.5 px-4 rounded-xl bg-gradient-to-r from-cyan-500 to-blue-600 hover:from-cyan-400 hover:to-blue-500 text-white font-bold text-xs flex items-center justify-center gap-2 shadow-lg shadow-cyan-500/20 transition active:scale-95">
          <span>🌐 تحميل العرض التفاعلي (index.html)</span>
        </button>
      </div>
    </div>

    <!-- Direct Windows Access Box -->
    <div class="p-4 rounded-2xl bg-slate-950/70 border border-cyan-500/30 text-xs">
      <div class="flex items-center gap-2 mb-2">
        <span class="text-cyan-400 font-bold">⚡ طريقة فورية بديلة لفتح الملفات على جهازك دون تحميل:</span>
      </div>
      <p class="text-slate-300 mb-2 leading-relaxed">
        الملفات محفوظة بالفعل على حاسوبك! اضغط في لوحة المفاتيح على <kbd class="px-1.5 py-0.5 bg-slate-800 text-cyan-300 rounded border border-slate-700 font-mono">Win + R</kbd>، ثم الصق هذا المسار واضغط <strong>Enter</strong>:
      </p>
      <div class="flex items-center gap-2">
        <input type="text" readonly value="C:\\Users\\Moham\\.gemini\\antigravity\\scratch\\informatique_presentation" class="w-full bg-slate-900 text-cyan-300 p-2 rounded-xl border border-slate-700 text-xs font-mono select-all focus:outline-none">
        <button onclick="copyPath()" id="copy-btn" class="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-white font-bold whitespace-nowrap transition">
          نسخ المسار
        </button>
      </div>
    </div>

  </div>

  <script>
    const B64_PPTX = "{b64_pptx}";
    const B64_HTML = "{b64_html}";

    function downloadPPTX() {{
      try {{
        const bin = window.atob(B64_PPTX);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) {{
          bytes[i] = bin.charCodeAt(i);
        }}
        const blob = new Blob([bytes], {{ type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }});
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'Informatique_Cours_Complet.pptx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        a.remove();
      }} catch (e) {{
        alert('تعذر التنزيل: ' + e.message);
      }}
    }}

    function downloadHTML() {{
      try {{
        const bin = window.atob(B64_HTML);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) {{
          bytes[i] = bin.charCodeAt(i);
        }}
        const blob = new Blob([bytes], {{ type: 'text/html;charset=utf-8' }});
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'Informatique_Presentation_Cinematic.html';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        a.remove();
      }} catch (e) {{
        alert('تعذر التنزيل: ' + e.message);
      }}
    }}

    function copyPath() {{
      const path = 'C:\\\\Users\\\\Moham\\\\.gemini\\\\antigravity\\\\scratch\\\\informatique_presentation';
      navigator.clipboard.writeText(path).then(() => {{
        const btn = document.getElementById('copy-btn');
        btn.textContent = 'تم النسخ! ✓';
        setTimeout(() => {{ btn.textContent = 'نسخ المسار'; }}, 2000);
      }});
    }}
  </script>
</body>
</html>
"""

brain_path = r"C:\Users\Moham\.gemini\antigravity\brain\8e28d94e-0064-4605-ad65-07261bf39388\presentation.html"
with open(brain_path, "w", encoding="utf-8") as f:
    f.write(new_presentation_html)

print("Updated presentation.html in brain successfully.")
