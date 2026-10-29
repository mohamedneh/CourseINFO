# -*- coding: utf-8 -*-
"""
Python script to generate the 26-slide index.html with interactive auto-collapsing header,
desktop simulator, icon breakdown, and embedded PPTX base64.
"""

import base64
import os

pptx_path = r"C:\Users\Moham\.gemini\antigravity\scratch\informatique_presentation\Informatique_Cours_Complet.pptx"
with open(pptx_path, "rb") as f:
    b64_pptx = base64.b64encode(f.read()).decode("utf-8")

html_code = '''<!DOCTYPE html>
<html lang="ar" dir="rtl" class="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>مقياس الإعلام الآلي | إعداد الأستاذ نهمار محمد - العرض الشامل التفاعلي (26 شريحة)</title>
  <!-- Tailwind CSS via official allowed gstatic CDN -->
  <script src="https://www.gstatic.com/antigravity/web/dev/tailwindcss.min.js"></script>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;500;600;700;800;900&family=Tajawal:wght@400;500;700;800;900&family=JetBrains+Mono:wght@400;500;600;700&display=swap');

    :root {
      --font-arabic: 'Cairo', 'Tajawal', system-ui, -apple-system, sans-serif;
      --font-mono: 'JetBrains Mono', monospace;
    }

    body {
      font-family: var(--font-arabic);
      background-color: #070b14;
      color: #f1f5f9;
      overflow: hidden;
      user-select: none;
    }

    #bg-canvas {
      position: fixed;
      top: 0;
      left: 0;
      width: 100vw;
      height: 100vh;
      z-index: 0;
      pointer-events: none;
    }

    /* Interactive Auto-Collapsing Header */
    #main-header {
      transition: transform 0.35s cubic-bezier(0.16, 1, 0.3, 1), opacity 0.3s ease;
      transform: translateY(0);
      opacity: 1;
    }

    #main-header.header-hidden {
      transform: translateY(-100%);
      opacity: 0;
      pointer-events: none;
    }

    /* Reveal pill when header is hidden */
    #header-reveal-pill {
      transition: opacity 0.3s ease, transform 0.3s ease;
      opacity: 0;
      pointer-events: none;
      transform: translateY(-10px);
    }

    #header-reveal-pill.pill-visible {
      opacity: 1;
      pointer-events: auto;
      transform: translateY(0);
    }

    /* Slide transition containers with ample top padding so titles are NEVER covered */
    .slide {
      position: absolute;
      inset: 0;
      opacity: 0;
      visibility: hidden;
      transform: scale(0.95) translateY(20px);
      transition: all 0.6s cubic-bezier(0.16, 1, 0.3, 1);
      z-index: 10;
      display: flex;
      flex-direction: column;
      padding: 5rem 3rem 5.8rem 3rem; /* 5rem top padding ensures titles are 100% visible */
      overflow-y: auto;
    }

    .slide.active {
      opacity: 1;
      visibility: visible;
      transform: scale(1) translateY(0);
      z-index: 20;
    }

    .slide.prev {
      opacity: 0;
      visibility: hidden;
      transform: scale(1.04) translateY(-16px);
    }

    .glass-card {
      background: rgba(18, 25, 42, 0.72);
      backdrop-filter: blur(20px);
      -webkit-backdrop-filter: blur(20px);
      border: 1px solid rgba(255, 255, 255, 0.08);
      box-shadow: 0 20px 40px -15px rgba(0, 0, 0, 0.6), inset 0 1px 0 0 rgba(255, 255, 255, 0.1);
      transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
    }

    .glass-card:hover {
      border-color: rgba(56, 189, 248, 0.35);
      transform: translateY(-3px);
      box-shadow: 0 25px 50px -12px rgba(56, 189, 248, 0.18), inset 0 1px 0 0 rgba(56, 189, 248, 0.25);
    }

    .glass-header {
      background: rgba(13, 18, 32, 0.88);
      backdrop-filter: blur(20px);
      border-bottom: 1px solid rgba(255, 255, 255, 0.08);
    }

    .glass-footer {
      background: rgba(10, 15, 26, 0.90);
      backdrop-filter: blur(24px);
      border-top: 1px solid rgba(255, 255, 255, 0.08);
    }

    .read-aloud-box {
      background: linear-gradient(135deg, rgba(30, 41, 59, 0.85) 0%, rgba(15, 23, 42, 0.95) 100%);
      border-right: 4px solid #38bdf8;
      border-radius: 14px;
      padding: 0.9rem 1.2rem;
      margin-bottom: 0.8rem;
      box-shadow: 0 8px 24px rgba(0, 0, 0, 0.35);
    }

    .teacher-badge {
      display: inline-flex;
      align-items: center;
      gap: 0.35rem;
      padding: 0.2rem 0.65rem;
      border-radius: 9999px;
      background: rgba(56, 189, 248, 0.15);
      border: 1px solid rgba(56, 189, 248, 0.35);
      color: #7dd3fc;
      font-size: 0.72rem;
      font-weight: 700;
      margin-bottom: 0.4rem;
    }

    .student-badge {
      display: inline-flex;
      align-items: center;
      gap: 0.35rem;
      padding: 0.2rem 0.65rem;
      border-radius: 9999px;
      background: rgba(52, 211, 153, 0.15);
      border: 1px solid rgba(52, 211, 153, 0.35);
      color: #6ee7b7;
      font-size: 0.72rem;
      font-weight: 700;
      margin-bottom: 0.4rem;
    }

    .analogy-badge {
      display: inline-flex;
      align-items: center;
      gap: 0.35rem;
      padding: 0.2rem 0.65rem;
      border-radius: 9999px;
      background: rgba(251, 191, 36, 0.15);
      border: 1px solid rgba(251, 191, 36, 0.35);
      color: #fde047;
      font-size: 0.72rem;
      font-weight: 700;
      margin-bottom: 0.4rem;
    }

    .slide::-webkit-scrollbar {
      width: 6px;
    }
    .slide::-webkit-scrollbar-track {
      background: rgba(15, 23, 42, 0.5);
    }
    .slide::-webkit-scrollbar-thumb {
      background: rgba(56, 189, 248, 0.35);
      border-radius: 9999px;
    }

    .code-font {
      font-family: var(--font-mono);
      direction: ltr;
      display: inline-block;
    }
  </style>
</head>
<body class="h-screen w-screen relative overflow-hidden select-none bg-[#070b14] text-slate-100">

  <!-- Background Particle Canvas -->
  <canvas id="bg-canvas"></canvas>

  <!-- Ambient Glows -->
  <div class="fixed top-[-10vw] right-[-10vw] w-[45vw] h-[45vw] rounded-full bg-cyan-600/10 blur-[140px] pointer-events-none z-0"></div>
  <div class="fixed bottom-[-10vw] left-[-10vw] w-[45vw] h-[45vw] rounded-full bg-indigo-600/12 blur-[140px] pointer-events-none z-0"></div>
  <div class="fixed top-1/2 left-1/3 w-[30vw] h-[30vw] rounded-full bg-emerald-600/6 blur-[130px] pointer-events-none z-0"></div>

  <!-- Header Reveal Hover/Click Area & Floating Pill -->
  <div id="header-hover-zone" onmouseenter="showHeader()" class="fixed top-0 left-0 right-0 h-3 z-[65] cursor-pointer"></div>
  <div id="header-reveal-pill" onclick="showHeader()" class="fixed top-1 left-1/2 -translate-x-1/2 z-[70] bg-slate-900/90 hover:bg-cyan-600/80 text-cyan-300 hover:text-white px-4 py-1 rounded-full border border-cyan-500/40 text-[11px] font-bold cursor-pointer transition flex items-center gap-1.5 shadow-xl">
    <span>⬇️</span> <span>إظهار شريط المقياس</span>
  </div>

  <!-- Top Global Header Bar (Interactive Auto-Collapsing) -->
  <header id="main-header" class="fixed top-0 left-0 right-0 h-16 glass-header z-50 flex items-center justify-between px-6">
    <div class="flex items-center gap-3.5">
      <div class="w-10 h-10 rounded-xl bg-gradient-to-tr from-cyan-500 via-sky-600 to-blue-700 flex items-center justify-center shadow-lg shadow-cyan-500/25 text-white font-black text-xl">
        💻
      </div>
      <div>
        <h1 class="text-sm font-extrabold text-white tracking-wide flex items-center gap-2">
          مقياس الإعلام الآلي (Informatique)
          <span class="text-[11px] font-bold px-2.5 py-0.5 rounded-full bg-cyan-500/20 text-cyan-300 border border-cyan-500/35">
            المنهاج الشامل والمفصل
          </span>
        </h1>
        <p class="text-xs text-slate-300 font-semibold flex items-center gap-2">
          <span>أستاذ المقياس: <strong class="text-amber-400 font-bold">السيد نهمار محمد</strong></span>
          <span class="text-slate-500">•</span>
          <span class="text-slate-400">محاضرات موجهة + أعمال تطبيقية</span>
        </p>
      </div>
    </div>

    <!-- Progress Tracker -->
    <div class="flex items-center gap-4">
      <div class="hidden lg:flex flex-col items-end">
        <span class="text-xs font-bold text-cyan-400" id="slide-tracker">شريحة 1 من 26</span>
        <span class="text-[11px] text-slate-300 font-medium truncate max-w-xs" id="slide-title-preview">الافتتاحية والأهداف</span>
      </div>
      <div class="w-36 h-2 bg-slate-800 rounded-full overflow-hidden border border-slate-700/60">
        <div id="progress-bar" class="h-full bg-gradient-to-l from-cyan-400 via-sky-500 to-indigo-500 transition-all duration-500 w-[3.84%]"></div>
      </div>

      <!-- Action Buttons -->
      <div class="flex items-center gap-1.5 border-r border-slate-700/60 pr-3 mr-1">
        <button id="btn-sound" onclick="toggleSound()" title="تشغيل / كتم المؤثرات الصوتية" class="w-9 h-9 rounded-lg bg-slate-800/80 hover:bg-slate-700 text-slate-300 hover:text-white flex items-center justify-center transition border border-slate-700/50">
          <span id="sound-icon">🔊</span>
        </button>
        <button id="btn-teacher-speech" onclick="toggleTeacherSpeechModal()" title="نص إلقاء الأستاذ نهمار محمد" class="w-9 h-9 rounded-lg bg-cyan-950/80 hover:bg-cyan-900 text-cyan-300 hover:text-white flex items-center justify-center transition border border-cyan-500/40">
          🎙️
        </button>
        <button id="btn-hide-header" onclick="hideHeader()" title="إخفاء الشريط العلوي لتركيز العنوان" class="w-9 h-9 rounded-lg bg-slate-800/80 hover:bg-slate-700 text-slate-300 hover:text-white flex items-center justify-center transition border border-slate-700/50">
          ▲
        </button>
        <button id="btn-fullscreen" onclick="toggleFullScreen()" title="ملء الشاشة السينمائي (F)" class="w-9 h-9 rounded-lg bg-slate-800/80 hover:bg-slate-700 text-slate-300 hover:text-white flex items-center justify-center transition border border-slate-700/50">
          ⛶
        </button>
      </div>
    </div>
  </header>

  <!-- ==================== MAIN SLIDES VIEWPORT (Clicking here hides header!) ==================== -->
  <main id="slides-viewport" onclick="handleViewportClick(event)" class="relative w-full h-full">

    <!-- SLIDE 1 -->
    <section class="slide active" data-index="0" data-tag="الافتتاحية الرسمية" data-title="الترحيب وبطاقة المقياس | أ. نهمار محمد">
      <div class="max-w-5xl mx-auto my-auto text-center flex flex-col items-center">
        <div class="inline-flex items-center gap-2.5 px-4 py-1.5 rounded-full bg-cyan-950/70 border border-cyan-500/40 text-cyan-300 text-xs font-bold mb-4">
          <span class="w-2.5 h-2.5 rounded-full bg-cyan-400 animate-ping"></span>
          التعليم العالي والبحث العلمي • مقياس الإعلام الآلي التأسيسي (Informatique)
        </div>
        <h1 class="text-3xl lg:text-5xl font-black text-white leading-tight mb-4 tracking-tight">
          المدخل الأكاديمي الشامل للحوسبة وهندسة نظام <span class="bg-gradient-to-r from-cyan-400 via-sky-300 to-blue-500 bg-clip-text text-transparent">Windows</span>
        </h1>
        <div class="glass-card px-6 py-3 rounded-2xl border-cyan-500/40 mb-5 flex items-center gap-3 bg-gradient-to-r from-slate-900/90 via-cyan-950/40 to-slate-900/90">
          <div class="w-12 h-12 rounded-xl bg-gradient-to-tr from-amber-500 to-amber-600 flex items-center justify-center text-2xl shadow-lg">👨‍🏫</div>
          <div class="text-right">
            <div class="text-xs text-slate-300 font-semibold">إعداد وتقديم أستاذ المقياس:</div>
            <div class="text-lg font-black text-amber-300">السيد نهمار محمد (M. Nehmar Mohamed)</div>
          </div>
        </div>
        <div class="read-aloud-box text-right w-full max-w-4xl border-r-amber-400">
          <div class="teacher-badge">🗣️ نص كلمة الأستاذ نهمار محمد للطلبة (اقرأ هذا النص مباشرة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed font-normal">
            "بسم الله الرحمن الرحيم، والصلاة والسلام على رسوله الكريم. مرحباً بكم جميعاً أعزائي الطلبة في أولى محاضرات مقياس الإعلام الآلي. هذا المقياس صُمم لكي يرافقكم خطوة بخطوة من الصفر. لا تقلقوا إذا كنتم لا تملكون خلفية تقنية مسبقة؛ سنفكك اليوم كل المفاهيم المعقدة بأمثلة وتشبيهات بسيطة من واقع حياتنا اليومية، وسننتقل معاً إلى الجانب التطبيقي المباشر على الحواسيب لنتدرب على كل حركة وأمر وخصيصة في نظام التشغيل Windows."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4 w-full max-w-4xl mb-6">
          <div class="glass-card p-4 rounded-2xl text-right border-t-2 border-t-cyan-500">
            <h3 class="text-sm font-bold text-white mb-1">1. تاريخ الحوسبة والأجيال</h3>
            <p class="text-xs text-slate-300 leading-relaxed">فهم كيفية تحول الحاسوب من وحش حديدي بوزن 30 طناً إلى شريحة صغيرة في جيوبنا.</p>
          </div>
          <div class="glass-card p-4 rounded-2xl text-right border-t-2 border-t-emerald-500">
            <h3 class="text-sm font-bold text-white mb-1">2. هندسة المعالجة والإقلاع</h3>
            <p class="text-xs text-slate-300 leading-relaxed">كيف يفكر الحاسوب بدورة IPOS ولغة 0 و 1، وماذا يفعل الـ BIOS عند تشغيل الجهاز.</p>
          </div>
          <div class="glass-card p-4 rounded-2xl text-right border-t-2 border-t-amber-500">
            <h3 class="text-sm font-bold text-white mb-1">3. احتراف بيئة Windows 11</h3>
            <p class="text-xs text-slate-300 leading-relaxed">ورشات تطبيقية مباشرة أمامكم: إدارة سطح المكتب، ضبط الوقت واللغات، والتحكم الشجري بالملفات.</p>
          </div>
        </div>
        <button onclick="nextSlide()" class="px-8 py-3 rounded-xl bg-gradient-to-r from-cyan-500 via-sky-600 to-blue-600 hover:from-cyan-400 hover:to-sky-500 text-white font-bold text-sm shadow-xl shadow-cyan-500/25 flex items-center gap-2.5 transition transform hover:scale-105 active:scale-95">
          <span>ابدأ المحاضرة والشرح مع الطلبة</span> <span>←</span>
        </button>
      </div>
    </section>

    <!-- SLIDE 2 -->
    <section class="slide" data-index="1" data-tag="المحور الأول: المفاهيم التأسيسية" data-title="ما هو الحاسوب؟ والفرق بين البيانات والمعلومات">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">المحور الأول: المفاهيم التأسيسية</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">ما هو الحاسوب (Computer)؟ وما هو علم الإعلام الآلي (Informatique)؟</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "أعزائي الطلبة، قبل أن نتحدث عن الأجهزة، دعونا نفهم كلمة 'إعلام آلي'. في اللغة الفرنسية، كلمة <strong>Informatique</strong> هي دمج لكلمتين: <em>Information</em> (معلومات) و <em>Automatique</em> (آلي). أي أن الإعلام الآلي هو: <strong>علم معالجة المعلومات بطريقة آلية وبدون تدخل بشري يدوي معقد</strong>. والحاسوب (Computer) ليس جهازاً ذكياً بطبعه، بل هو آلة إلكترونية تطيع الأوامر الرياضية بسرعة خيالية تفوق عقل الإنسان بمليارات المرات."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4 my-auto">
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-cyan-400">
            <h3 class="text-sm font-bold text-white mb-2 flex items-center gap-2"><span>📊</span> الفارق الجوهري: البيانات مقابل المعلومات</h3>
            <div class="space-y-3 text-xs">
              <div class="p-3 rounded-xl bg-slate-900/80 border border-slate-700/60">
                <strong class="text-cyan-400 block mb-1">1. البيانات (Data / Données):</strong>
                <p class="text-slate-300 leading-relaxed">الحقائق والأرقام الخام غير المنظمة التي لا تعطي معنى بمفردها.<br><span class="text-slate-400 font-semibold">• مثال: الأرقام (14، 18، 12) لا تفهم منها شيئاً حتى الآن!</span></p>
              </div>
              <div class="p-3 rounded-xl bg-slate-900/80 border border-slate-700/60">
                <strong class="text-emerald-400 block mb-1">2. المعلومات (Information):</strong>
                <p class="text-slate-300 leading-relaxed">البيانات بعد معالجتها وتنظيمها وترتيبها لتصبح مفيدة.<br><span class="text-slate-400 font-semibold">• مثال: "معدل علامات الطالب أحمد هو 14.66 وهو ناجح". الآن أصبح لدينا قرار ومعنى!</span></p>
              </div>
            </div>
          </div>
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-amber-400 flex flex-col justify-between">
            <div>
              <div class="analogy-badge">💡 تشبيه المطبخ البسيط:</div>
              <h3 class="text-sm font-bold text-white mb-2">كيف يشبه الحاسوب مطبخاً متكاملاً؟</h3>
              <ul class="text-xs text-slate-300 space-y-2 leading-relaxed">
                <li><strong class="text-cyan-400">• البيانات الخام (Data):</strong> تشبه الخضار واللحم غير المطبوخ (مواد أولية لا تؤكل بمفردها).</li>
                <li><strong class="text-amber-400">• المعالجة (Processing):</strong> تشبه عمل الطاهي في التقطيع والطهي والمزج.</li>
                <li><strong class="text-emerald-400">• المعلومات (Information):</strong> هي الوجبة الشهية الجاهزة للأكل!</li>
              </ul>
            </div>
            <div class="p-3 rounded-xl bg-slate-900/90 text-xs text-amber-300 border border-amber-500/30 mt-3">
              📌 <strong>خلاصة الأستاذ للطلبة:</strong> الحاسوب لا يبتكر المعلومات من العدم؛ هو فقط يأخذ بياناتكم الخام ويطبخها وينسقها بسرعة ليقدم لكم النتيجة!
            </div>
          </div>
        </div>
        <div class="glass-card p-2.5 rounded-xl bg-slate-900/60 text-xs text-slate-300 text-center">
          الأستاذ نهمار محمد: "تذكروا دائماً: الحاسوب ينفذ ما تأمره به، فإن أخطأ فالخطأ في المدخلات التي أدخلها الإنسان!"
        </div>
      </div>
    </section>

    <!-- SLIDE 3 -->
    <section class="slide" data-index="2" data-tag="المحور الأول: أجيال الحواسيب" data-title="الجيل الأول: الصمامات المفرغة (1940 - 1956)">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">المحور الأول: تاريخ الحوسبة</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">الجيل الأول (1940 - 1956): الصمامات المفرغة (Vacuum Tubes)</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "طلبتي الأعزاء، هل رأيتم من قبل مصباح الغرفة الزجاجي القديم الذي يحوي سلكاً متوهجاً من الداخل؟ هذا هو بالضبط <strong>الصمام المفرغ (Vacuum Tube)</strong>! في الأربعينيات، لم تكن هناك معالجات سيليكون صغيرة. استخدم المهندسون آلاف المصابيح الزجاجية المفرغة من الهواء. إذا أضاء المصباح يمر تيار كهربائي ويعبر عن الرقم (1)، وإذا انطفأ يعبر عن الرقم (0). بهذا المبدأ البدائي بني أول حاسوب عملاق في العالم وهو حاسوب <strong>ENIAC</strong>."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4 my-auto">
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-red-400">
            <h3 class="text-sm font-bold text-white mb-2 flex items-center gap-2"><span>🦕</span> ديناصور الحواسيب: حاسوب ENIAC (حقائق مذهلة)</h3>
            <ul class="text-xs text-slate-300 space-y-2 leading-relaxed">
              <li><strong class="text-red-400">• الوزن المهول:</strong> كان يزن أكثر من <strong>30 طناً</strong> (وزن 20 سيارة مجتمعة!).</li>
              <li><strong class="text-amber-400">• المساحة والحرارة:</strong> قاعة كاملة بمساحة 167 متراً مربعاً، و 18,000 صمام زجاجي تطلق حرارة خيالية.</li>
              <li><strong class="text-cyan-400">• استهلاك الكهرباء:</strong> 150 كيلوواط؛ قيل إن أضواء مدينة فيلادلفيا كانت تخفت عند تشغيله!</li>
              <li><strong class="text-emerald-400">• السرعة:</strong> 5000 عملية جمع في الثانية وتتم برمجته بتوصيل أسلاك باليد.</li>
            </ul>
          </div>
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-amber-400 flex flex-col justify-between">
            <div>
              <div class="analogy-badge">🐛 قصة تاريخية مسلية:</div>
              <h3 class="text-sm font-bold text-white mb-2">من أين جاءت كلمة "Bug" (خطأ برمجي)؟</h3>
              <p class="text-xs text-slate-300 leading-relaxed mb-3">
                في عام 1947، توقف حاسوب Harvard Mark II فجأة، وعندما فحصت العالمة <strong>غريس هوبر</strong> الصمامات، وجدت <strong>حشرة عثة حقيقية (Moth)</strong> انحشرت بين الأقطاب فاحترقت وأوقفت النظام! ألصقتها في السجل وكتبت: <em>"First actual case of bug being found"</em> ومنذ ذلك اليوم سمي كل خطأ تقني <strong>Bug</strong> وتصليحه <strong>Debugging</strong>!
              </p>
            </div>
            <div class="p-2.5 rounded-xl bg-slate-900/80 text-[11px] text-amber-300 border border-amber-500/30">
              📌 عيوب الجيل الأول: الحجم الهائل، استهلاك الكهرباء، واحتراق الصمامات يومياً.
            </div>
          </div>
        </div>
        <div class="glass-card p-2.5 rounded-xl bg-slate-900/60 text-xs text-slate-300 text-center">
          الأستاذ نهمار محمد: "كان المهندسون يقضون ساعات في البحث عن الصمام المحترق لاستبداله يدوياً!"
        </div>
      </div>
    </section>

    <!-- SLIDE 4 -->
    <section class="slide" data-index="3" data-tag="المحور الأول: أجيال الحواسيب" data-title="الجيل الثاني: ثورة الترانزستور (1956 - 1963)">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">المحور الأول: تاريخ الحوسبة</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">الجيل الثاني (1956 - 1963): ثورة الترانزستور (The Transistor)</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "جاء الاختراع الذي غير وجه الأرض: <strong>الترانزستور (Transistor)</strong>. استبدلنا المصباح الزجاجي بمفتاح إلكتروني صلب مصنوع من مادة السيليكون (رمال الشاطئ المنقاة!). لا ينكسر، لا يولد حرارة شديدة، وحجمه أصغر بـ 200 مرة من الصمام! جعل الحواسيب بحجم خزانة ملابس بدل قاعة كاملة."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4 my-auto">
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-emerald-400">
            <h3 class="text-sm font-bold text-white mb-2 flex items-center gap-2"><span>⚡</span> معجزة الترانزستور العتادية</h3>
            <div class="space-y-2.5 text-xs">
              <div class="p-2.5 rounded-xl bg-slate-900/80 border border-slate-700/60">
                <strong class="text-emerald-400 block mb-0.5">• متانة خارقة:</strong>
                <span class="text-slate-300">يعمل لسنوات طويلة دون انقطاع، وقضى على الأعطال اليومية.</span>
              </div>
              <div class="p-2.5 rounded-xl bg-slate-900/80 border border-slate-700/60">
                <strong class="text-cyan-400 block mb-0.5">• استهلاك طاقة ضئيل:</strong>
                <span class="text-slate-300">يستهلك جزءاً يسيراً من الواط ولا يحتاج محطات تبريد عملاقة.</span>
              </div>
              <div class="p-2.5 rounded-xl bg-slate-900/80 border border-slate-700/60">
                <strong class="text-amber-400 block mb-0.5">• تصغير الحجم:</strong>
                <span class="text-slate-300">أصبحت الحواسيب بحجم غرف المكاتب وخزائن الشركات مثل IBM 1401.</span>
              </div>
            </div>
          </div>
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-sky-400 flex flex-col justify-between">
            <div>
              <div class="teacher-badge">💻 القفزة البرمجية:</div>
              <h3 class="text-sm font-bold text-white mb-2">وداعاً لسلاسل الأصفار والآحاد المعقدة!</h3>
              <ul class="text-xs text-slate-300 space-y-2 leading-relaxed">
                <li><strong class="text-sky-300">1. لغة التجميع (Assembly):</strong> استخدام اختصارات مثل (ADD و SUB).</li>
                <li><strong class="text-emerald-300">2. اللغات عالية المستوى:</strong> ولادة FORTRAN للهندسة، و COBOL للمصارف والشركات.</li>
              </ul>
            </div>
            <div class="p-2.5 rounded-xl bg-slate-900/80 text-[11px] text-sky-300 border border-sky-500/30">
              📌 بفضل الترانزستور، خرج الحاسوب من المعامل العسكرية ليدخل كبرى الشركات والجامعات.
            </div>
          </div>
        </div>
        <div class="glass-card p-2.5 rounded-xl bg-slate-900/60 text-xs text-slate-300 text-center">
          الأستاذ نهمار محمد: "نال مخترعو الترانزستور الثلاثة جائزة نوبل في الفيزياء سنة 1956!"
        </div>
      </div>
    </section>

    <!-- SLIDE 5 -->
    <section class="slide" data-index="4" data-tag="المحور الأول: أجيال الحواسيب" data-title="الجيل الثالث والرابع: الدارات المتكاملة والمعالج الدقيق">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">المحور الأول: تاريخ الحوسبة</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">الجيل الثالث والرابع: الدارات المتكاملة (IC) والمعالجات الدقيقة (Microprocessors)</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "في الجيل الثالث (1964)، طبع المهندسون آلاف الترانزستورات دفعة واحدة على شريحة سيليكون واحدة بحجم ظفر الأصبع: <strong>الدارة المتكاملة (Integrated Circuit - IC)</strong>. وفي الجيل الرابع (1971)، طبعوا المعالج كاملاً مع كل دوائره على رقاقة واحدة: فولد <strong>المعالج الدقيق (Microprocessor)</strong> ومعه الحاسوب الشخصي (PC)."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4 my-auto">
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-cyan-400">
            <h3 class="text-sm font-bold text-white mb-2">الجيل الثالث (1964 - 1971): رقاقة السيليكون (IC)</h3>
            <ul class="text-xs text-slate-300 space-y-2 leading-relaxed">
              <li><strong class="text-cyan-400">• دمج آلاف المكونات:</strong> التخلي عن الأسلاك اليدوية واعتماد شرائح السيليكون الدقيقة.</li>
              <li><strong class="text-amber-400">• لوحة المفاتيح والشاشة:</strong> وداعاً للبطاقات المثقبة؛ ولأول مرة يكتب المستخدم ويرى النتيجة على شاشة.</li>
              <li><strong class="text-emerald-400">• أنظمة التشغيل التشاركية (Time-Sharing):</strong> تشغيل برامج متعددة وتقسيم الوقت.</li>
            </ul>
          </div>
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-indigo-400">
            <h3 class="text-sm font-bold text-white mb-2">الجيل الرابع (1971 - الحاضر): المعالج الدقيق (VLSI)</h3>
            <ul class="text-xs text-slate-300 space-y-2 leading-relaxed">
              <li><strong class="text-indigo-400">• رقاقة Intel 4004:</strong> أول معالج دقيق تجاري يضم كل مكونات الـ CPU في رقاقة واحدة.</li>
              <li><strong class="text-sky-400">• الحاسوب الشخصي (PC):</strong> أجهزة مثل IBM PC؛ تحول الحاسوب إلى جهاز يوضع فوق مكتب أي طالب وموظف.</li>
              <li><strong class="text-emerald-400">• الهواتف الذكية:</strong> معالج هاتفك اليوم يضم أكثر من 15 مليار ترانزستور نانوي!</li>
            </ul>
          </div>
        </div>
        <div class="glass-card p-2.5 rounded-xl bg-slate-900/60 text-xs text-slate-300 flex items-center justify-between">
          <span>قانون مور (Moore's Law): "عدد الترانزستورات يتضاعف كل عامين تقريباً."</span>
          <span class="text-cyan-400 font-bold">الجيل الخامس: الذكاء الاصطناعي والحوسبة الكمومية</span>
        </div>
      </div>
    </section>

    <!-- SLIDE 6 -->
    <section class="slide" data-index="5" data-tag="المحور الأول: تصنيف الحواسيب" data-title="أنواع الحواسيب وتصنيفاتها وفق القوة والحجم">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">المحور الأول: تصنيف الحواسيب</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">أنواع الحواسيب الحديثة: كيف نصنف الأجهزة من حولنا؟</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "كلمة حاسوب تشمل تصنيفات متعددة تختلف بحسب حجمها وقدرتها الحسابية والغرض الذي صنعت من أجله. سنقسمها علمياً إلى 4 عائلات كبرى من الأقوى والأضخم إلى الأصغر حجماً."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-4 gap-3.5 my-auto">
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-cyan-400 flex flex-col justify-between">
            <div>
              <div class="w-10 h-10 rounded-lg bg-cyan-500/20 text-cyan-400 flex items-center justify-center text-xl mb-2">🚀</div>
              <h3 class="text-xs font-bold text-white mb-1">الحواسيب العملاقة</h3>
              <p class="text-[11px] text-cyan-300 font-semibold mb-2">Supercomputers</p>
              <p class="text-xs text-slate-300 leading-relaxed mb-2">تقاس سرعتها بالـ Petaflops و Exaflops (ملايين المليارات من العمليات في الثانية).</p>
            </div>
            <div class="text-[11px] bg-slate-900/80 text-cyan-300 p-2 rounded-lg border border-cyan-500/30">أمثلة: التنبؤ بالأعاصير، محاكاة نووية، وأبحاث الجينوم (Frontier).</div>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-amber-400 flex flex-col justify-between">
            <div>
              <div class="w-10 h-10 rounded-lg bg-amber-500/20 text-amber-400 flex items-center justify-center text-xl mb-2">🏦</div>
              <h3 class="text-xs font-bold text-white mb-1">الحواسيب الكبيرة</h3>
              <p class="text-[11px] text-amber-300 font-semibold mb-2">Mainframe Computers</p>
              <p class="text-xs text-slate-300 leading-relaxed mb-2">خوادم لا تنطفئ لعقود (موثوقية 99.999%) لمعالجة ملايين المعاملات المالية المتزامنة في نفس اللحظة.</p>
            </div>
            <div class="text-[11px] bg-slate-900/80 text-amber-300 p-2 rounded-lg border border-amber-500/30">أمثلة: البنوك والتحويلات المالية العالمية، وحجوزات الطيران.</div>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-emerald-400 flex flex-col justify-between">
            <div>
              <div class="w-10 h-10 rounded-lg bg-emerald-500/20 text-emerald-400 flex items-center justify-center text-xl mb-2">💻</div>
              <h3 class="text-xs font-bold text-white mb-1">الحواسيب الشخصية</h3>
              <p class="text-[11px] text-emerald-300 font-semibold mb-2">PCs & Laptops</p>
              <p class="text-xs text-slate-300 leading-relaxed mb-2">مصممة لاستخدام الفرد الواحد في المكاتب والمنازل والجامعات لإنجاز البحوث والتصميم والبرمجة.</p>
            </div>
            <div class="text-[11px] bg-slate-900/80 text-emerald-300 p-2 rounded-lg border border-emerald-500/30">أمثلة: المكتبي (Desktop)، المحمول (Laptop)، واللوحي.</div>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-purple-400 flex flex-col justify-between">
            <div>
              <div class="w-10 h-10 rounded-lg bg-purple-500/20 text-purple-400 flex items-center justify-center text-xl mb-2">🚗</div>
              <h3 class="text-xs font-bold text-white mb-1">الأنظمة المدمجة</h3>
              <p class="text-[11px] text-purple-300 font-semibold mb-2">Embedded Systems</p>
              <p class="text-xs text-slate-300 leading-relaxed mb-2">حواسيب مصغرة مبرمجة داخل أجهزة أخرى لمهمة واحدة محددة بكفاءة مطلقة.</p>
            </div>
            <div class="text-[11px] bg-slate-900/80 text-purple-300 p-2 rounded-lg border border-purple-500/30">أمثلة: فرامل السيارات (ABS)، الغسالات الذكية، وإنترنت الأشياء (تشكل 90% من معالجات العالم!).</div>
          </div>
        </div>
        <div class="glass-card p-2.5 rounded-xl bg-slate-900/60 text-xs text-slate-300 text-center">
          الأستاذ نهمار محمد: "الأنظمة المدمجة تمثل أكثر من 90% من مجموع المعالجات المصنعة في العالم!"
        </div>
      </div>
    </section>

    <!-- SLIDE 7 -->
    <section class="slide" data-index="6" data-tag="المحور الثاني: هندسة المعالجة" data-title="طريقة عمل الحاسوب ودورة معالجة البيانات (IPOS)">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">المحور الثاني: هندسة المعالجة</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">طريقة عمل الحاسوب: دورة معالجة البيانات (IPOS Cycle)</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "طريقة عمل أي حاسوب تخضع لقانون منطقي واحد ثابت يسمى دورة <strong>IPOS</strong>: <strong>الإدخال (Input)</strong> ➔ <strong>المعالجة (Processing)</strong> ➔ <strong>الإخراج (Output)</strong>، وتتوسطها مرحلة <strong>التخزين (Storage)</strong>."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-4 gap-3.5 my-auto">
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-cyan-400 text-center">
            <div class="w-10 h-10 mx-auto rounded-full bg-cyan-500/20 text-cyan-400 flex items-center justify-center text-lg mb-2">⌨️</div>
            <span class="text-[10px] font-bold text-cyan-400 block mb-1">المرحلة 1</span>
            <h3 class="text-xs font-bold text-white mb-1.5">1. الإدخال (Input)</h3>
            <p class="text-[11px] text-slate-300 leading-relaxed mb-2">تلقي الأوامر والبيانات عبر لوحة المفاتيح، الفأرة، الميكروفون، الماسح، أو الكاميرا.</p>
            <span class="text-[10px] bg-slate-900 text-cyan-300 px-2 py-0.5 rounded font-mono">بيانات أولية (Data)</span>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-amber-400 text-center">
            <div class="w-10 h-10 mx-auto rounded-full bg-amber-500/20 text-amber-400 flex items-center justify-center text-lg mb-2">🧠</div>
            <span class="text-[10px] font-bold text-amber-400 block mb-1">المرحلة 2</span>
            <h3 class="text-xs font-bold text-white mb-1.5">2. المعالجة (Processing)</h3>
            <p class="text-[11px] text-slate-300 leading-relaxed mb-2">وحدة المعالجة المركزية (CPU) تقوم بالحسابات والمنطق وفرز النتائج وتفسير التعليمات.</p>
            <span class="text-[10px] bg-slate-900 text-amber-300 px-2 py-0.5 rounded font-mono">القلب المفكر (CPU)</span>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-emerald-400 text-center">
            <div class="w-10 h-10 mx-auto rounded-full bg-emerald-500/20 text-emerald-400 flex items-center justify-center text-lg mb-2">💾</div>
            <span class="text-[10px] font-bold text-emerald-400 block mb-1">المرحلة 3</span>
            <h3 class="text-xs font-bold text-white mb-1.5">3. التخزين (Storage)</h3>
            <p class="text-[11px] text-slate-300 leading-relaxed mb-2">حفظ البيانات مؤقتاً في RAM لتسريع العمل، أو دائماً في القرص الصلب (SSD / HDD).</p>
            <span class="text-[10px] bg-slate-900 text-emerald-300 px-2 py-0.5 rounded font-mono">ذاكرة مؤقتة ودائمة</span>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-blue-400 text-center">
            <div class="w-10 h-10 mx-auto rounded-full bg-blue-500/20 text-blue-400 flex items-center justify-center text-lg mb-2">🖥️</div>
            <span class="text-[10px] font-bold text-blue-400 block mb-1">المرحلة 4</span>
            <h3 class="text-xs font-bold text-white mb-1.5">4. الإخراج (Output)</h3>
            <p class="text-[11px] text-slate-300 leading-relaxed mb-2">إظهار النتيجة للمستخدم: نص على الشاشة، ورقة مطبوعة، أو صوت من السماعات.</p>
            <span class="text-[10px] bg-slate-900 text-blue-300 px-2 py-0.5 rounded font-mono">معلومات جاهزة (Output)</span>
          </div>
        </div>
        <div class="glass-card p-3 rounded-xl bg-slate-900/60 text-xs text-slate-300 flex items-center justify-between">
          <span>مثال: تكتب <kbd class="px-1.5 py-0.5 bg-slate-800 rounded font-mono text-cyan-300">5 + 3 =</kbd> (إدخال) ➔ المعالج يحسب (معالجة) ➔ يحفظ في RAM (تخزين) ➔ يظهر <strong class="text-emerald-400">8</strong> على الشاشة (إخراج)!</span>
          <button onclick="simulateIPOSCycle()" class="px-3 py-1 rounded-lg bg-cyan-600 hover:bg-cyan-500 text-white font-bold text-[11px] transition">▶ تجربة المحاكاة</button>
        </div>
        <div id="ipos-notification" class="hidden glass-card p-2 rounded-xl text-center text-xs text-emerald-300 border border-emerald-500/40 bg-emerald-950/60">
          🚀 جاري تدفق البيانات: إدخال الكيبورد ➔ وحدة ALU بالمعالج ➔ حفظ في RAM ➔ عرض النتيجة على الشاشة فوراً!
        </div>
      </div>
    </section>

    <!-- SLIDE 8 -->
    <section class="slide" data-index="7" data-tag="المحور الثاني: هندسة المعالجة" data-title="تشريح المعالج المركزي (CPU) ومعمارية فون نيومان">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">المحور الثاني: هندسة المعالجة</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">تشريح وحدة المعالجة المركزية (CPU): ماذا يوجد في عقل الحاسوب؟</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "يسمى المعالج <strong>CPU</strong> بـ 'دماغ الحاسوب'، ويتكون من 3 أقسام رئيسية تتكامل بدقة وفق <strong>معمارية فون نيومان</strong>: وحدة التحكم، وحدة الحساب والمنطق، والمسجلات وذاكرة الكاش."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4 my-auto">
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-cyan-400">
            <div class="w-10 h-10 rounded-lg bg-cyan-500/20 text-cyan-400 flex items-center justify-center text-xl mb-2">👮‍♂️</div>
            <h3 class="text-xs font-bold text-white mb-1">وحدة التحكم (CU)</h3>
            <p class="text-[11px] text-cyan-300 font-semibold mb-2">Control Unit - شرطي المرور</p>
            <p class="text-xs text-slate-300 leading-relaxed mb-2">تجلب الأوامر من الذاكرة، تفك تشفيرها، وترسل إشارات التزامن إلى كافة القطع لتعمل في تناغم دون تصادم.</p>
            <div class="text-[11px] text-slate-400 pt-2 border-t border-slate-700/60">الوظيفة: جلب التعليمات وفك التشفير والتنسيق.</div>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-emerald-400">
            <div class="w-10 h-10 rounded-lg bg-emerald-500/20 text-emerald-400 flex items-center justify-center text-xl mb-2">🔢</div>
            <h3 class="text-xs font-bold text-white mb-1">وحدة الحساب والمنطق (ALU)</h3>
            <p class="text-[11px] text-emerald-300 font-semibold mb-2">Arithmetic Logic Unit - الآلة الحاسبة</p>
            <p class="text-xs text-slate-300 leading-relaxed mb-2">تنفذ العمليات الرياضية (+, -, ×, ÷) والمنطقية (&gt;, &lt;, ==, AND, OR, NOT) بمليارات العمليات في الثانية.</p>
            <div class="text-[11px] text-slate-400 pt-2 border-t border-slate-700/60">الوظيفة: الحساب الرياضي والمقارنات المنطقية.</div>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-amber-400">
            <div class="w-10 h-10 rounded-lg bg-amber-500/20 text-amber-400 flex items-center justify-center text-xl mb-2">⚡</div>
            <h3 class="text-xs font-bold text-white mb-1">المسجلات وذاكرة الكاش</h3>
            <p class="text-[11px] text-amber-300 font-semibold mb-2">Registers & Cache - الملاحظات الفورية</p>
            <p class="text-xs text-slate-300 leading-relaxed mb-2">أسرع ذواكر داخل قلب المعالج (L1, L2, L3) لتغذية الحسابات فوراً وتفادي انتظار بطء الذاكرة العشوائية RAM.</p>
            <div class="text-[11px] text-slate-400 pt-2 border-t border-slate-700/60">الوظيفة: تغذية المعالج بالبيانات الفورية.</div>
          </div>
        </div>
        <div class="glass-card p-3 rounded-xl bg-slate-900/60 text-xs text-slate-300 flex items-center justify-between">
          <span>🔄 <strong>دورة الآلة الأكاديمية:</strong></span>
          <span class="text-cyan-400 font-mono text-[11px]">1. الجلب (Fetch) ➔ 2. فك التشفير (Decode) ➔ 3. التنفيذ (Execute) ➔ 4. الحفظ (Store)</span>
        </div>
      </div>
    </section>

    <!-- SLIDE 9 -->
    <section class="slide" data-index="8" data-tag="المحور الثاني: هندسة المعالجة" data-title="سر لغة الآلة: النظام الثنائي ووحدات قياس الذاكرة">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">المحور الثاني: هندسة المعالجة</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">سر النظام الثنائي (Binary): لماذا يتحدث الحاسوب بالأصفار والآحاد فقط؟</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "الحاسوب جهاز كهربائي، والكهرباء فيزيائياً تملك حالتين فقط: إما <strong>تيار يمر (On = 1)</strong>، أو <strong>لا تيار (Off = 0)</strong>! كل نص وصورة وفيديو يترجم في النهاية إلى نبضات كهربائية من 0 و 1."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4 my-auto">
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-cyan-400">
            <h3 class="text-sm font-bold text-white mb-2 flex items-center gap-2"><span>📏</span> سلم وحدات قياس الذاكرة الرقمية</h3>
            <div class="space-y-2 text-xs">
              <div class="p-2 rounded-xl bg-slate-900/80 border border-slate-700/60 flex items-center justify-between">
                <div><strong class="text-cyan-400">1 بت (Bit):</strong><span class="text-slate-300 mr-2">أصغر وحدة؛ خانة واحدة إما (0) أو (1).</span></div>
                <span class="text-[10px] text-slate-400 font-mono">0 or 1</span>
              </div>
              <div class="p-2 rounded-xl bg-slate-900/80 border border-slate-700/60 flex items-center justify-between">
                <div><strong class="text-emerald-400">1 بايت (Byte):</strong><span class="text-slate-300 mr-2">حزمة من <strong>8 بتات</strong> تكفي لتخزين <strong>حرف واحد فقط</strong>.</span></div>
                <span class="text-[10px] text-emerald-400 font-mono">8 Bits</span>
              </div>
              <div class="p-2 rounded-xl bg-slate-900/80 border border-slate-700/60 flex items-center justify-between">
                <div><strong class="text-amber-400">1 كيلوبايت (KB):</strong><span class="text-slate-300 mr-2">1024 بايت (صفحة نصية بسيطة).</span></div>
                <span class="text-[10px] text-slate-400 font-mono">1024 B</span>
              </div>
              <div class="p-2 rounded-xl bg-slate-900/80 border border-slate-700/60 flex items-center justify-between">
                <div><strong class="text-sky-400">1 ميغابايت (MB):</strong><span class="text-slate-300 mr-2">1024 كيلوبايت (كتاب كامل أو أغنية).</span></div>
                <span class="text-[10px] text-slate-400 font-mono">1024 KB</span>
              </div>
              <div class="p-2 rounded-xl bg-slate-900/80 border border-slate-700/60 flex items-center justify-between">
                <div><strong class="text-purple-400">1 غيغابايت (GB):</strong><span class="text-slate-300 mr-2">1024 ميغابايت (فيلم عالي الدقة).</span></div>
                <span class="text-[10px] text-slate-400 font-mono">1024 MB</span>
              </div>
              <div class="p-2 rounded-xl bg-slate-900/80 border border-slate-700/60 flex items-center justify-between">
                <div><strong class="text-pink-400">1 تيرابايت (TB):</strong><span class="text-slate-300 mr-2">1024 غيغابايت (سعة قرص صلب كامل).</span></div>
                <span class="text-[10px] text-slate-400 font-mono">1024 GB</span>
              </div>
            </div>
          </div>
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-purple-400 flex flex-col justify-between">
            <div>
              <div class="teacher-badge">🧪 تجربة تفاعلية أمام الطلبة:</div>
              <h3 class="text-sm font-bold text-white mb-2">كيف يرى الحاسوب الحرف الذي تكتبه؟</h3>
              <p class="text-xs text-slate-300 leading-relaxed mb-3">طبقاً لمعيار ASCII، اكتب أي حرف بالإنجليزية لترى كيف يتحول إلى 8 خانات ثنائية أمام أعين الطلبة:</p>
              <div class="p-3.5 rounded-xl bg-slate-950/80 border border-purple-500/40 mb-3">
                <div class="flex items-center gap-3 mb-2">
                  <input type="text" id="binary-input" maxlength="1" value="A" oninput="updateBinaryPreview()" class="w-16 text-center text-xl font-bold bg-slate-800 text-cyan-300 rounded-lg border border-cyan-500/50 py-1.5 focus:outline-none focus:border-cyan-400">
                  <span class="text-xs text-slate-300 font-semibold">⬅ الحرف المدخل (جرّب تغييره)</span>
                </div>
                <div class="p-2.5 rounded-lg bg-slate-900 border border-slate-700/80 flex items-center justify-between">
                  <span class="text-xs text-slate-400">التمثيل الثنائي (1 Byte):</span>
                  <span id="binary-output" class="text-sm font-mono font-bold text-emerald-400 tracking-widest">01000001</span>
                </div>
              </div>
            </div>
            <div class="p-2.5 rounded-xl bg-slate-900/80 text-[11px] text-purple-300 border border-purple-500/30">
              💡 لماذا 1024 وليس 1000؟ لأن النظام الثنائي يبنى على قوى العدد 2: (2^10 = 1024)!
            </div>
          </div>
        </div>
        <div class="glass-card p-2.5 rounded-xl bg-slate-900/60 text-xs text-slate-300 text-center">
          الأستاذ نهمار محمد: "1 Byte = 8 Bits. فلاش ديسك 32GB يخزن أكثر من 34 مليار حرف!"
        </div>
      </div>
    </section>

    <!-- SLIDE 10 -->
    <section class="slide" data-index="9" data-tag="المحور الثالث: أنظمة التشغيل" data-title="نظام التشغيل (OS) وتاريخ مايكروسوفت ويندوز">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">المحور الثالث: أنظمة التشغيل</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">نظام التشغيل (Operating System): المدير العام للحاسوب وتاريخ Windows</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "نظام التشغيل هو المدير العام والوسيط الشامل الذي يدير كل قطع العتاد ويوفر واجهة رسومية جميلة تسمح للمستخدم بتشغيل البرامج دون الحاجة للتخاطب المباشر مع الأسلاك والدارات."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-6 gap-3 my-auto">
          <div class="glass-card p-3 rounded-xl border-t-2 border-t-slate-500 flex flex-col justify-between">
            <div>
              <span class="text-[10px] font-bold text-slate-400">1985 - 1992</span>
              <h3 class="text-xs font-bold text-white mt-1 mb-1">Win 1.0 & 3.1</h3>
              <p class="text-[11px] text-slate-300 leading-relaxed">واجهة رسومية بسيطة جداً كانت تعمل فوق نظام الأوامر النصية MS-DOS.</p>
            </div>
            <span class="text-[10px] text-slate-400 font-semibold mt-2">البداية الأولى</span>
          </div>
          <div class="glass-card p-3 rounded-xl border-t-2 border-t-amber-500 flex flex-col justify-between">
            <div>
              <span class="text-[10px] font-bold text-amber-400">1995</span>
              <h3 class="text-xs font-bold text-white mt-1 mb-1">Windows 95</h3>
              <p class="text-[11px] text-slate-300 leading-relaxed">الانفجار الثوري الأكبر: ظهور زر ابدأ (Start)، شريط المهام، و Plug & Play.</p>
            </div>
            <span class="text-[10px] text-amber-400 font-semibold mt-2">الثورة الكبرى</span>
          </div>
          <div class="glass-card p-3 rounded-xl border-t-2 border-t-blue-500 flex flex-col justify-between">
            <div>
              <span class="text-[10px] font-bold text-blue-400">2001</span>
              <h3 class="text-xs font-bold text-white mt-1 mb-1">Windows XP</h3>
              <p class="text-[11px] text-slate-300 leading-relaxed">الأيقونة الخالدة: استقرار هائل بنواة NT والخلفية العشبية الخضراء الشهيرة.</p>
            </div>
            <span class="text-[10px] text-blue-400 font-semibold mt-2">عصر الاستقرار</span>
          </div>
          <div class="glass-card p-3 rounded-xl border-t-2 border-t-cyan-400 flex flex-col justify-between">
            <div>
              <span class="text-[10px] font-bold text-cyan-400">2009</span>
              <h3 class="text-xs font-bold text-white mt-1 mb-1">Windows 7</h3>
              <p class="text-[11px] text-slate-300 leading-relaxed">القمة الجمالية: واجهة زجاجية Aero Glass، استقرار وسرعة فائقة، والأكثر شعبية.</p>
            </div>
            <span class="text-[10px] text-cyan-400 font-semibold mt-2">الأكثر استقراراً</span>
          </div>
          <div class="glass-card p-3 rounded-xl border-t-2 border-t-indigo-500 flex flex-col justify-between">
            <div>
              <span class="text-[10px] font-bold text-indigo-400">2015</span>
              <h3 class="text-xs font-bold text-white mt-1 mb-1">Windows 10</h3>
              <p class="text-[11px] text-slate-300 leading-relaxed">التوحيد السحابي: متجر التطبيقات، كورتانا، واعتماد نموذج ويندوز كخدمة.</p>
            </div>
            <span class="text-[10px] text-indigo-400 font-semibold mt-2">نظام كخدمة</span>
          </div>
          <div class="glass-card p-3 rounded-xl border-t-2 border-t-emerald-400 flex flex-col justify-between">
            <div>
              <span class="text-[10px] font-bold text-emerald-400">2021 - الآن</span>
              <h3 class="text-xs font-bold text-white mt-1 mb-1">Windows 11</h3>
              <p class="text-[11px] text-slate-300 leading-relaxed">العصر الحديث: تصميم Fluent زجاجي، توسط شريط المهام، ودمج Copilot AI.</p>
            </div>
            <span class="text-[10px] text-emerald-400 font-semibold mt-2">الذكاء الاصطناعي</span>
          </div>
        </div>
        <div class="glass-card p-3 rounded-xl bg-slate-900/60 text-xs text-slate-300 flex items-center justify-between">
          <span>حصة ويندوز تفوق <strong>72%</strong> عالمياً في الحواسيب الشخصية.</span>
          <span class="text-cyan-400 font-bold">إتقان Windows مهارة وظيفية أساسية!</span>
        </div>
      </div>
    </section>

    <!-- SLIDE 11 -->
    <section class="slide" data-index="10" data-tag="المحور الرابع: بدء تشغيل النظام" data-title="رحلة بدء التشغيل: ماذا يحدث من زر الطاقة إلى شاشة الدخول؟">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">المحور الرابع: بدء تشغيل النظام</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">بدء تشغيل النظام (Boot Process): الرحلة الحرجة في الثواني الأولى</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "من لحظة الضغط على زر الطاقة (Power Button)، يخوض الحاسوب في ثوانٍ معدودة معركة تقنية دقيقة للتأكد من سلامة كل قطعة عتادية قبل أن يسلم القيادة لنظام التشغيل ويندوز."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-4 gap-3.5 my-auto">
          <div class="glass-card p-4 rounded-xl border-r-4 border-r-red-400">
            <span class="text-[10px] font-bold text-red-400 block mb-1">المرحلة الأولى</span>
            <h3 class="text-xs font-bold text-white mb-1.5">1. الطاقة وفحص POST</h3>
            <p class="text-[11px] text-slate-300 leading-relaxed mb-2">تصل الكهرباء فتبدأ اللوحة فحص العتاد الذاتي (Power-On Self-Test) لفحص المعالج والذاكرة وكارت الشاشة.</p>
            <div class="text-[10px] bg-red-950/40 text-red-300 p-1.5 rounded border border-red-500/30">⚠️ أصوات البيب (Beep Codes) تصدر عند فشل POST!</div>
          </div>
          <div class="glass-card p-4 rounded-xl border-r-4 border-r-amber-400">
            <span class="text-[10px] font-bold text-amber-400 block mb-1">المرحلة الثانية</span>
            <h3 class="text-xs font-bold text-white mb-1.5">2. BIOS / UEFI</h3>
            <p class="text-[11px] text-slate-300 leading-relaxed mb-2">برنامج اللوحة الأم. UEFI الحديث تفوق على BIOS بدعم الفأرة والواجهة الرسومية وأقراص عملاقة GPT وميزة الإقلاع الآمن Secure Boot.</p>
            <div class="text-[10px] bg-amber-950/40 text-amber-300 p-1.5 rounded border border-amber-500/30">🔑 مفاتيح الدخول: F2 أو Del أو F12</div>
          </div>
          <div class="glass-card p-4 rounded-xl border-r-4 border-r-cyan-400">
            <span class="text-[10px] font-bold text-cyan-400 block mb-1">المرحلة الثالثة</span>
            <h3 class="text-xs font-bold text-white mb-1.5">3. محمّل الإقلاع ونواة ويندوز</h3>
            <p class="text-[11px] text-slate-300 leading-relaxed mb-2">يسلم UEFI القيادة لـ Windows Boot Manager الذي يقرأ نواة ويندوز (ntoskrnl.exe) ويحملها في ذاكرة RAM مع تعريفات العتاد.</p>
            <div class="text-[10px] bg-cyan-950/40 text-cyan-300 p-1.5 rounded border border-cyan-500/30">⚡ شاشة شعار ويندوز الدوارة.</div>
          </div>
          <div class="glass-card p-4 rounded-xl border-r-4 border-r-emerald-400">
            <span class="text-[10px] font-bold text-emerald-400 block mb-1">المرحلة الرابعة</span>
            <h3 class="text-xs font-bold text-white mb-1.5">4. تسجيل الدخول وسطح المكتب</h3>
            <p class="text-[11px] text-slate-300 leading-relaxed mb-2">تنشط خدمة Winlogon وشاشة القفل، والتحقق عبر Windows Hello، ثم تحميل سطح المكتب ومستكشف الملفات.</p>
            <div class="text-[10px] bg-emerald-950/40 text-emerald-300 p-1.5 rounded border border-emerald-500/30">🎉 النظام جاهز للعمل والإنتاج!</div>
          </div>
        </div>
        <div class="glass-card p-3 rounded-xl bg-slate-900/60 text-xs text-slate-300 flex items-center justify-between">
          <span>الـ UEFI يحمي الحاسوب من الفيروسات التي تحاول الإقلاع قبل ويندوز عبر تقنية Secure Boot!</span>
        </div>
      </div>
    </section>

    <!-- SLIDE 12 -->
    <section class="slide" data-index="11" data-tag="أعمال تطبيقية (1)" data-title="ورشة تطبيقية: الدخول للـ BIOS والوضع الآمن (Safe Mode)">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">أعمال تطبيقية (1): مهارات الإقلاع والإنقاذ</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">ورشة تطبيقية (1): مفاتيح الدخول للـ BIOS وحل مشاكل الإقلاع بالوضع الآمن</h2>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4 my-auto">
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-amber-400">
            <div class="teacher-badge">👨‍🏫 النموذج التوجيهي للأستاذ نهمار محمد (طبقه أمام الطلبة):</div>
            <h3 class="text-sm font-bold text-white mb-2">خطوات دخول شاشة إعدادات BIOS / UEFI</h3>
            <ol class="text-xs text-slate-300 space-y-2 leading-relaxed">
              <li><strong class="text-amber-400">1. إعادة تشغيل الحاسوب:</strong> بمجرد أن تنطفئ الشاشة وتضيء من جديد، اضغط بشكل متقطع ومتكرر على مفتاح الدخول (F2 أو Del أو F12).</li>
              <li><strong class="text-cyan-400">2. استعراض قائمة الإقلاع (Boot Order):</strong> أظهر للطلبة كيف يختار الحاسوب الإقلاع من القرص الصلب أولاً أو من قرص USB Flash.</li>
              <li><strong class="text-emerald-400">3. الخروج الآمن:</strong> اضغط F10 للحفظ وإعادة التشغيل (Save and Exit).</li>
            </ol>
            <div class="mt-3 p-2.5 rounded-xl bg-amber-950/40 border border-amber-500/30 text-[11px] text-amber-300">
              ⚠️ تحذير: إياكم وتغيير خيارات الجهد الكهربائي أو كلمة مرور BIOS دون إشراف!
            </div>
          </div>
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-emerald-400 flex flex-col justify-between">
            <div>
              <div class="student-badge">👨‍💻 التمرين التطبيقي للطلبة في المخبر:</div>
              <h3 class="text-sm font-bold text-white mb-2">تمرين: تشغيل الوضع الآمن (Safe Mode) لإصلاح الأعطال</h3>
              <p class="text-xs text-slate-300 leading-relaxed mb-3">كيف ندخل للوضع الآمن المخصص لإصلاح الشاشة الزرقاء وتعطيل البرامج الخبيثة؟</p>
              <div class="p-3 rounded-xl bg-slate-900/80 border border-slate-700/60 text-xs text-slate-200 space-y-1.5">
                <div><strong>1.</strong> افتح قائمة ابدأ واضغط زر الطاقة ⏻.</div>
                <div><strong>2.</strong> <strong>اضغط باستمرار على زر Shift في الكيبورد</strong>، وأثناء ذلك انقر على <strong>Restart</strong> بالفأرة!</div>
                <div><strong>3.</strong> من القائمة الزرقاء: Troubleshoot ➔ Advanced options ➔ Startup Settings ➔ Safe Mode.</div>
              </div>
            </div>
            <div class="p-2.5 rounded-xl bg-emerald-950/40 border border-emerald-500/30 text-[11px] text-emerald-300 mt-3">
              🎯 <strong>النتيجة:</strong> الوضع الآمن يعطل البرامج الضارة ويشغل النظام بأبسط إمكانياته لإنقاذ الملفات.
            </div>
          </div>
        </div>
        <div class="glass-card p-2.5 rounded-xl bg-slate-900/60 text-xs text-slate-300 text-center">
          الأستاذ نهمار محمد: "الوضع الآمن هو الملاذ الأخير لإنقاذ ملفات الحاسوب قبل اللجوء للفورمات!"
        </div>
      </div>
    </section>

    <!-- SLIDE 13 (NEW): Interactive Windows 11 Desktop Simulator -->
    <section class="slide" data-index="12" data-tag="الواجهة الرسومية: المحاكي التفاعلي" data-title="محاكي سطح مكتب Windows 11 وتوزيع الأيقونات">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">الواجهة الرسومية: المحاكي التفاعلي</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-1">المحاكي التفاعلي لسطح المكتب Windows 11 وتوزيع الأيقونات الأساسية</h2>
          <p class="text-xs text-slate-300 mb-2">انقر على الأيقونات داخل الشاشة التفاعلية لتسمع صوت التفاعل وترى الشرح الأكاديمي المباشر لكل عنصر:</p>
        </div>

        <!-- Realistic Desktop Simulator Box -->
        <div class="glass-card rounded-2xl overflow-hidden border border-slate-700/80 my-auto shadow-2xl">
          <div class="relative h-64 md:h-72 bg-gradient-to-tr from-slate-900 via-sky-950 to-indigo-950 p-4 flex flex-col justify-between overflow-hidden">
            <!-- Simulated Icons Grid -->
            <div class="flex flex-col gap-2.5 w-fit z-10">
              <div onclick="clickSimIcon('💻 هذا الكمبيوتر الشخصي (This PC): بوابة النفاذ للأقراص الصلبة C: و D: ومحركات الأقراص وفحص السعات التخزينية المتبقية.')" class="flex items-center gap-2 p-1.5 rounded-lg hover:bg-white/10 cursor-pointer transition w-40 group">
                <span class="text-2xl group-hover:scale-110 transition transform">💻</span>
                <div>
                  <div class="text-xs font-bold text-white">هذا الكمبيوتر</div>
                  <div class="text-[9px] text-cyan-300">أيقونة نظام أصلية</div>
                </div>
              </div>

              <div onclick="clickSimIcon('🗑️ سلة المحذوفات (Recycle Bin): مستودع وقائي للملفات المحذوفة مؤقتاً لحمايتها من الخطأ البشري؛ تتيح استعادتها بنقرة Restore.')" class="flex items-center gap-2 p-1.5 rounded-lg hover:bg-white/10 cursor-pointer transition w-40 group">
                <span class="text-2xl group-hover:scale-110 transition transform">🗑️</span>
                <div>
                  <div class="text-xs font-bold text-white">سلة المحذوفات</div>
                  <div class="text-[9px] text-red-300">مستودع وقائي</div>
                </div>
              </div>

              <div onclick="clickSimIcon('👤 مجلد ملفات المستخدم: يضم المستندات، التنزيلات (Downloads)، الصور، وسطح المكتب نفسه الخاص بحساب الطالب.')" class="flex items-center gap-2 p-1.5 rounded-lg hover:bg-white/10 cursor-pointer transition w-40 group">
                <span class="text-2xl group-hover:scale-110 transition transform">👤</span>
                <div>
                  <div class="text-xs font-bold text-white">ملفات الطالب</div>
                  <div class="text-[9px] text-amber-300">المجلد الشخصي</div>
                </div>
              </div>

              <div onclick="clickSimIcon('⚙️ لوحة التحكم (Control Panel): المركز الإداري الكلاسيكي لإعدادات العتاد وإلغاء تثبيت البرامج وإدارة المستخدمين.')" class="flex items-center gap-2 p-1.5 rounded-lg hover:bg-white/10 cursor-pointer transition w-40 group">
                <span class="text-2xl group-hover:scale-110 transition transform">⚙️</span>
                <div>
                  <div class="text-xs font-bold text-white">لوحة التحكم</div>
                  <div class="text-[9px] text-emerald-300">مركز الإدارة</div>
                </div>
              </div>

              <div onclick="clickSimIcon('🌐 شبكة الاتصال (Network): استعراض الأجهزة والخوادم والمجلدات المشتركة داخل الشبكة المحلية (LAN).')" class="flex items-center gap-2 p-1.5 rounded-lg hover:bg-white/10 cursor-pointer transition w-40 group">
                <span class="text-2xl group-hover:scale-110 transition transform">🌐</span>
                <div>
                  <div class="text-xs font-bold text-white">شبكة الاتصال</div>
                  <div class="text-[9px] text-sky-300">أجهزة ومجلدات LAN</div>
                </div>
              </div>
            </div>

            <!-- Dynamic Information Banner -->
            <div id="sim-info-banner" class="absolute top-4 left-4 max-w-sm glass-card p-3 rounded-xl border border-cyan-500/40 text-xs text-slate-200 z-20 shadow-lg">
              👆 <strong>انقر على أي أيقونة في الشاشة</strong> لترى الشرح الأكاديمي المباشر وتسمع صوت التفاعل!
            </div>

            <!-- Simulated Taskbar -->
            <div class="h-11 -mx-4 -mb-4 bg-slate-950/90 backdrop-blur-md border-t border-slate-700/60 flex items-center justify-between px-4 z-20">
              <div class="text-[11px] text-slate-400 hidden sm:flex items-center gap-1.5">
                <span>⛅ 22°C</span> <span>قاعة الإعلام الآلي</span>
              </div>
              <div class="flex items-center gap-2">
                <button onclick="toggleSimStartMenu()" class="w-8 h-8 rounded-lg bg-cyan-600/30 hover:bg-cyan-500/50 text-cyan-300 flex items-center justify-center transition border border-cyan-500/40" title="قائمة ابدأ">🪟</button>
                <div onclick="clickSimIcon('🔍 البحث السريع (Win + S): النفاذ للبرامج والملفات والإعدادات فوراً.')" class="w-8 h-8 rounded-lg hover:bg-white/10 flex items-center justify-center text-slate-300 cursor-pointer">🔍</div>
                <div onclick="clickSimIcon('📁 مستكشف الملفات (Win + E): المتصفح الشجري للأقراص والملفات.')" class="w-8 h-8 rounded-lg hover:bg-white/10 flex items-center justify-center text-amber-400 cursor-pointer">📁</div>
                <div onclick="clickSimIcon('⚙️ الإعدادات (Win + I): نافذة التحكم العصرية بخصائص النظام واللغات.')" class="w-8 h-8 rounded-lg hover:bg-white/10 flex items-center justify-center text-cyan-400 cursor-pointer">⚙️</div>
              </div>
              <div onclick="clickSimIcon('🔔 صينية النظام (System Tray): تعرض الساعة، مستوى الصوت، اتصال الإنترنت، ولغة الإدخال AR.')" class="flex items-center gap-2 text-xs text-slate-300 hover:bg-white/10 p-1.5 rounded-lg cursor-pointer transition">
                <span class="text-[11px] font-bold text-cyan-400">AR</span>
                <span>🔊</span> <span>📶</span>
                <span class="font-mono text-[11px] text-slate-200">15:30</span>
              </div>
            </div>
          </div>
        </div>

        <div class="glass-card p-2.5 rounded-xl bg-slate-900/60 text-xs text-slate-300 text-center">
          الأستاذ نهمار محمد: "هذه الأيقونات الخمس تمثل البنية التحتية لإدارة نظام ويندوز، وحفظ وظيفة كل أيقونة يسهل عملكم اليومي!"
        </div>
      </div>
    </section>

    <!-- SLIDE 14 (NEW): Academic Icon Breakdown -->
    <section class="slide" data-index="13" data-tag="تشريح الأيقونات الأساسية الخمسة" data-title="الشرح الأكاديمي المفصل لأيقونات سطح المكتب الأساسية">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">تشريح الأيقونات الأساسية الخمسة</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">الشرح الأكاديمي المعمق للأيقونات الأساسية الخمسة على سطح المكتب</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "طلبتي الكرام، تلاحظون أن أيقونات النظام هذه لا تحمل سهماً أزرق صغيراً في زاويتها مثل الاختصارات العادية. لماذا؟ لأنها ليست روابط لملفات، بل هي كائنات نظام أصلية (Special System Objects). إليكم الدور الهندسي والوظيفي لكل واحدة منها بالتفصيل."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-3.5 my-auto text-xs">
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-cyan-400">
            <h3 class="font-bold text-white mb-1.5 flex items-center gap-1.5"><span class="text-lg">💻</span> 1. هذا الكمبيوتر الشخصي (This PC)</h3>
            <p class="text-slate-300 leading-relaxed mb-2">بوابة النفاذ الرئيسية لكافة وحدات التخزين الفيزيائية المتصلة بالحاسوب:</p>
            <ul class="text-slate-400 space-y-1 text-[11px]">
              <li>• أقسام القرص الصلب (C: للنظام، D: لبيانات المستخدم).</li>
              <li>• فحص السعات التخزينية المتبقية وتفريغ المساحة.</li>
              <li>• النقر باليمين ➔ <strong>Manage</strong> للدخول لإدارة الأقراص والعتاد.</li>
            </ul>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-red-400">
            <h3 class="font-bold text-white mb-1.5 flex items-center gap-1.5"><span class="text-lg">🗑️</span> 2. سلة المحذوفات (Recycle Bin)</h3>
            <p class="text-slate-300 leading-relaxed mb-2">المستودع الوقائي للملفات والمجلدات المحذوفة مؤقتاً لحمايتها من الخطأ البشري:</p>
            <ul class="text-slate-400 space-y-1 text-[11px]">
              <li>• لا يتم تحرير مساحة القرص إلا بعد النقر على <strong>Empty Recycle Bin</strong>.</li>
              <li>• استعادة أي ملف لمكانه الأصلي بنقرة يمين واختيار <strong>Restore</strong>.</li>
              <li>• يمكن تحديد سعتها القصوى من نافذة الخصائص (Properties).</li>
            </ul>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-amber-400">
            <h3 class="font-bold text-white mb-1.5 flex items-center gap-1.5"><span class="text-lg">👤</span> 3. مجلد ملفات المستخدم</h3>
            <p class="text-slate-300 leading-relaxed mb-2">المجلد الحاضن لكافة بيانات الطالب الشخصية المحمية بأذونات الوصول:</p>
            <ul class="text-slate-400 space-y-1 text-[11px]">
              <li>• يضم: المستندات (Documents)، التنزيلات (Downloads)، وسطح المكتب.</li>
              <li>• مساره الحقيقي في القرص: <code class="code-font text-cyan-300">C:\Users\Username</code>.</li>
              <li>• نصيحة الأستاذ: احتفظوا بملفاتكم الهامة في القرص D: لحمايتها عند الفورمات!</li>
            </ul>
          </div>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-3.5">
          <div class="glass-card p-3 rounded-xl border-t-2 border-t-emerald-400 text-xs">
            <strong class="text-emerald-400 block mb-1">⚙️ 4. لوحة التحكم (Control Panel):</strong>
            <span class="text-slate-300 text-[11px]">المركز الإداري الكلاسيكي؛ يتيح إزالة وتثبيت البرامج (Programs and Features)، إدارة حسابات المستخدمين وكلمات المرور، وإعدادات الطابعات والشبكات.</span>
          </div>
          <div class="glass-card p-3 rounded-xl border-t-2 border-t-sky-400 text-xs">
            <strong class="text-sky-400 block mb-1">🌐 5. شبكة الاتصال (Network):</strong>
            <span class="text-slate-300 text-[11px]">تتيح استعراض كافة الحواسيب والخوادم والطابعات المشتركة في مخبر الإعلام الآلي وتبادل الملفات عبر الشبكة المحلية (LAN) دون الحاجة لفلاش ديسك.</span>
          </div>
        </div>
      </div>
    </section>

    <!-- SLIDE 15 (NEW): Taskbar & Power States -->
    <section class="slide" data-index="14" data-tag="شريط المهام وخيارات الطاقة" data-title="تشريح شريط المهام وخيارات الطاقة (Sleep vs Hibernate)">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">شريط المهام وخيارات الطاقة</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">تشريح شريط المهام والفرق الجوهري بين خيارات الطاقة الأربعة</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "عندما ننقر على زر الطاقة في قائمة ابدأ، تظهر لنا خيارات متعددة: إيقاف التشغيل، إعادة التشغيل، وضع السكون (Sleep)، ووضع الإسبات (Hibernate). كثير من الطلبة لا يعرف الفارق التقني الدقيق بينها؛ دعونا نفكك كيف يتعامل المعالج والذاكرة مع كل خيار."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-4 gap-3.5 my-auto text-xs">
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-red-400">
            <div class="w-9 h-9 rounded-lg bg-red-500/20 text-red-400 flex items-center justify-center text-lg mb-2">⏻</div>
            <h3 class="font-bold text-white mb-1">1. إيقاف التشغيل</h3>
            <p class="text-red-300 text-[10px] font-semibold mb-2">Shut Down</p>
            <p class="text-slate-300 leading-relaxed mb-2">إغلاق كافة البرامج والنوافذ وحفظ إعدادات النظام وفصل التيار الكهربائي تماماً عن اللوحة الأم والمعالج والقرص.</p>
            <div class="text-[10px] text-slate-400 pt-1 border-t border-slate-700/60">استهلاك الطاقة: 0 واط تماماً.</div>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-cyan-400">
            <div class="w-9 h-9 rounded-lg bg-cyan-500/20 text-cyan-400 flex items-center justify-center text-lg mb-2">🔄</div>
            <h3 class="font-bold text-white mb-1">2. إعادة التشغيل</h3>
            <p class="text-cyan-300 text-[10px] font-semibold mb-2">Restart</p>
            <p class="text-slate-300 leading-relaxed mb-2">تفريغ الذاكرة العشوائية RAM بالكامل من الشوائب والبرامج العالقة، وإعادة تحميل نواة ويندوز من القرص لحل مشاكل البطء والتجمد.</p>
            <div class="text-[10px] text-cyan-300 pt-1 border-t border-slate-700/60">ضرورية بعد تثبيت التحديثات.</div>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-amber-400">
            <div class="w-9 h-9 rounded-lg bg-amber-500/20 text-amber-400 flex items-center justify-center text-lg mb-2">🌙</div>
            <h3 class="font-bold text-white mb-1">3. وضع السكون</h3>
            <p class="text-amber-300 text-[10px] font-semibold mb-2">Sleep Mode</p>
            <p class="text-slate-300 leading-relaxed mb-2">حفظ جلسة العمل في ذاكرة RAM مع إبقاء تيار كهربائي ضئيل لتغذيتها. يستيقظ الجهاز فوراً في ثانية واحدة بلمس الفأرة!</p>
            <div class="text-[10px] text-amber-300 pt-1 border-t border-slate-700/60">⚠️ تنبيه: انقطاع الكهرباء يفقد البيانات!</div>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-emerald-400">
            <div class="w-9 h-9 rounded-lg bg-emerald-500/20 text-emerald-400 flex items-center justify-center text-lg mb-2">💤</div>
            <h3 class="font-bold text-white mb-1">4. وضع الإسبات</h3>
            <p class="text-emerald-300 text-[10px] font-semibold mb-2">Hibernate Mode</p>
            <p class="text-slate-300 leading-relaxed mb-2">نسخ كامل محتوى ذاكرة RAM إلى ملف مخفي بالقرص الصلب (<code class="code-font text-emerald-300">hiberfil.sys</code>) وفصل الكهرباء 100%! وعند التشغيل يعود كل شيء كما تركته دون استنزاف البطارية.</p>
            <div class="text-[10px] text-emerald-300 pt-1 border-t border-slate-700/60">مثالي للحواسيب المحمولة عند السفر.</div>
          </div>
        </div>
        <div class="glass-card p-2.5 rounded-xl bg-slate-900/60 text-xs text-slate-300 text-center">
          الأستاذ نهمار محمد: "في وضع السكون الذاكرة تحتاج كهرباء، بينما في وضع الإسبات الكهرباء مفصولة تماماً والملفات محفوظة بأمان على القرص!"
        </div>
      </div>
    </section>

    <!-- SLIDE 16 (Preserving existing Slide 13) -->
    <section class="slide" data-index="15" data-tag="المحور الخامس: تشريح الواجهة" data-title="الواجهة الرسومية: سطح المكتب وشريط المهام والأيقونات">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">المحور الخامس: تشريح الواجهة</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">الواجهة الرسومية (GUI): سطح المكتب، شريط المهام، وتصنيف الأيقونات</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "الأيقونات تنقسم لثلاثة أصناف رئيسية: أيقونات النظام الأصلية، أيقونات الاختصار التي تتميز بسهم أزرق صغير، وأيقونات الملفات والمجلدات. وتذكروا أن حذف الاختصار يزيل فقط الرابط ولا يمس البرنامج الأصلي."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4 my-auto">
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-cyan-400">
            <h3 class="text-xs font-bold text-white mb-1.5 flex items-center gap-1.5"><span>🖥️</span> 1. سطح المكتب (Desktop)</h3>
            <p class="text-xs text-slate-300 leading-relaxed mb-3">هو المساحة المرئية الواسعة خلف كل النوافذ المفتوحة. يمثل المجلد الأساسي لحسابك.</p>
            <ul class="text-[11px] text-slate-300 space-y-1.5">
              <li>• <strong>الخلفية:</strong> صورة أو لون لتخصيص بيئة العمل.</li>
              <li>• <strong>الترتيب التلقائي:</strong> يمين الفأرة ➔ عرض ➔ ترتيب الأيقونات تلقائياً حسب شبكة خفية.</li>
              <li>• <strong>إخفاء الأيقونات:</strong> يمين الفأرة ➔ عرض ➔ إلغاء "إظهار أيقونات سطح المكتب".</li>
            </ul>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-amber-400">
            <h3 class="text-xs font-bold text-white mb-1.5 flex items-center gap-1.5"><span>📌</span> 2. شريط المهام (Taskbar)</h3>
            <p class="text-xs text-slate-300 leading-relaxed mb-3">الشريط الأفقي أسفل الشاشة (في ويندوز 11 توسطت أيقوناته لأول مرة).</p>
            <ul class="text-[11px] text-slate-300 space-y-1.5">
              <li>• <strong>زر ابدأ (Start):</strong> بوابة الوصول لكافة البرامج وخيارات الطاقة.</li>
              <li>• <strong>البرامج النشطة:</strong> يظهر تحتها خط أزرق مضيء صغير.</li>
              <li>• <strong>صينية النظام (System Tray):</strong> الساعة، Wi-Fi، الصوت، واللغة.</li>
            </ul>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-emerald-400">
            <h3 class="text-xs font-bold text-white mb-1.5 flex items-center gap-1.5"><span>📁</span> 3. تصنيف الأيقونات (Icons)</h3>
            <p class="text-xs text-slate-300 leading-relaxed mb-3">الأيقونة رمز صوري صغير يمثل عنصراً في الحاسوب:</p>
            <ul class="text-[11px] text-slate-300 space-y-1.5">
              <li>• <strong>أيقونات النظام:</strong> "هذا الكمبيوتر الشخصي" و "سلة المحذوفات".</li>
              <li>• <strong>أيقونات الاختصار (Shortcuts):</strong> <strong>تتميز بسهم أزرق مقوس</strong>؛ حذفه يحذف الرابط فقط ويبقى البرنامج سليماً.</li>
              <li>• <strong>أيقونات الملفات والمجلدات:</strong> مستنداتك وصورك الحقيقية.</li>
            </ul>
          </div>
        </div>
        <div class="glass-card p-3 rounded-xl bg-slate-900/60 text-xs text-slate-300 flex items-center justify-between">
          <span>حركات الفأرة الأساسية: النقر المفرد (تحديد) • النقر المزدوج (فتح وتشغيل) • النقر الأيمن (القائمة السياقية Context Menu) • السحب والإفلات (نقل وتكرار).</span>
        </div>
      </div>
    </section>

    <!-- SLIDE 17 (Preserving existing Slide 14) -->
    <section class="slide" data-index="16" data-tag="أعمال تطبيقية (2)" data-title="ورشة تطبيقية: استعادة الأيقونات وخماسية الاختصارات">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">أعمال تطبيقية (2): احتراف سطح المكتب</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">ورشة تطبيقية (2): مهارات سطح المكتب المتقدمة وخماسية الاختصارات الذهبية</h2>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4 my-auto">
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-cyan-400">
            <div class="teacher-badge">👨‍🏫 النموذج التوجيهي للأستاذ نهمار محمد (طبقه أمام الطلبة):</div>
            <h3 class="text-sm font-bold text-white mb-2">استعادة أيقونة "هذا الكمبيوتر" المفقودة بالطريقة الصحيحة</h3>
            <div class="p-3 rounded-xl bg-slate-900/80 text-xs text-slate-200 border border-slate-700/60 space-y-1">
              <div>1. انقر بيمين الفأرة على سطح المكتب ➔ اختر <strong>تخصيص (Personalize)</strong>.</div>
              <div>2. ادخل إلى قسم <strong>النسق (Themes)</strong>.</div>
              <div>3. انزل واضغط على <strong>إعدادات أيقونات سطح المكتب (Desktop icon settings)</strong>.</div>
              <div>4. ضع علامة (✓) أمام "الكمبيوتر" وسلة المحذوفات ➔ اضغط تطبيق وموافق!</div>
            </div>
          </div>
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-emerald-400 flex flex-col justify-between">
            <div>
              <div class="student-badge">👨‍💻 المهمة الموجهة للطلبة على حواسيبهم في المخبر:</div>
              <h3 class="text-sm font-bold text-white mb-2">المهمة: عزل بيئة الدراسة عن الترفيه بأسطح المكتب الافتراضية</h3>
              <div class="p-3 rounded-xl bg-slate-900/80 text-xs text-slate-200 border border-slate-700/60 space-y-1">
                <div>• اضغط <kbd class="px-1.5 py-0.5 bg-slate-800 text-emerald-300 rounded font-mono">Win + Tab</kbd> لفتح واجهة المهام (Task View).</div>
                <div>• انقر على <strong>(+) New Desktop</strong> وسمّه "مشاريع الإعلام الآلي".</div>
                <div>• تنقل بين السطحين باستخدام: <kbd class="px-1.5 py-0.5 bg-slate-800 text-emerald-300 rounded font-mono">Ctrl + Win + سهم يمين / يسار</kbd>!</div>
              </div>
            </div>
            <div class="p-2.5 rounded-xl bg-emerald-950/40 border border-emerald-500/30 text-[11px] text-emerald-300 mt-3">
              🎯 <strong>النتيجة:</strong> لن تختلط نوافذك أبداً أثناء العمل والبحث الجامعي!
            </div>
          </div>
        </div>
        <div class="grid grid-cols-2 sm:grid-cols-5 gap-2 text-center text-xs">
          <div class="glass-card p-2 rounded-xl">
            <kbd class="px-1.5 py-0.5 bg-slate-800 text-cyan-300 rounded font-mono text-[10px]">Win + D</kbd>
            <span class="block text-[10px] text-slate-300 mt-1">إظهار / إخفاء سطح المكتب</span>
          </div>
          <div class="glass-card p-2 rounded-xl">
            <kbd class="px-1.5 py-0.5 bg-slate-800 text-cyan-300 rounded font-mono text-[10px]">Win + E</kbd>
            <span class="block text-[10px] text-slate-300 mt-1">فتح مستكشف الملفات</span>
          </div>
          <div class="glass-card p-2 rounded-xl">
            <kbd class="px-1.5 py-0.5 bg-slate-800 text-cyan-300 rounded font-mono text-[10px]">Win + L</kbd>
            <span class="block text-[10px] text-slate-300 mt-1">قفل الشاشة فوراً</span>
          </div>
          <div class="glass-card p-2 rounded-xl">
            <kbd class="px-1.5 py-0.5 bg-slate-800 text-cyan-300 rounded font-mono text-[10px]">Win + Shift + S</kbd>
            <span class="block text-[10px] text-slate-300 mt-1">لقطة شاشة محددة</span>
          </div>
          <div class="glass-card p-2 rounded-xl">
            <kbd class="px-1.5 py-0.5 bg-slate-800 text-cyan-300 rounded font-mono text-[10px]">Alt + Tab</kbd>
            <span class="block text-[10px] text-slate-300 mt-1">التبديل بين النوافذ</span>
          </div>
        </div>
      </div>
    </section>

    <!-- SLIDE 18 (Preserving existing Slide 15) -->
    <section class="slide" data-index="17" data-tag="المحور السادس: خصائص النظام" data-title="خصائص النظام: الوقت والتاريخ وبطارية CMOS واللغات">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">المحور السادس: خصائص النظام</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">خصائص النظام: ضبط التاريخ والوقت، سر بطارية CMOS، وإدارة اللغات</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "كيف يعرف الحاسوب الوقت والتاريخ بدقة حتى لو فصلته عن الكهرباء لشهور؟ الجواب يكمن في قطعة عتادية تسمى <strong>بطارية CMOS</strong>، وبروتوكول برمجي يسمى <strong>NTP</strong>. واليوم سنفهم أيضاً الفرق بين لغة العرض ولغة لوحة المفاتيح."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4 my-auto">
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-cyan-400">
            <div class="w-10 h-10 rounded-lg bg-cyan-500/20 text-cyan-400 flex items-center justify-center text-xl mb-2">⏰</div>
            <h3 class="text-xs font-bold text-white mb-1">المزامنة التلقائية (NTP)</h3>
            <p class="text-[11px] text-cyan-300 font-semibold mb-2">Network Time Protocol</p>
            <p class="text-xs text-slate-300 leading-relaxed mb-3">يتواصل ويندوز عبر بروتوكول NTP مع خوادم مايكروسوفت الذرية لمطابقة الثواني تلقائياً بحسب خط المنطقة الزمنية لبلدك.</p>
            <div class="text-[11px] text-slate-400 pt-2 border-t border-slate-700/60">اختصار الدخول السريع: <kbd class="px-1.5 py-0.5 bg-slate-800 rounded font-mono text-[10px]">Win + I</kbd></div>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-amber-400">
            <div class="w-10 h-10 rounded-lg bg-amber-500/20 text-amber-400 flex items-center justify-center text-xl mb-2">🔋</div>
            <h3 class="text-xs font-bold text-white mb-1">سر بطارية السيموس (CMOS)</h3>
            <p class="text-[11px] text-amber-300 font-semibold mb-2">بطارية CR2032 على اللوحة الأم</p>
            <p class="text-xs text-slate-300 leading-relaxed mb-3">بطارية دائرية صغيرة وظيفتها إبقاء شريحة الوقت وإعدادات BIOS حية دون انقطاع حتى عند انقطاع الكهرباء تماماً.</p>
            <div class="p-2 rounded-lg bg-amber-950/40 border border-amber-500/30 text-[10px] text-amber-300">⚠️ إذا رجع التاريخ لسنة 2000 فيجب تغييرها فوراً!</div>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-emerald-400">
            <div class="w-10 h-10 rounded-lg bg-emerald-500/20 text-emerald-400 flex items-center justify-center text-xl mb-2">🌐</div>
            <h3 class="text-xs font-bold text-white mb-1">إدارة اللغات والمناطق</h3>
            <p class="text-[11px] text-emerald-300 font-semibold mb-2">Display vs Input Language</p>
            <p class="text-xs text-slate-300 leading-relaxed mb-3">• <strong>لغة العرض (Display):</strong> لغة القوائم والنوافذ.<br>• <strong>لوحة الإدخال (Input):</strong> تخطيط المفاتيح (عربية، فرنسية، إنجليزية).</p>
            <div class="p-2 rounded-lg bg-emerald-950/40 border border-emerald-500/30 text-[10px] text-emerald-300">التبديل الفوري: <kbd class="px-1.5 py-0.5 bg-slate-800 rounded font-mono text-[10px]">Alt + Shift</kbd></div>
          </div>
        </div>
        <div class="glass-card p-3 rounded-xl bg-slate-900/60 text-xs text-slate-300 flex items-center justify-between">
          <span>حل خطأ 'Your connection is not private': "غالباً ما يكون السبب هو عدم مطابقة ساعة حاسوبك مع خوادم الويب!"</span>
        </div>
      </div>
    </section>

    <!-- SLIDE 19 (Preserving existing Slide 16) -->
    <section class="slide" data-index="18" data-tag="أعمال تطبيقية (3)" data-title="ورشة تطبيقية: فحص مواصفات الحاسوب العتادية وأوامر dxdiag">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">أعمال تطبيقية (3): استكشاف العتاد</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">ورشة تطبيقية (3): فحص مواصفات الجهاز والتمييز بين 32-bit و 64-bit</h2>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4 my-auto">
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-cyan-400">
            <div class="teacher-badge">👨‍🏫 النموذج التوجيهي للأستاذ نهمار محمد (طبقه أمام الطلبة):</div>
            <h3 class="text-sm font-bold text-white mb-2">ثلاث طرق لفحص عتاد الحاسوب دون برامج خارجية</h3>
            <div class="space-y-2 text-xs text-slate-300">
              <div class="p-2.5 rounded-xl bg-slate-900/80 border border-slate-700/60">
                <strong class="text-cyan-400 block mb-0.5">1. الإعدادات السريعة:</strong>
                <span><kbd class="px-1.5 py-0.5 bg-slate-800 text-cyan-300 rounded font-mono">Win + I</kbd> ➔ النظام ➔ حول (About).</span>
              </div>
              <div class="p-2.5 rounded-xl bg-slate-900/80 border border-slate-700/60">
                <strong class="text-emerald-400 block mb-0.5">2. أداة التشخيص dxdiag:</strong>
                <span><kbd class="px-1.5 py-0.5 bg-slate-800 text-emerald-300 rounded font-mono">Win + R</kbd> واكتب <code class="code-font text-emerald-300">dxdiag</code> لعرض كارت الشاشة والبيوس.</span>
              </div>
              <div class="p-2.5 rounded-xl bg-slate-900/80 border border-slate-700/60">
                <strong class="text-amber-400 block mb-0.5">3. تقرير معلومات النظام msinfo32:</strong>
                <span><kbd class="px-1.5 py-0.5 bg-slate-800 text-amber-300 rounded font-mono">Win + R</kbd> واكتب <code class="code-font text-amber-300">msinfo32</code> لتقرير هندسي شامل.</span>
              </div>
            </div>
          </div>
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-emerald-400 flex flex-col justify-between">
            <div>
              <div class="student-badge">👨‍💻 المهمة الموجهة للطلبة على حواسيبهم في المخبر:</div>
              <h3 class="text-sm font-bold text-white mb-2">المهمة: استخراج بطاقة المواصفات الفنية لحاسوبك في المخبر</h3>
              <div class="p-3 rounded-xl bg-slate-900/80 text-xs text-slate-200 border border-slate-700/60 space-y-2">
                <div><strong class="text-cyan-400">1. طراز وسرعة المعالج (Processor):</strong> (مثال: Intel Core i5 @ 3.20GHz)</div>
                <div><strong class="text-emerald-400">2. حجم الذاكرة العشوائية المثبتة (RAM):</strong> (مثال: 8.00 GB)</div>
                <div><strong class="text-amber-400">3. نوع النظام (System type):</strong> (هل هو 64-bit أو 32-bit؟)</div>
              </div>
            </div>
            <div class="p-2.5 rounded-xl bg-emerald-950/40 border border-emerald-500/30 text-[11px] text-emerald-300 mt-3">
              🔍 <strong>الفارق الأكاديمي الحاسم:</strong> أنظمة 32-bit لا تقرأ أكثر من 4GB رام، بينما 64-bit تدعم ذواكر عملاقة وتشغل البرامج الهندسية الثقيلة!
            </div>
          </div>
        </div>
        <div class="glass-card p-2.5 rounded-xl bg-slate-900/60 text-xs text-slate-300 text-center">
          الأستاذ نهمار محمد: "تأكدوا دائماً قبل تحميل أي برنامج جامعي من أن حاسوبكم 64-bit!"
        </div>
      </div>
    </section>

    <!-- SLIDE 20 (Preserving existing Slide 17) -->
    <section class="slide" data-index="19" data-tag="المحور السابع: إدارة البيانات" data-title="المفهوم الجذري للملفات والمجلدات وأهمية الامتداد">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">المحور السابع: إدارة البيانات</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">المفهوم الجذري للملفات والمجلدات (Files & Folders) وسر الامتدادات</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "كيف يعرف الحاسوب نوع الملف؟ السر يكمن في <strong>امتداد الملف (File Extension)</strong>. الملف والمجلد هما حجر الأساس لتنظيم البيانات على القرص الصلب."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4 my-auto">
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-cyan-400">
            <h3 class="text-sm font-bold text-white mb-2 flex items-center gap-2"><span>📄</span> تشريح اسم الملف (File Structure)</h3>
            <div class="p-3.5 rounded-xl bg-slate-950 border border-cyan-500/30 mb-3 flex items-center justify-center gap-1 font-mono text-base md:text-lg font-bold">
              <span class="px-3 py-1 rounded bg-cyan-950 text-cyan-400 border border-cyan-500/40">cours_informatique</span>
              <span class="text-white text-xl font-black">.</span>
              <span class="px-3 py-1 rounded bg-emerald-950 text-emerald-400 border border-emerald-500/40">pptx</span>
            </div>
            <div class="space-y-2 text-xs">
              <div class="p-2.5 rounded-xl bg-slate-900/80 border border-slate-700/60">
                <strong class="text-cyan-400 block mb-0.5">1. اسم الملف (File Name):</strong>
                <span class="text-slate-300">يختاره المستخدم ليدل على المحتوى (يمنع 9 محارف: <code class="code-font text-red-300">\ / : * ? " &lt; &gt; |</code>).</span>
              </div>
              <div class="p-2.5 rounded-xl bg-slate-900/80 border border-slate-700/60">
                <strong class="text-emerald-400 block mb-0.5">2. الامتداد (File Extension):</strong>
                <span class="text-slate-300">يحدد لنظام التشغيل نوع البيانات والبرنامج المسؤول عن فتحها.</span>
              </div>
            </div>
          </div>
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-amber-400 flex flex-col justify-between">
            <div>
              <h3 class="text-sm font-bold text-white mb-2 flex items-center gap-2"><span>📁</span> ما هو المجلد (Folder)؟</h3>
              <p class="text-xs text-slate-300 leading-relaxed mb-3">حاوية افتراضية لا تملك امتداداً، وظيفتها احتضان الملفات وتنظيمها في شجرة هرمية.</p>
              <div class="p-2.5 rounded-xl bg-slate-950 border border-amber-500/30 text-xs font-mono text-amber-300 mb-3 overflow-x-auto text-left" dir="ltr">
                C:\Universite\Informatique\Semestre1\Projet.docx
              </div>
              <div class="p-3 rounded-xl bg-amber-950/40 border border-amber-500/30 text-xs text-amber-200 leading-relaxed">
                🔒 <strong>نصيحة أمنية جوهرية:</strong> يقوم ويندوز افتراضياً بإخفاء امتدادات الملفات. هذا خطأ أمني خطير! لأن الفيروسات تتنكر بأسماء مثل <code class="code-font text-red-300">photo.jpg.exe</code>! يجب تفعيل إظهار الامتدادات دائماً!
              </div>
            </div>
            <div class="text-[11px] text-slate-400 pt-2 border-t border-slate-700/60">
              💡 حجم المجلد يساوي مجموع أحجام كافة الملفات بداخله.
            </div>
          </div>
        </div>
        <div class="glass-card p-2.5 rounded-xl bg-slate-900/60 text-xs text-slate-300 text-center">
          الأستاذ نهمار محمد: "لا تغيروا امتداد الملف عند إعادة تسميته إلا إذا كنتم متأكدين مما تفعلون!"
        </div>
      </div>
    </section>

    <!-- SLIDE 21 (Preserving existing Slide 18) -->
    <section class="slide" data-index="20" data-tag="المحور السابع: إدارة البيانات" data-title="جدول أشهر امتدادات الملفات وتصنيفاتها البرمجية">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">المحور السابع: إدارة البيانات</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">جدول أشهر امتدادات الملفات الرقمية في عالم الحوسبة</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "إليكم هذا الدليل المرجعي الذي يجب على كل طالب حفظه عن ظهر قلب؛ أشهر عائلات الامتدادات مصنفة بحسب نوع المحتوى والتطبيق المسؤول."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-4 gap-3.5 my-auto text-xs">
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-blue-400">
            <h3 class="text-xs font-bold text-white mb-2 flex items-center gap-1.5"><span>📝</span> المستندات والنصوص</h3>
            <div class="space-y-2">
              <div class="p-2 rounded-lg bg-slate-900/80 border border-slate-700/60"><span class="code-font text-blue-400 font-bold">.docx</span><span class="block text-[11px] text-slate-300">مستند نصي منسق (Word)</span></div>
              <div class="p-2 rounded-lg bg-slate-900/80 border border-slate-700/60"><span class="code-font text-red-400 font-bold">.pdf</span><span class="block text-[11px] text-slate-300">وثيقة محمولة ثابتة</span></div>
              <div class="p-2 rounded-lg bg-slate-900/80 border border-slate-700/60"><span class="code-font text-emerald-400 font-bold">.xlsx</span><span class="block text-[11px] text-slate-300">جداول حسابية (Excel)</span></div>
              <div class="p-2 rounded-lg bg-slate-900/80 border border-slate-700/60"><span class="code-font text-amber-400 font-bold">.pptx</span><span class="block text-[11px] text-slate-300">عروض تقديمية (PowerPoint)</span></div>
            </div>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-purple-400">
            <h3 class="text-xs font-bold text-white mb-2 flex items-center gap-1.5"><span>🎬</span> الصور والوسائط</h3>
            <div class="space-y-2">
              <div class="p-2 rounded-lg bg-slate-900/80 border border-slate-700/60"><span class="code-font text-purple-400 font-bold">.jpg / .jpeg</span><span class="block text-[11px] text-slate-300">صور فوتوغرافية مضغوطة</span></div>
              <div class="p-2 rounded-lg bg-slate-900/80 border border-slate-700/60"><span class="code-font text-pink-400 font-bold">.png</span><span class="block text-[11px] text-slate-300">صور عالية الدقة شفافة</span></div>
              <div class="p-2 rounded-lg bg-slate-900/80 border border-slate-700/60"><span class="code-font text-cyan-400 font-bold">.mp4 / .mkv</span><span class="block text-[11px] text-slate-300">فيديو رقمي مدمج</span></div>
              <div class="p-2 rounded-lg bg-slate-900/80 border border-slate-700/60"><span class="code-font text-yellow-400 font-bold">.mp3 / .wav</span><span class="block text-[11px] text-slate-300">تسجيلات صوتية</span></div>
            </div>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-amber-400">
            <h3 class="text-xs font-bold text-white mb-2 flex items-center gap-1.5"><span>📦</span> الأرشيف والضغط</h3>
            <div class="space-y-2">
              <div class="p-2 rounded-lg bg-slate-900/80 border border-slate-700/60"><span class="code-font text-amber-400 font-bold">.zip</span><span class="block text-[11px] text-slate-300">أرشيف مضغوط قياسي</span></div>
              <div class="p-2 rounded-lg bg-slate-900/80 border border-slate-700/60"><span class="code-font text-orange-400 font-bold">.rar</span><span class="block text-[11px] text-slate-300">أرشيف عالي الضغط</span></div>
              <div class="p-2 rounded-lg bg-slate-900/80 border border-slate-700/60"><span class="code-font text-red-300 font-bold">.iso</span><span class="block text-[11px] text-slate-300">صورة قرص افتراضي</span></div>
            </div>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-red-400">
            <h3 class="text-xs font-bold text-white mb-2 flex items-center gap-1.5"><span>⚙️</span> البرامج التنفيذية</h3>
            <div class="space-y-2">
              <div class="p-2 rounded-lg bg-slate-900/80 border border-slate-700/60"><span class="code-font text-red-400 font-bold">.exe</span><span class="block text-[11px] text-slate-300">برنامج تنفيذي مباشر</span></div>
              <div class="p-2 rounded-lg bg-slate-900/80 border border-slate-700/60"><span class="code-font text-cyan-300 font-bold">.msi</span><span class="block text-[11px] text-slate-300">حزمة تثبيت ويندوز</span></div>
              <div class="p-2 rounded-lg bg-slate-900/80 border border-slate-700/60"><span class="code-font text-yellow-300 font-bold">.bat</span><span class="block text-[11px] text-slate-300">ملف أوامر برمجية</span></div>
            </div>
          </div>
        </div>
        <div class="glass-card p-3 rounded-xl bg-slate-900/60 text-xs text-slate-300 flex items-center justify-between">
          <span>تحذير أمني: إياكم وفتح أي ملف مجهول يحمل امتداد <code class="code-font text-red-400">.exe</code> مرسل في البريد المشبوه!</span>
        </div>
      </div>
    </section>

    <!-- SLIDE 22 (Preserving existing Slide 19) -->
    <section class="slide" data-index="21" data-tag="المحور السابع: إدارة البيانات" data-title="العمليات الأساسية والاحترافية على الملفات والمجلدات">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">المحور السابع: إدارة البيانات</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">العمليات الأساسية والاحترافية على الملفات والمجلدات</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "هناك 4 فوارق جوهرية: الفرق بين النسخ والقص، الفرق بين الحذف المؤقت والحذف النهائي، ضغط الملفات، وفحص خصائص الملف وحجمه على القرص."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4 my-auto">
          <div class="glass-card p-4 rounded-xl border-r-4 border-r-cyan-400">
            <h3 class="text-xs font-bold text-white mb-2 flex items-center gap-1.5"><span>📋</span> 1. النسخ (Copy) مقابل القص (Cut)</h3>
            <div class="space-y-2 text-xs">
              <div class="p-2.5 rounded-lg bg-slate-900/80 border border-slate-700/60">
                <strong class="text-cyan-400 block mb-0.5">• النسخ (Ctrl + C):</strong>
                <span class="text-slate-300">أخذ نسخة مكررة مع بقاء الأصل مكانه دون مساس.</span>
              </div>
              <div class="p-2.5 rounded-lg bg-slate-900/80 border border-slate-700/60">
                <strong class="text-amber-400 block mb-0.5">• القص (Ctrl + X):</strong>
                <span class="text-slate-300">نقل الملف الفعلي لموضع جديد عند اللصق.</span>
              </div>
              <div class="p-2 rounded-lg bg-slate-900/80 border border-slate-700/60 text-slate-300">
                ⚡ اللصق دائماً عبر: <kbd class="px-1.5 py-0.5 bg-slate-800 text-emerald-400 rounded font-mono">Ctrl + V</kbd>.
              </div>
            </div>
          </div>
          <div class="glass-card p-4 rounded-xl border-r-4 border-r-red-400">
            <h3 class="text-xs font-bold text-white mb-2 flex items-center gap-1.5"><span>🗑️</span> 2. الحذف العادي مقابل الحذف النهائي</h3>
            <div class="space-y-2 text-xs">
              <div class="p-2.5 rounded-lg bg-slate-900/80 border border-slate-700/60">
                <strong class="text-amber-400 block mb-0.5">• الحذف العادي (Delete):</strong>
                <span class="text-slate-300">إجراء وقائي ينقل الملف لسلة المحذوفات ويمكن استرجاعه (Restore).</span>
              </div>
              <div class="p-2.5 rounded-lg bg-slate-900/80 border border-slate-700/60">
                <strong class="text-red-400 block mb-0.5">• الحذف النهائي (Shift + Delete):</strong>
                <span class="text-slate-300">حذف فوري يتجاوز سلة المحذوفات ولا يمكن استرجاعه بالطرق العادية!</span>
              </div>
              <div class="p-2 rounded-lg bg-red-950/40 text-[11px] text-red-300 border border-red-500/30">
                ⚠️ استخدم Shift+Delete فقط عندما تكون متأكداً 100% أنك لا تحتاج الملف!
              </div>
            </div>
          </div>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div class="glass-card p-3 rounded-xl text-xs flex items-center justify-between">
            <div>
              <strong class="text-amber-400 block">📦 الضغط والأرشفة (ZIP):</strong>
              <span class="text-slate-300">يمين الفأرة ➔ Compress to ZIP file؛ يقلص الحجم لتسهيل الإرسال بالبريد و Moodle.</span>
            </div>
          </div>
          <div class="glass-card p-3 rounded-xl text-xs flex items-center justify-between">
            <div>
              <strong class="text-purple-400 block">⚙️ نافذة الخصائص (Alt + Enter):</strong>
              <span class="text-slate-300">معرفة الحجم الفعلي وتفعيل سمة "للقراءة فقط (Read-only)" أو "مخفي (Hidden)".</span>
            </div>
          </div>
        </div>
      </div>
    </section>

    <!-- SLIDE 23 (Preserving existing Slide 20) -->
    <section class="slide" data-index="22" data-tag="أعمال تطبيقية (4)" data-title="ورشة تطبيقية: بناء الهيكل الشجري للملفات والمجلدات وضغطها">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">أعمال تطبيقية (4): إدارة البيانات</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">ورشة تطبيقية (4): بناء شجرة المجلدات الأكاديمية وإظهار الامتدادات المخفية</h2>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4 my-auto">
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-cyan-400">
            <div class="teacher-badge">👨‍🏫 النموذج التوجيهي للأستاذ نهمار محمد (طبقه أمام الطلبة):</div>
            <h3 class="text-sm font-bold text-white mb-2">1. تفعيل خيار إظهار امتدادات الملفات أولاً</h3>
            <div class="p-3 rounded-xl bg-slate-900/80 text-xs text-slate-200 border border-slate-700/60 mb-3 space-y-1">
              <div>• افتح مستكشف الملفات عبر <kbd class="px-1.5 py-0.5 bg-slate-800 text-cyan-300 rounded font-mono">Win + E</kbd>.</div>
              <div>• انقر على <strong>عرض (View)</strong> ➔ <strong>إظهار (Show)</strong>.</div>
              <div>• ضع علامة صح (✓) على <strong>امتدادات أسماء الملفات (File name extensions)</strong>.</div>
              <div>• ستلاحظ ظهور <code class="code-font text-emerald-400">.docx</code> و <code class="code-font text-blue-400">.pdf</code> خلف كل ملف!</div>
            </div>
            <div class="text-[11px] text-amber-300">💡 هذه الخطوة تحمي من الفيروسات وتكشف نوع الملف الحقيقي بدقة.</div>
          </div>
          <div class="glass-card p-5 rounded-2xl border-r-4 border-r-emerald-400 flex flex-col justify-between">
            <div>
              <div class="student-badge">👨‍💻 المهمة الموجهة للطلبة على حواسيبهم في المخبر:</div>
              <h3 class="text-sm font-bold text-white mb-2">المهمة: بناء الهيكل الشجري الأكاديمي لمقياس الإعلام الآلي</h3>
              <p class="text-xs text-slate-300 leading-relaxed mb-2">المطلوب التوجه إلى القرص <strong>D:</strong> وبناء الشجرة التالية:</p>
              <div class="p-3 rounded-xl bg-slate-900/80 text-xs text-slate-200 border border-slate-700/60 font-mono space-y-1 text-left" dir="ltr">
                📁 جامعة_العلوم_الاعلام_الالي_2026/<br>
                ├── 📁 01_المحاضرات_النظرية/ (ملف نصي.txt)<br>
                ├── 📁 02_الأعمال_التطبيقية/ (تقرير_ويندوز.docx)<br>
                └── 📁 03_المشاريع_والواجبات/ (علامات.xlsx)
              </div>
            </div>
            <div class="p-2.5 rounded-xl bg-emerald-950/40 border border-emerald-500/30 text-[11px] text-emerald-300 mt-2">
              🎯 <strong>مرحلة التسليم:</strong> انقر بيمين الفأرة على المجلد الرئيسي واختر <strong>Compress to ZIP file</strong> لتسليم ملف الأرشيف للأستاذ نهمار محمد.
            </div>
          </div>
        </div>
        <div class="glass-card p-2.5 rounded-xl bg-slate-900/60 text-xs text-slate-300 text-center">
          الأستاذ نهمار محمد: "من لا يرتب شجرة مجلداته الجامعية، سيقضي نصف وقته يبحث عن ملفاته التائهة ليلة الامتحان!"
        </div>
      </div>
    </section>

    <!-- SLIDE 24 (Preserving existing Slide 21) -->
    <section class="slide" data-index="23" data-tag="التحدي الختامي الشامل" data-title="المشروع التطبيقي الشامل لمقياس الإعلام الآلي">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">التحدي الختامي الشامل</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-2">مشروع التحدي العملي المتكامل: ورشة التقييم الميداني للطلبة</h2>
        </div>
        <div class="read-aloud-box text-right">
          <div class="teacher-badge">🗣️ نص إلقاء الأستاذ نهمار محمد (اقرأ للطلبة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "حان وقت وضع كل ما تعلمناه اليوم موضع التنفيذ في مشروع واحد متكامل؛ يجمع بين فحص مواصفات العتاد، لقطات الشاشة الاحترافية، بناء الشجرة، وأرشفة الملفات لتقييم علامة الأعمال التطبيقية (TP)."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-4 gap-3.5 my-auto">
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-cyan-400">
            <span class="text-[10px] font-bold text-cyan-400 block mb-1">المرحلة 1</span>
            <h3 class="text-xs font-bold text-white mb-1.5">تهيئة البيئة والشجرة</h3>
            <p class="text-[11px] text-slate-300 leading-relaxed mb-2">إنشاء المجلد الرئيسي باسم الطالب وتفريع 3 مجلدات فرعية للمحاضرات والتطبيقات والمشاريع.</p>
            <span class="text-[10px] text-cyan-300">الوزن: 5 نقاط</span>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-emerald-400">
            <span class="text-[10px] font-bold text-emerald-400 block mb-1">المرحلة 2</span>
            <h3 class="text-xs font-bold text-white mb-1.5">فحص وتوثيق العتاد</h3>
            <p class="text-[11px] text-slate-300 leading-relaxed mb-2">فتح أداة dxdiag أو About، والتقاط لقطة شاشة بـ Win+Shift+S وحفظها بصيغة PNG.</p>
            <span class="text-[10px] text-emerald-300">الوزن: 5 نقاط</span>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-amber-400">
            <span class="text-[10px] font-bold text-amber-400 block mb-1">المرحلة 3</span>
            <h3 class="text-xs font-bold text-white mb-1.5">إدارة الملفات والامتدادات</h3>
            <p class="text-[11px] text-slate-300 leading-relaxed mb-2">تفعيل إظهار الامتدادات، إنشاء ملف نصي وتسميته، وتعديل خصائصه ليصبح "للقراءة فقط".</p>
            <span class="text-[10px] text-amber-300">الوزن: 5 نقاط</span>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-purple-400">
            <span class="text-[10px] font-bold text-purple-400 block mb-1">المرحلة 4</span>
            <h3 class="text-xs font-bold text-white mb-1.5">الضغط والتسليم النهائي</h3>
            <p class="text-[11px] text-slate-300 leading-relaxed mb-2">ضغط المجلد الرئيسي بصيغة ZIP ورفعه على منصة التعليم للأستاذ نهمار محمد.</p>
            <span class="text-[10px] text-purple-300">الوزن: 5 نقاط</span>
          </div>
        </div>
        <div class="glass-card p-3 rounded-xl bg-slate-900/60 text-xs text-slate-300 flex items-center justify-between">
          <span>المدة الزمنية في المخبر: <strong>45 دقيقة</strong> | معايير التقييم: الدقة، التسمية الصحيحة، وعدم نسيان أي ملف!</span>
        </div>
      </div>
    </section>

    <!-- SLIDE 25 (Preserving existing Slide 22) -->
    <section class="slide" data-index="24" data-tag="التقييم الذاتي" data-title="اختبار التقييم الذاتي للطلبة (3 أسئلة مفصلية مشروحة)">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div>
          <span class="text-xs font-bold text-cyan-400 tracking-wider">التقييم الذاتي</span>
          <h2 class="text-2xl lg:text-3xl font-extrabold text-white mb-1">اختبار التقييم الذاتي المباشر: هل استوعبت مفاهيم اليوم؟</h2>
          <p class="text-xs text-slate-300 mb-2">انقر على الإجابة الصحيحة لكل سؤال، وسيعرض لك النظام التصحيح التلقائي والشرح الأكاديمي الفوري:</p>
        </div>
        <div class="glass-card p-5 rounded-2xl border border-cyan-500/30 my-auto shadow-2xl">
          <div id="quiz-question-container">
            <div class="flex items-center justify-between mb-3">
              <span class="text-xs font-bold text-cyan-400" id="quiz-counter">سؤال 1 من 3</span>
              <span class="text-xs text-emerald-400 font-bold" id="quiz-score-badge">النقاط: 0</span>
            </div>
            <h3 class="text-sm md:text-base font-bold text-white mb-4 leading-relaxed" id="quiz-question-text">
              1. ما هو الفارق الجوهري بين نظام البرامج الثابتة الحديث (UEFI) ونظام (BIOS) القديم؟
            </h3>
            <div class="space-y-2.5 mb-3" id="quiz-options-container">
              <button onclick="handleQuizAnswer(0)" class="w-full text-right p-3 rounded-xl bg-slate-900/80 hover:bg-slate-800 border border-slate-700/80 hover:border-cyan-400 text-xs text-slate-200 transition flex items-center justify-between">
                <span>أ) دعم أقراص أكبر من 2.2TB مع تقسيم GPT وميزة الإقلاع الآمن Secure Boot وواجهة رسومية.</span>
                <span class="text-slate-500 text-xs font-bold">A</span>
              </button>
              <button onclick="handleQuizAnswer(1)" class="w-full text-right p-3 rounded-xl bg-slate-900/80 hover:bg-slate-800 border border-slate-700/80 hover:border-cyan-400 text-xs text-slate-200 transition flex items-center justify-between">
                <span>ب) إلغاء الحاجة لوجود معالج مركزي CPU في الحاسوب.</span>
                <span class="text-slate-500 text-xs font-bold">B</span>
              </button>
              <button onclick="handleQuizAnswer(2)" class="w-full text-right p-3 rounded-xl bg-slate-900/80 hover:bg-slate-800 border border-slate-700/80 hover:border-cyan-400 text-xs text-slate-200 transition flex items-center justify-between">
                <span>ج) تخزين ملفات المستخدم تلقائياً على سلة المحذوفات.</span>
                <span class="text-slate-500 text-xs font-bold">C</span>
              </button>
            </div>
            <div id="quiz-feedback" class="hidden p-3 rounded-xl text-xs leading-relaxed"></div>
          </div>
          <div id="quiz-results" class="hidden text-center py-6">
            <div class="text-5xl mb-2">🎓</div>
            <h3 class="text-xl font-black text-white mb-1">تهانينا! أكملت اختبار المقياس بنجاح</h3>
            <p class="text-sm text-slate-300 mb-4" id="quiz-final-score-text">درجتك: 3 من 3 (100%)</p>
            <button onclick="resetQuiz()" class="px-5 py-2.5 rounded-xl bg-cyan-600 hover:bg-cyan-500 text-white font-bold text-xs transition">🔄 إعادة الاختبار مرة أخرى</button>
          </div>
        </div>
        <div class="glass-card p-2.5 rounded-xl bg-slate-900/60 text-xs text-slate-300 text-center">
          الأستاذ نهمار محمد: "الاختبار وسيلة لتثبيت المعلومات الهندسية في أذهانكم!"
        </div>
      </div>
    </section>

    <!-- SLIDE 26 (Preserving existing Slide 23) -->
    <section class="slide" data-index="25" data-tag="الخاتمة والتوصيات" data-title="خلاصة المقياس ووصايا الأستاذ نهمار محمد">
      <div class="max-w-5xl mx-auto w-full flex flex-col h-full justify-between">
        <div class="text-center">
          <span class="text-xs font-bold text-cyan-400 tracking-wider">خلاصة المحاضرة والوصايا الأكاديمية</span>
          <h2 class="text-3xl lg:text-4xl font-black text-white mb-2">تم بحمد الله وتوفيقه | حصاد مقياس الإعلام الآلي</h2>
          <p class="text-xs text-slate-300">إعداد وتقديم أستاذ المقياس: <strong class="text-amber-300 font-bold">السيد نهمار محمد (M. Nehmar Mohamed)</strong></p>
        </div>
        <div class="read-aloud-box text-right max-w-4xl mx-auto w-full">
          <div class="teacher-badge">🗣️ الكلمة الختامية للأستاذ نهمار محمد (اقرأها للطلبة في نهاية الحصة):</div>
          <p class="text-xs md:text-sm text-slate-200 leading-relaxed">
            "أبنائي وبناتي الطلبة، لقد قطعنا اليوم شوطاً رائعاً في فهم أبجديات الحوسبة. لم نعد نرى الحاسوب مجرد صندوق أسود غامض، بل صرنا ندرك حركة النبضات في الصمامات والترانزستورات، ودورة المعالجة في قلب الـ CPU، وفهمنا الرحلة الحرجة لإقلاع ويندوز عبر الـ UEFI، وكيف ننظم ملفاتنا بشجرة هرمية آمنة. الإعلام الآلي ليس مجرد مادة نظرية تمتحنون فيها، بل هو سلاحكم وأداتكم اليومية في مسيرتكم الجامعية والمهنية القادمة."
          </p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4 my-auto">
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-cyan-400 text-right">
            <h3 class="text-xs font-bold text-white mb-2 flex items-center gap-1.5"><span>🎓</span> أبرز محطات اليوم</h3>
            <ul class="text-[11px] text-slate-300 space-y-1.5 leading-relaxed">
              <li>✓ أجيال الحواسيب ودورة IPOS.</li>
              <li>✓ فحص POST والـ UEFI مقارنة بـ BIOS.</li>
              <li>✓ تشريح سطح المكتب والأيقونات وشريط المهام.</li>
              <li>✓ بطارية CMOS وضبط الوقت واللغات.</li>
              <li>✓ كشف امتدادات الملفات وبناء الشجرة الهيكلية.</li>
            </ul>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-amber-400 text-right">
            <h3 class="text-xs font-bold text-white mb-2 flex items-center gap-1.5"><span>⭐</span> وصايا الأستاذ نهمار الذهبية</h3>
            <ul class="text-[11px] text-slate-300 space-y-1.5 leading-relaxed">
              <li>1. اعتمدوا اختصارات لوحة المفاتيح؛ فهي توفر مئات الساعات سنوياً.</li>
              <li>2. لا تراكموا ملفاتكم أبداً على سطح المكتب؛ اعتمدوا التقسيم الشجري بالقرص D:.</li>
              <li>3. احتفظوا دائماً بنسخة احتياطية سحابية (Cloud Backup).</li>
            </ul>
          </div>
          <div class="glass-card p-4 rounded-xl border-t-2 border-t-emerald-400 text-right flex flex-col justify-between">
            <div>
              <h3 class="text-xs font-bold text-white mb-2 flex items-center gap-1.5"><span>📥</span> تحميل ملفات العرض</h3>
              <p class="text-[11px] text-slate-300 leading-relaxed mb-3">تنزيل نسخة العرض الرسمية بصيغة PowerPoint (.pptx) أو إعادة تشغيل العرض:</p>
            </div>
            <div class="space-y-2">
              <button onclick="downloadPPTXFile()" class="w-full py-2 px-3 rounded-xl bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-bold text-xs flex items-center justify-center gap-2 transition shadow-lg shadow-emerald-600/20 active:scale-95">
                <span>📊 تحميل عرض PowerPoint (26 شريحة)</span>
              </button>
              <button onclick="goToSlide(0)" class="w-full py-2 px-3 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 font-semibold text-xs transition border border-slate-700">↺ إعادة العرض من البداية</button>
            </div>
          </div>
        </div>
        <div class="text-center text-xs text-slate-400 pt-1">
          شكراً لحسن استماعكم ومتابعتكم • تمنيات أستاذ المقياس السيد نهمار محمد لكم بالتوفيق والنجاح الدائم!
        </div>
      </div>
    </section>

  </main>

  <!-- ==================== CINEMATIC BOTTOM CONTROL BAR ==================== -->
  <footer class="fixed bottom-0 left-0 right-0 h-16 glass-footer z-50 flex items-center justify-between px-6">
    <div class="flex items-center gap-2">
      <button onclick="prevSlide()" class="h-10 px-4 rounded-xl bg-slate-800/90 hover:bg-slate-700 text-slate-200 font-bold text-xs flex items-center gap-1.5 transition border border-slate-700/80 active:scale-95">
        <span>→</span> <span>السابق</span>
      </button>
      <button onclick="toggleCinematicAutoPlay()" id="btn-play-pause" class="h-10 px-4 rounded-xl bg-cyan-600/30 hover:bg-cyan-500/50 text-cyan-300 border border-cyan-500/40 font-bold text-xs flex items-center gap-2 transition active:scale-95">
        <span id="play-pause-icon">▶</span>
        <span id="play-pause-text">عرض تلقائي</span>
        <span id="autoplay-timer" class="hidden text-[10px] bg-cyan-500/30 px-1.5 py-0.5 rounded-full font-mono">12s</span>
      </button>
      <button onclick="nextSlide()" class="h-10 px-4 rounded-xl bg-gradient-to-r from-cyan-500 to-blue-600 hover:from-cyan-400 hover:to-blue-500 text-white font-bold text-xs flex items-center gap-1.5 transition shadow-lg shadow-cyan-500/20 active:scale-95">
        <span>التالي</span> <span>←</span>
      </button>
    </div>

    <div class="flex items-center gap-2">
      <button onclick="toggleSlideIndexDrawer()" class="h-10 px-4 rounded-xl bg-slate-800/80 hover:bg-slate-700 text-slate-300 hover:text-white text-xs font-semibold flex items-center gap-2 transition border border-slate-700/60">
        <span>📑</span> <span>فهرس الشرائح الشامل (26 شريحة)</span>
      </button>
    </div>

    <div class="hidden lg:flex items-center gap-3 text-xs text-slate-400">
      <span class="text-amber-300 font-semibold">أ. نهمار محمد</span>
      <span class="text-slate-600">•</span>
      <span class="flex items-center gap-1"><kbd class="px-1.5 py-0.5 bg-slate-800 rounded border border-slate-700 font-mono text-[10px]">Space</kbd> / <kbd class="px-1.5 py-0.5 bg-slate-800 rounded border border-slate-700 font-mono text-[10px]">←</kbd> التنقل</span>
      <span class="flex items-center gap-1"><kbd class="px-1.5 py-0.5 bg-slate-800 rounded border border-slate-700 font-mono text-[10px]">F</kbd> ملء الشاشة</span>
      <span class="flex items-center gap-1"><kbd class="px-1.5 py-0.5 bg-slate-800 rounded border border-slate-700 font-mono text-[10px]">M</kbd> كتم الصوت</span>
    </div>
  </footer>

  <!-- ==================== MODALS & DRAWERS ==================== -->

  <!-- Slide Index Drawer (26 slides) -->
  <div id="slide-drawer" class="fixed inset-0 bg-black/75 backdrop-blur-md z-[100] hidden items-center justify-center p-6">
    <div class="glass-card max-w-4xl w-full max-h-[85vh] rounded-3xl p-6 flex flex-col justify-between border border-cyan-500/30">
      <div class="flex items-center justify-between pb-3 border-b border-slate-700/60 mb-3">
        <div class="flex items-center gap-2">
          <span class="text-xl">📑</span>
          <h3 class="text-sm md:text-base font-bold text-white">فهرس المحاضرة الشامل - الأستاذ نهمار محمد (26 شريحة)</h3>
        </div>
        <button onclick="toggleSlideIndexDrawer()" class="w-8 h-8 rounded-lg bg-slate-800 text-slate-300 hover:text-white flex items-center justify-center">✕</button>
      </div>
      <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2 overflow-y-auto pr-1" id="drawer-slide-list"></div>
      <div class="pt-3 border-t border-slate-700/60 mt-3 text-center text-xs text-slate-400">
        اختر أي شريحة للانتقال الفوري إليها
      </div>
    </div>
  </div>

  <!-- Teacher Speech Teleprompter Modal -->
  <div id="speech-modal" class="fixed inset-0 bg-black/75 backdrop-blur-md z-[100] hidden items-center justify-center p-6">
    <div class="glass-card max-w-3xl w-full rounded-3xl p-6 border border-cyan-500/40">
      <div class="flex items-center justify-between pb-3 border-b border-slate-700/60 mb-3">
        <div class="flex items-center gap-2">
          <span class="text-xl">🎙️</span>
          <h3 class="text-sm md:text-base font-bold text-white">شاشة إلقاء الأستاذ نهمار محمد (اقرأ هذا النص للطلبة)</h3>
        </div>
        <button onclick="toggleTeacherSpeechModal()" class="w-8 h-8 rounded-lg bg-slate-800 text-slate-300 hover:text-white flex items-center justify-center">✕</button>
      </div>
      <div class="p-4 rounded-2xl bg-slate-950/80 border border-cyan-500/30 text-xs md:text-sm text-slate-100 leading-relaxed max-h-72 overflow-y-auto mb-4" id="speech-modal-content"></div>
      <div class="flex items-center justify-between">
        <span class="text-xs text-slate-400">مكتوب بصياغة إلقاء مباشرة وميسرة</span>
        <button onclick="toggleTeacherSpeechModal()" class="px-5 py-2 rounded-xl bg-cyan-600 hover:bg-cyan-500 text-white font-bold text-xs transition">إغلاق</button>
      </div>
    </div>
  </div>

  <!-- ==================== JAVASCRIPT LOGIC ==================== -->
  <script>
    let currentSlideIndex = 0;
    const slides = document.querySelectorAll('.slide');
    const totalSlides = slides.length;
    let isSoundMuted = false;
    let autoPlayInterval = null;
    let autoPlaySecondsRemaining = 12;
    const AUTOPLAY_DURATION = 12;
    let isHeaderVisible = true;

    // Auto-collapsing Header Management
    function hideHeader() {
      const header = document.getElementById('main-header');
      const pill = document.getElementById('header-reveal-pill');
      if (header && !header.classList.contains('header-hidden')) {
        header.classList.add('header-hidden');
        if (pill) pill.classList.add('pill-visible');
        isHeaderVisible = false;
      }
    }

    function showHeader() {
      const header = document.getElementById('main-header');
      const pill = document.getElementById('header-reveal-pill');
      if (header && header.classList.contains('header-hidden')) {
        header.classList.remove('header-hidden');
        if (pill) pill.classList.remove('pill-visible');
        isHeaderVisible = true;
      }
    }

    function handleViewportClick(e) {
      // If user clicks in the middle/content of the slide, hide the header so titles get full unobstructed view!
      if (isHeaderVisible && e.clientY > 70) {
        hideHeader();
      }
    }

    // Speech scripts for all 26 slides for Mr. Nehmar Mohamed
    const teacherSpeechScripts = [
      "بسم الله الرحمن الرحيم. مرحباً بكم أعزائي الطلبة في مقياس الإعلام الآلي. معكم أستاذ المقياس السيد نهمار محمد. اليوم سنبدأ معاً رحلة ممتعة وتأسيسية من الصفر، لنفهم كيف يعمل الحاسوب، ونكتشف معاً أسرار نظام Windows وتطبيقاته العملية خطوة بخطوة.",
      "أعزائي الطلبة، كلمة Informatique هي دمج لكلمتين: Information (معلومات) و Automatique (آلي). أي علم معالجة المعلومات آلياً. وتذكروا دائماً: البيانات هي المادة الخام الصامتة، بينما المعلومات هي الناتج المفيد المنظم الذي يساعدنا على اتخاذ القرار.",
      "في الجيل الأول (1940-1956)، استخدم المهندسون الصمامات المفرغة وهي مصابيح زجاجية تضيء لتمرير التيار وتطفأ لمنعه. كان حاسوب ENIAC يزن 30 طناً ويستهلك كهرباء مدينة كاملة! وقصة كلمة Bug جاءت عندما دخلت حشرة عثة حقيقية في أحد الصمامات وأوقفت النظام.",
      "في الجيل الثاني (1956-1963)، حدثت الثورة الكبرى باختراع الترانزستور. مفتاح كهربائي مجهري صلب مصنوع من السيليكون؛ صغّر الحجم بمئات المرات وخفض الحرارة، وظهرت لغات برمجة أسهل بكثير مثل FORTRAN و COBOL.",
      "في الجيل الثالث والرابع، جمع العلماء آلاف الترانزستورات على رقاقة سيليكون واحدة فظهرت الدارة المتكاملة IC، ثم طبعوا المعالج كاملاً على شريحة واحدة فظهر المعالج الدقيق Microprocessor وولد الحاسوب الشخصي PC في كل بيت ومكتب.",
      "نصنف الحواسيب لأربعة أصناف: الحواسيب العملاقة Supercomputers للمناخ والفيزياء، الحواسيب الكبيرة Mainframes للبنوك واستقرار المعاملات، الحواسيب الشخصية PCs للاستخدام اليومي، والأنظمة المدمجة Embedded داخل السيارات والغسالات والتي تمثل 90% من معالجات الكوكب!",
      "دورة IPOS هي قانون عمل أي حاسوب: إدخال البيانات (Input) عبر الكيبورد ➔ معالجتها في الـ CPU (Processing) ➔ تخزينها مؤقتاً في RAM أو دائماً في SSD (Storage) ➔ وإظهارها للمستخدم على الشاشة أو الطابعة كمعلومات مفيدة (Output).",
      "يتكون المعالج CPU من 3 غرف أساسية: وحدة التحكم CU كشرطي المرور الذي يوجه الأوامر، وحدة الحساب والمنطق ALU كالآلة الحاسبة المنفذة، وذاكرة الكاش والمسجلات Registers فائقة السرعة لتفادي بطء الذاكرة العادية.",
      "يتحدث الحاسوب بلغة 0 و 1 فقط لأن الكهرباء تملك حالتين: تيار يمر (1) أو لا يمر (0). وكل 8 بتات تكون 1 بايت القادر على تخزين حرف واحد. و 1024 بايت = 1 كيلوبايت، ومضاعفاتها تصل للتيرابايت.",
      "نظام التشغيل OS هو المدير العام الحاكم؛ بدون ويندوز لن نرى إلا شاشة سوداء. ورحلة ويندوز بدأت من Windows 1.0 البسيط، ثم ثورة Windows 95 وزر ابدأ، واستقرار XP، وأناقة 7، وصولاً إلى Windows 11 وتكامل الذكاء الاصطناعي Copilot.",
      "عند الضغط على زر التشغيل: تصل الطاقة ➔ يبدأ فحص العتاد الذاتي POST (وإذا حدث عطل تصدر صفارات Beep Codes) ➔ يستيقظ UEFI لفحص الأقراص بحماية Secure Boot ➔ ثم تُحمّل نواة ويندوز ntoskrnl.exe في RAM وتفتح شاشة الدخول.",
      "في هذا التطبيق العملي، سأريكم أولاً كيف ندخل للـ BIOS بالنقر المتكرر على F2 أو Del فور التشغيل. ثم سأطلب من كل طالب الدخول للوضع الآمن Safe Mode بالضغط على زر Shift أثناء النقر على Restart لإصلاح مشاكل الشاشة الزرقاء.",
      "هذا هو محاكي سطح مكتب Windows 11 التفاعلي. انظروا إلى كيفية توزيع الأيقونات الرئيسية: هذا الكمبيوتر، سلة المحذوفات، ملفات الطالب، لوحة التحكم، وشبكة الاتصال. عند النقر على أي أيقونة ستسمعون نغمة تفاعل ويظهر الشرح المباشر.",
      "سنشرح الآن الدور الأكاديمي للأيقونات الخمسة: هذا الكمبيوتر لفحص سعات الأقراص، سلة المحذوفات كصمام أمان قبل الحذف النهائي، مجلد المستخدم المحمي بالأذونات، لوحة التحكم لإدارة البرامج، وشبكة الاتصال لتبادل الملفات في المخبر.",
      "في شريط المهام، يجب التمييز بدقة بين خيارات الطاقة: إيقاف التشغيل يفصل الكهرباء، إعادة التشغيل تفرغ RAM، وضع السكون Sleep يبقي كهرباء ضعيفة للاستيقاظ السريع، أما وضع الإسبات Hibernate فيحفظ RAM في القرص الصلب ويفصل الكهرباء 100% دون ضياع بياناتك!",
      "واجهة ويندوز تشبه طاولة مكتبك الحقيقية: سطح المكتب Desktop هو مساحة العمل، شريط المهام Taskbar يعرض برامجك المفتوحة والساعة، والأيقونات تنقسم لأيقونات نظام، وأيقونات اختصار تتميز بسهم أزرق صغير، وأيقونات ملفات ومجلدات.",
      "سأريكم الآن عملياً كيف نستعيد أيقونة 'هذا الكمبيوتر' الأصلية عبر إعدادات النسق Themes بدل الاختصار المشوه، ثم ستقومون أنتم بإنشاء أسطح مكتب افتراضية بـ Win + Tab لتنظيم الدراسة والترفيه، وتجربة خماسية الاختصارات السريعة.",
      "حاسوبكم يعرف الوقت دائماً بفضل بطارية CMOS الصغيرة CR2032 على اللوحة الأم التي تغذي شريحة الوقت حتى عند فصل الكهرباء تماماً. وإذا رجع التاريخ لسنة 2000 فيجب تغييرها. وفرقوا بين لغة عرض ويندوز ولغة لوحة المفاتيح والتبديل بـ Alt + Shift.",
      "سأشرح لكم الآن أوامر فحص العتاد: أمر dxdiag وأمر msinfo32 من نافذة Win + R. وعلى كل طالب منكم الآن استخراج اسم وسرعة معالجه، وحجم الرام، والتأكد من أن نظامه 64-bit ليتحمل برامج التخصص الجامعية.",
      "اسم أي ملف يتكون من: [اسم_الملف].[الامتداد]. الاسم يختاره المستخدم (ويمنع استخدام المحارف التسعة مثل النجمة وعلامة الاستفهام)، بينما الامتداد يحدد لويندوز البرنامج المسؤول عن فتحه. والمجلد هو مجرد حاوية هرمية لتنظيم المسار.",
      "احفظوا هذا الجدول: .docx للورد، .pdf للوثائق المحمولة، .xlsx للإكسل، .pptx للبوربوينت، .jpg و .png للصور، .mp4 للفيديو، .zip للضغط، و .exe للملفات التنفيذية التي يجب الحذر منها إن كانت مجهولة المصدر.",
      "الفرق بين النسخ Ctrl+C والقص Ctrl+X: الأول يكرر الملف والثاني ينقله. والفرق بين Delete و Shift+Delete: الأول ينقله لسلة المحذوفات لاسترجاعه، والثاني يحذفه نهائياً مباشرة لتوفير المساحة بحذر.",
      "سأريكم أولاً كيف نفعل خيار إظهار امتدادات أسماء الملفات File name extensions من قائمة View لحماية الجهاز من الفيروسات. ثم ستقومون أنتم في المخبر ببناء الهيكلة الشجرية لمجلد الإعلام الآلي 2026 وضغطه بصيغة ZIP.",
      "هذا هو مشروع التحدي التقييمي الشامل: تهيئة الشجرة، التقاط صورة لعتاد الجهاز، إنشاء وتسمية الملفات وتعديل خصائصها، وضغط المجلد بالكامل لتسليمه للأستاذ نهمار محمد في مخبر الأعمال التطبيقية.",
      "اختبار التقييم الذاتي الفوري: اختبروا معلوماتكم في الأسئلة الثلاثة حول UEFI وبطارية CMOS والفرق بين أوامر الحذف لتروا درجتكم النهائية والشرح التفصيلي لكل إجابة.",
      "تم بحمد الله وتوفيقه. أشكركم على حسن الاستماع والانتباه؛ لا تنسوا الوصايا الثلاث: احترفوا الاختصارات، رتبوا ملفاتكم في مجلدات هرمية، واحتفظوا دائماً بنسخة سحابية احتياطية. باب الأسئلة مفتوح لكم الآن!"
    ];

    // Audio SFX Synthesizer
    let audioCtx = null;
    function getAudioContext() {
      if (!audioCtx) {
        const AudioContextClass = window.AudioContext || window.webkitAudioContext;
        if (AudioContextClass) audioCtx = new AudioContextClass();
      }
      if (audioCtx && audioCtx.state === 'suspended') audioCtx.resume();
      return audioCtx;
    }

    function playSound(type) {
      if (isSoundMuted) return;
      try {
        const ctx = getAudioContext();
        if (!ctx) return;
        const now = ctx.currentTime;
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.connect(gain);
        gain.connect(ctx.destination);

        if (type === 'slide') {
          osc.type = 'sine';
          osc.frequency.setValueAtTime(150, now);
          osc.frequency.exponentialRampToValueAtTime(320, now + 0.15);
          gain.gain.setValueAtTime(0.06, now);
          gain.gain.exponentialRampToValueAtTime(0.001, now + 0.18);
          osc.start(now);
          osc.stop(now + 0.18);
        } else if (type === 'click') {
          osc.type = 'triangle';
          osc.frequency.setValueAtTime(550, now);
          osc.frequency.exponentialRampToValueAtTime(220, now + 0.05);
          gain.gain.setValueAtTime(0.04, now);
          gain.gain.exponentialRampToValueAtTime(0.001, now + 0.05);
          osc.start(now);
          osc.stop(now + 0.05);
        } else if (type === 'fanfare') {
          const notes = [523.25, 659.25, 783.99, 1046.50];
          notes.forEach((freq, idx) => {
            const noteOsc = ctx.createOscillator();
            const noteGain = ctx.createGain();
            noteOsc.connect(noteGain);
            noteGain.connect(ctx.destination);
            noteOsc.type = 'sine';
            noteOsc.frequency.setValueAtTime(freq, now + idx * 0.08);
            noteGain.gain.setValueAtTime(0.05, now + idx * 0.08);
            noteGain.gain.exponentialRampToValueAtTime(0.001, now + idx * 0.08 + 0.28);
            noteOsc.start(now + idx * 0.08);
            noteOsc.stop(now + idx * 0.08 + 0.28);
          });
        }
      } catch (e) {
        console.warn('Audio play error:', e);
      }
    }

    function toggleSound() {
      isSoundMuted = !isSoundMuted;
      document.getElementById('sound-icon').textContent = isSoundMuted ? '🔇' : '🔊';
      if (!isSoundMuted) playSound('click');
    }

    // Slide Navigation
    function updateSlideUI() {
      slides.forEach((slide, idx) => {
        slide.classList.remove('active', 'prev');
        if (idx === currentSlideIndex) {
          slide.classList.add('active');
        } else if (idx < currentSlideIndex) {
          slide.classList.add('prev');
        }
      });

      const activeSlide = slides[currentSlideIndex];
      const tag = activeSlide.getAttribute('data-tag') || '';
      const title = activeSlide.getAttribute('data-title') || '';
      document.getElementById('slide-tracker').textContent = `شريحة ${currentSlideIndex + 1} من ${totalSlides}`;
      document.getElementById('slide-title-preview').textContent = `${tag}: ${title}`;

      const progressPct = ((currentSlideIndex + 1) / totalSlides) * 100;
      document.getElementById('progress-bar').style.width = `${progressPct}%`;

      document.getElementById('speech-modal-content').innerHTML = `
        <div class="mb-2 text-cyan-400 font-bold">📍 الشريحة ${currentSlideIndex + 1}: ${title}</div>
        <p class="text-slate-200 leading-relaxed font-medium">${teacherSpeechScripts[currentSlideIndex] || 'اقرأ المحتوى المكتوب في الشريحة للطلبة.'}</p>
      `;

      playSound('slide');
    }

    function goToSlide(index) {
      if (index >= 0 && index < totalSlides) {
        currentSlideIndex = index;
        updateSlideUI();
      }
    }

    function nextSlide() {
      if (currentSlideIndex < totalSlides - 1) {
        currentSlideIndex++;
        updateSlideUI();
      } else {
        if (autoPlayInterval) {
          currentSlideIndex = 0;
          updateSlideUI();
        }
      }
      resetAutoPlayTimer();
    }

    function prevSlide() {
      if (currentSlideIndex > 0) {
        currentSlideIndex--;
        updateSlideUI();
      }
      resetAutoPlayTimer();
    }

    function toggleCinematicAutoPlay() {
      if (autoPlayInterval) {
        clearInterval(autoPlayInterval);
        autoPlayInterval = null;
        document.getElementById('play-pause-icon').textContent = '▶';
        document.getElementById('play-pause-text').textContent = 'عرض تلقائي';
        document.getElementById('autoplay-timer').classList.add('hidden');
      } else {
        autoPlaySecondsRemaining = AUTOPLAY_DURATION;
        document.getElementById('play-pause-icon').textContent = '⏸';
        document.getElementById('play-pause-text').textContent = 'إيقاف مؤقت';
        document.getElementById('autoplay-timer').classList.remove('hidden');
        document.getElementById('autoplay-timer').textContent = `${autoPlaySecondsRemaining}s`;

        autoPlayInterval = setInterval(() => {
          autoPlaySecondsRemaining--;
          if (autoPlaySecondsRemaining <= 0) {
            nextSlide();
            autoPlaySecondsRemaining = AUTOPLAY_DURATION;
          }
          document.getElementById('autoplay-timer').textContent = `${autoPlaySecondsRemaining}s`;
        }, 1000);
      }
    }

    function resetAutoPlayTimer() {
      if (autoPlayInterval) {
        autoPlaySecondsRemaining = AUTOPLAY_DURATION;
        document.getElementById('autoplay-timer').textContent = `${autoPlaySecondsRemaining}s`;
      }
    }

    function toggleFullScreen() {
      if (!document.fullscreenElement) {
        document.documentElement.requestFullscreen().catch(() => {});
      } else {
        if (document.exitFullscreen) document.exitFullscreen();
      }
    }

    function toggleSlideIndexDrawer() {
      const drawer = document.getElementById('slide-drawer');
      const isHidden = drawer.classList.contains('hidden');
      if (isHidden) {
        renderDrawerSlideList();
        drawer.classList.remove('hidden');
        drawer.classList.add('flex');
      } else {
        drawer.classList.add('hidden');
        drawer.classList.remove('flex');
      }
    }

    function renderDrawerSlideList() {
      const container = document.getElementById('drawer-slide-list');
      container.innerHTML = '';
      slides.forEach((slide, idx) => {
        const title = slide.getAttribute('data-title') || `شريحة ${idx + 1}`;
        const tag = slide.getAttribute('data-tag') || '';
        const item = document.createElement('button');
        item.className = `p-2.5 rounded-xl text-right transition border text-xs flex flex-col justify-between ${
          idx === currentSlideIndex
            ? 'bg-cyan-950/80 border-cyan-400 text-cyan-200'
            : 'bg-slate-900/80 border-slate-700/80 hover:bg-slate-800 text-slate-300'
        }`;
        item.innerHTML = `
          <div class="flex items-center justify-between w-full mb-1">
            <span class="text-[10px] font-bold ${idx === currentSlideIndex ? 'text-cyan-400' : 'text-slate-500'}">شريحة ${idx + 1}</span>
            <span class="text-[10px] text-slate-400">${tag}</span>
          </div>
          <div class="font-bold text-white truncate w-full text-[11px]">${title}</div>
        `;
        item.onclick = () => {
          goToSlide(idx);
          toggleSlideIndexDrawer();
        };
        container.appendChild(item);
      });
    }

    function toggleTeacherSpeechModal() {
      const modal = document.getElementById('speech-modal');
      modal.classList.toggle('hidden');
      modal.classList.toggle('flex');
    }

    // Keyboard Shortcuts
    window.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowLeft' || e.key === ' ' || e.key === 'PageDown') {
        e.preventDefault();
        nextSlide();
      } else if (e.key === 'ArrowRight' || e.key === 'PageUp') {
        e.preventDefault();
        prevSlide();
      } else if (e.key === 'f' || e.key === 'F') {
        toggleFullScreen();
      } else if (e.key === 'm' || e.key === 'M') {
        toggleSound();
      } else if (e.key === 'p' || e.key === 'P') {
        toggleCinematicAutoPlay();
      } else if (e.key === 'h' || e.key === 'H') {
        if (isHeaderVisible) hideHeader(); else showHeader();
      } else if (e.key === 'Escape') {
        document.getElementById('slide-drawer').classList.add('hidden');
        document.getElementById('slide-drawer').classList.remove('flex');
        document.getElementById('speech-modal').classList.add('hidden');
        document.getElementById('speech-modal').classList.remove('flex');
      }
    });

    // In-slide simulators
    function simulateIPOSCycle() {
      playSound('click');
      const notif = document.getElementById('ipos-notification');
      notif.classList.remove('hidden');
      setTimeout(() => { playSound('fanfare'); }, 600);
      setTimeout(() => { notif.classList.add('hidden'); }, 5000);
    }

    function updateBinaryPreview() {
      const input = document.getElementById('binary-input').value;
      const output = document.getElementById('binary-output');
      if (input.length > 0) {
        const charCode = input.charCodeAt(0);
        const binaryStr = charCode.toString(2).padStart(8, '0');
        output.textContent = binaryStr;
        playSound('click');
      } else {
        output.textContent = '00000000';
      }
    }

    function clickSimIcon(description) {
      playSound('click');
      const banner = document.getElementById('sim-info-banner');
      if (banner) {
        banner.innerHTML = `💡 <strong>الشرح الأكاديمي للأيقونة:</strong><br>${description}`;
      }
    }

    function toggleSimStartMenu() {
      playSound('click');
      clickSimIcon('🪟 قائمة ابدأ (Start Menu): النفاذ السريع للتطبيقات المثبتة وملفات المستخدم وخيارات الطاقة (Shut Down, Restart, Sleep, Hibernate).');
    }

    // Quiz System
    const quizQuestions = [
      {
        question: "1. ما هو الفارق الجوهري بين نظام البرامج الثابتة الحديث (UEFI) ونظام (BIOS) القديم؟",
        options: [
          "أ) دعم أقراص أكبر من 2.2TB مع تقسيم GPT وميزة الإقلاع الآمن Secure Boot وواجهة رسومية.",
          "ب) إلغاء الحاجة لوجود معالج مركزي CPU في الحاسوب.",
          "ج) تخزين ملفات المستخدم تلقائياً على سلة المحذوفات."
        ],
        correct: 0,
        explanation: "صحيح تماماً! نظام UEFI يدعم معمارية GPT الحديثة والواجهات الرسومية ويوفر ميزة Secure Boot لمنع البرمجيات الخبيثة من الإقلاع قبل ويندوز."
      },
      {
        question: "2. إذا كان تاريخ الحاسوب وساعته يرجعان دائماً إلى سنة 2000 عند كل تشغيل، فما هو التشخيص الهندسي؟",
        options: [
          "أ) امتلاء القرص الصلب C: بالملفات المؤقتة.",
          "ب) نفاد بطارية اللوحة الأم CMOS (CR2032) المسؤولة عن تغذية شريحة الوقت عند انقطاع الكهرباء.",
          "ج) خلل في سلك الشاشة أو لوحة المفاتيح."
        ],
        correct: 1,
        explanation: "إجابة ممتازة! بطارية CMOS الصغيرة على اللوحة الأم هي المسؤولة عن تزويد شريحة الوقت RTC بالطاقة عند فصل الكهرباء، ونفادها يعيد الساعة لتاريخ المصنع الافتراضي."
      },
      {
        question: "3. ما الفارق الجوهري بين الحذف بمفتاح (Delete) والحذف باختصار (Shift + Delete)؟",
        options: [
          "أ) لا يوجد أي فرق بينهما فكلاهما يقوم بنفس العمل بالضبط.",
          "ب) الأول ينقل الملف لسلة المحذوفات لاسترجاعه، بينما الثاني يحذفه نهائياً مباشرة لتوفير المساحة.",
          "ج) الحذف بـ Shift+Delete يغير فقط اسم الملف دون حذفه."
        ],
        correct: 1,
        explanation: "أحسنت! مفتاح Delete العادي هو حذف وقائي يمكن التراجع عنه من سلة المحذوفات، أما Shift+Delete فيزيل مؤشر الملف نهائياً من نظام الملفات."
      }
    ];

    let currentQuizIndex = 0;
    let quizScore = 0;

    function handleQuizAnswer(selectedOption) {
      const q = quizQuestions[currentQuizIndex];
      const feedback = document.getElementById('quiz-feedback');
      feedback.classList.remove('hidden');

      if (selectedOption === q.correct) {
        playSound('fanfare');
        quizScore++;
        document.getElementById('quiz-score-badge').textContent = `النقاط: ${quizScore}`;
        feedback.className = 'p-3 rounded-xl text-xs leading-relaxed bg-emerald-950/60 text-emerald-300 border border-emerald-500/40';
        feedback.innerHTML = `✓ <strong>إجابة صحيحة!</strong> ${q.explanation}`;
      } else {
        playSound('click');
        feedback.className = 'p-3 rounded-xl text-xs leading-relaxed bg-red-950/60 text-red-300 border border-red-500/40';
        feedback.innerHTML = `✗ <strong>إجابة خاطئة!</strong> ${q.explanation}`;
      }

      setTimeout(() => {
        currentQuizIndex++;
        if (currentQuizIndex < quizQuestions.length) {
          loadQuizQuestion(currentQuizIndex);
        } else {
          showQuizResults();
        }
      }, 2200);
    }

    function loadQuizQuestion(idx) {
      const q = quizQuestions[idx];
      document.getElementById('quiz-counter').textContent = `سؤال ${idx + 1} من ${quizQuestions.length}`;
      document.getElementById('quiz-question-text').textContent = q.question;
      document.getElementById('quiz-feedback').classList.add('hidden');

      const optionsContainer = document.getElementById('quiz-options-container');
      optionsContainer.innerHTML = '';
      q.options.forEach((optText, optIdx) => {
        const btn = document.createElement('button');
        btn.className = 'w-full text-right p-3 rounded-xl bg-slate-900/80 hover:bg-slate-800 border border-slate-700/80 hover:border-cyan-400 text-xs text-slate-200 transition flex items-center justify-between';
        btn.innerHTML = `
          <span>${optText}</span>
          <span class="text-slate-500 text-xs font-bold">${String.fromCharCode(65 + optIdx)}</span>
        `;
        btn.onclick = () => handleQuizAnswer(optIdx);
        optionsContainer.appendChild(btn);
      });
    }

    function showQuizResults() {
      document.getElementById('quiz-question-container').classList.add('hidden');
      document.getElementById('quiz-results').classList.remove('hidden');
      const pct = Math.round((quizScore / quizQuestions.length) * 100);
      document.getElementById('quiz-final-score-text').textContent = `درجتك النهائية: ${quizScore} من ${quizQuestions.length} (${pct}%)`;
      playSound('fanfare');
    }

    function resetQuiz() {
      currentQuizIndex = 0;
      quizScore = 0;
      document.getElementById('quiz-score-badge').textContent = 'النقاط: 0';
      document.getElementById('quiz-results').classList.add('hidden');
      document.getElementById('quiz-question-container').classList.remove('hidden');
      loadQuizQuestion(0);
    }

    // Dynamic Canvas
    function initCanvas() {
      const canvas = document.getElementById('bg-canvas');
      const ctx = canvas.getContext('2d');
      let width = canvas.width = window.innerWidth;
      let height = canvas.height = window.innerHeight;

      window.addEventListener('resize', () => {
        width = canvas.width = window.innerWidth;
        height = canvas.height = window.innerHeight;
      });

      const numParticles = 45;
      const particles = [];
      for (let i = 0; i < numParticles; i++) {
        particles.push({
          x: Math.random() * width,
          y: Math.random() * height,
          vx: (Math.random() - 0.5) * 0.35,
          vy: (Math.random() - 0.5) * 0.35,
          radius: Math.random() * 1.5 + 0.8,
          alpha: Math.random() * 0.5 + 0.2
        });
      }

      function render() {
        ctx.clearRect(0, 0, width, height);
        for (let i = 0; i < particles.length; i++) {
          for (let j = i + 1; j < particles.length; j++) {
            const dx = particles[i].x - particles[j].x;
            const dy = particles[i].y - particles[j].y;
            const dist = Math.sqrt(dx * dx + dy * dy);
            if (dist < 130) {
              ctx.beginPath();
              ctx.moveTo(particles[i].x, particles[i].y);
              ctx.lineTo(particles[j].x, particles[j].y);
              ctx.strokeStyle = `rgba(56, 189, 248, ${0.11 * (1 - dist / 130)})`;
              ctx.lineWidth = 0.7;
              ctx.stroke();
            }
          }
        }

        for (let i = 0; i < particles.length; i++) {
          const p = particles[i];
          p.x += p.vx;
          p.y += p.vy;
          if (p.x < 0) p.x = width;
          if (p.x > width) p.x = 0;
          if (p.y < 0) p.y = height;
          if (p.y > height) p.y = 0;

          ctx.beginPath();
          ctx.arc(p.x, p.y, p.radius, 0, Math.PI * 2);
          ctx.fillStyle = `rgba(56, 189, 248, ${p.alpha})`;
          ctx.fill();
        }
        requestAnimationFrame(render);
      }
      render();
    }

    // Embedded PPTX Downloader
    const PPTX_BASE64 = "''' + b64_pptx + '''";
    function downloadPPTXFile() {
      try {
        const binaryString = window.atob(PPTX_BASE64);
        const len = binaryString.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
          bytes[i] = binaryString.charCodeAt(i);
        }
        const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = 'Informatique_Cours_Complet.pptx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        playSound('fanfare');
        alert('✅ بدأ تحميل ملف PowerPoint (26 شريحة كاملة) بنجاح إلى مجلد التنزيلات (Downloads)!');
      } catch (err) {
        alert('حدث خطأ أثناء فك التشفير: ' + err.message);
      }
    }

    window.addEventListener('DOMContentLoaded', () => {
      initCanvas();
      updateSlideUI();
    });
  </script>
</body>
</html>'''

index_path = r"C:\Users\Moham\.gemini\antigravity\scratch\informatique_presentation\index.html"
with open(index_path, "w", encoding="utf-8") as f:
    f.write(html_code)

print("Generated 26-slide index.html successfully.")

# Also copy to cinematic_presentation.html in brain
brain_cinematic = r"C:\Users\Moham\.gemini\antigravity\brain\8e28d94e-0064-4605-ad65-07261bf39388\cinematic_presentation.html"
with open(brain_cinematic, "w", encoding="utf-8") as f:
    f.write(html_code)

print("Updated cinematic_presentation.html in brain.")
