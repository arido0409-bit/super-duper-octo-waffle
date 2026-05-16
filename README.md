<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>نظام شرق الاحتراف المحاسبي الذكي — v3.1</title>
<meta name="theme-color" content="#C9A84C">
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;800;900&family=Cairo:wght@300;400;600;700;900&display=swap" rel="stylesheet">
<style>
:root{--gold:#C9A84C;--gold-light:#E8C97A;--gold-dark:#8B6914;--bg:#0A0A0A;--card:#111;--card2:#161616;--card3:#1C1C1C;--bord:rgba(201,168,76,0.3);--bordh:rgba(201,168,76,0.7);--tp:#F0E6C8;--ts:#A89060;--tm:#5A4A2A;--ok:#2ECC71;--err:#E74C3C;--warn:#F39C12;--inf:#3498DB;--r:12px;--rs:8px}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Tajawal','Cairo',sans-serif;background:var(--bg);color:var(--tp);min-height:100vh;direction:rtl;overflow-x:hidden}
body::before{content:'';position:fixed;inset:0;background-image:linear-gradient(rgba(201,168,76,0.03) 1px,transparent 1px),linear-gradient(90deg,rgba(201,168,76,0.03) 1px,transparent 1px);background-size:40px 40px;pointer-events:none;z-index:0}

/* ── SPLASH ── */
#splash{position:fixed;inset:0;background:#000;z-index:9999;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:20px;transition:opacity .8s}
#splash.out{opacity:0;pointer-events:none}
.slogo{width:130px;height:130px;border:3px solid var(--gold);border-radius:50%;display:flex;align-items:center;justify-content:center;position:relative;animation:pr 2s infinite}
.slogo::before{content:'';position:absolute;inset:-8px;border-radius:50%;border:1px solid rgba(201,168,76,0.3);animation:pr 2s infinite .3s}
.stitle{font-size:24px;font-weight:900;background:linear-gradient(135deg,var(--gold-light),var(--gold));-webkit-background-clip:text;-webkit-text-fill-color:transparent;text-align:center}
.ssub{color:var(--ts);font-size:12px;letter-spacing:2px}
.sbar{width:260px;height:2px;background:rgba(201,168,76,0.2);border-radius:2px;overflow:hidden}
.sbar-f{height:100%;background:linear-gradient(90deg,var(--gold-dark),var(--gold-light));animation:lb 2.8s ease forwards}
@keyframes lb{from{width:0}to{width:100%}}
@keyframes pr{0%,100%{box-shadow:0 0 0 0 rgba(201,168,76,0.3)}50%{box-shadow:0 0 20px 8px rgba(201,168,76,0.08)}}

/* ── AUTH ── */
#auth-screen{position:fixed;inset:0;z-index:10000;display:flex;align-items:center;justify-content:center;overflow:hidden;transition:opacity .5s}
#auth-screen.hidden{display:none}
.auth-bg{position:absolute;inset:0;background:radial-gradient(ellipse at 50% 30%,#1a1200 0%,#0a0800 60%,#000 100%)}
.auth-bg::before{content:'';position:absolute;inset:0;background-image:linear-gradient(rgba(201,168,76,0.04) 1px,transparent 1px),linear-gradient(90deg,rgba(201,168,76,0.04) 1px,transparent 1px);background-size:44px 44px}
.auth-float{position:absolute;opacity:.15;color:var(--gold)}
.auth-float.tl{top:12%;left:8%;font-size:52px}
.auth-float.tr{top:10%;right:7%;font-size:48px}
.auth-float.bl{bottom:18%;left:6%;font-size:44px}
.auth-float.br{bottom:15%;right:8%;font-size:50px}
.auth-main{position:relative;z-index:1;width:100%;max-width:1200px;display:grid;grid-template-columns:1fr auto 1fr;gap:40px;align-items:center;padding:0 40px}
.auth-panel-left,.auth-panel-right{display:flex;flex-direction:column;gap:14px}
.auth-panel-card{background:rgba(20,16,4,0.85);border:1px solid rgba(201,168,76,0.25);border-radius:14px;padding:18px;backdrop-filter:blur(10px);transition:all .3s}
.auth-panel-card:hover{border-color:rgba(201,168,76,0.5);transform:translateY(-2px)}
.auth-panel-title{font-size:13px;font-weight:700;color:var(--gold);margin-bottom:10px}
.auth-panel-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:6px}
.auth-panel-cell{height:24px;background:rgba(201,168,76,0.08);border:1px solid rgba(201,168,76,0.15);border-radius:4px}
.auth-panel-cell.h{background:rgba(201,168,76,0.18);border-color:rgba(201,168,76,0.4)}
.auth-dept-card{background:rgba(20,16,4,0.9);border:1px solid rgba(201,168,76,0.35);border-radius:14px;padding:14px;display:flex;align-items:center;gap:12px}
.auth-dept-ico">⚙️</div>
        <div><div class="auth-dept-name">قسم الصب والتشكيل</div><div class="auth-dept-sub">Advanced Die-Casting & Forming Dept.</div></div>
      </div>
      <div class="auth-panel-card" style="padding:12px 16px">
        <div style="display:flex;gap:8px;align-items:center;margin-bottom:10px">
          <div style="width:8px;height:8px;border-radius:50%;background:var(--ok);box-shadow:0 0 6px var(--ok)"></div>
          <span style="font-size:11px;color:var(--ts)">النظام يعمل بكفاءة</span>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px">
          <div style="text-align:center;padding:8px;background:rgba(201,168,76,0.06);border-radius:8px;border:1px solid rgba(201,168,76,0.15)">
            <div id="auth-inv-count" style="font-size:18px;font-weight:800;color:var(--gold)">—</div>
            <div style="font-size:10px;color:var(--ts)">فاتورة</div>
          </div>
          <div style="text-align:center;padding:8px;background:rgba(46,204,113,0.06);border-radius:8px;border:1px solid rgba(46,204,113,0.15)">
            <div id="auth-entry-count" style="font-size:18px;font-weight:800;color:var(--ok)">—</div>
            <div style="font-size:10px;color:var(--ts)">قيد</div>
          </div>
        </div>
      </div>
    </div>

    <div class="auth-center">
      <div class="auth-emblem">
        <div class="emb-outer"></div><div class="emb-inner"></div>
        <span class="emb-crown">👑</span>
        <span class="emb-ll">🌿</span><span class="emb-lr">🌿</span>
        <div class="emb-body">
          <div class="emb-ar">شرق</div>
          <div class="emb-ar" style="font-size:17px">الإحتراف</div>
          <div class="emb-sub">١٤٤٥</div>
        </div>
      </div>
      <div class="auth-sys-title">نظام <span>شرق الاحتراف</span><br>المحاسبي الذكي الشامل</div>
      <div class="auth-sys-sub">FinBot OCR & Die-Casting Solutions — v3.1</div>

      <div class="auth-login-box" id="main-login-box">
        <div style="background:rgba(231,76,60,0.08);border:1px solid rgba(231,76,60,0.3);border-radius:10px;padding:8px 14px;margin-bottom:14px;display:flex;align-items:center;gap:8px;justify-content:center">
          <span style="font-size:14px">🔒</span>
          <span style="font-size:11px;color:var(--err);font-weight:700">النظام محمي — بدعوة رسمية فقط</span>
        </div>

        <button class="auth-enter-btn" onclick="initAdminLogin()" style="margin-bottom:16px">
          👑 دخول المدير العام
        </button>

        <div class="auth-divider">طلب دخول للموظفين</div>

        <div style="background:rgba(201,168,76,0.06);border:1px solid rgba(201,168,76,0.2);border-radius:10px;padding:8px 12px;margin-bottom:12px;text-align:center">
          <div style="font-size:10px;color:var(--ts);margin-bottom:3px">الطلب يُرسل مباشرة للمدير</div>
          <div style="font-size:12px;font-weight:700;color:var(--gold)">📧 arido0409@gmail.com</div>
        </div>

        <input class="req-input" type="text"  id="req-name"  placeholder="الاسم الكامل">
        <input class="req-input" type="email" id="req-email" placeholder="البريد الإلكتروني">
        <select id="req-role" style="background:#1a1200;border:1px solid var(--bord);border-radius:10px;padding:10px 14px;color:var(--ts);font-family:inherit;font-size:12px;outline:none;width:100%;margin-bottom:10px">
          <option value="accountant">🧾 محاسب</option>
          <option value="manager">📊 محاسب رئيسي</option>
          <option value="employee">👤 موظف</option>
        </select>
        <button onclick="requestAccess()" style="width:100%;padding:11px;background:rgba(201,168,76,0.08);border:1px solid var(--bord);border-radius:10px;color:var(--gold);font-family:'Tajawal',sans-serif;font-size:13px;font-weight:700;cursor:pointer;transition:all .2s;margin-bottom:12px" onmouseover="this.style.background='rgba(201,168,76,0.18)'" onmouseout="this.style.background='rgba(201,168,76,0.08)'">
          📩 إرسال طلب الدخول
        </button>
        <div class="auth-terms">الدخول بدون دعوة رسمية مرفوض<br>سيتم الرد على طلبك من <a href="#">المدير العام</a></div>
        <div id="auth-msg" style="min-height:18px;margin-top:8px;font-size:11px;text-align:center"></div>
      </div>

      <div class="auth-login-box" id="otp-box" style="display:none;width:340px">
        <div style="font-size:38px;margin-bottom:8px">🔐</div>
        <div style="font-size:16px;font-weight:800;color:var(--gold);margin-bottom:6px">كود التحقق</div>
        <div style="font-size:12px;color:var(--ts);margin-bottom:18px;line-height:1.7">
          أُرسل كود التحقق إلى<br>
          <strong style="color:var(--gold)">arido0409@gmail.com</strong><br>
          <span style="font-size:10px;color:var(--tm)">+ واتساب +966502283783</span>
        </div>
        <div style="display:flex;gap:8px;justify-content:center;margin-bottom:16px" id="otp-inputs-row">
          <input class="otp-digit" type="text" maxlength="1" id="d0" oninput="otpIn(0)" onkeydown="otpKey(event,0)">
          <input class="otp-digit" type="text" maxlength="1" id="d1" oninput="otpIn(1)" onkeydown="otpKey(event,1)">
          <input class="otp-digit" type="text" maxlength="1" id="d2" oninput="otpIn(2)" onkeydown="otpKey(event,2)">
          <input class="otp-digit" type="text" maxlength="1" id="d3" oninput="otpIn(3)" onkeydown="otpKey(event,3)">
          <input class="otp-digit" type="text" maxlength="1" id="d4" oninput="otpIn(4)" onkeydown="otpKey(event,4)">
          <input class="otp-digit" type="text" maxlength="1" id="d5" oninput="otpIn(5)" onkeydown="otpKey(event,5)">
        </div>
        <div style="font-size:12px;color:var(--ts);margin-bottom:14px;text-align:center">
          ⏱ الكود صالح: <strong id="otp-timer" style="color:var(--gold)">5:00</strong>
        </div>
        <button id="otp-btn" onclick="verifyOTP()" disabled style="width:100%;padding:13px;background:linear-gradient(135deg,var(--gold-dark),var(--gold));border:none;border-radius:12px;color:#000;font-family:'Tajawal',sans-serif;font-size:14px;font-weight:800;cursor:pointer;margin-bottom:10px;opacity:.4;transition:all .3s">
          ✅ تحقق ودخول
        </button>
        <div style="display:flex;gap:8px;justify-content:center">
          <button onclick="resendOTP()" style="padding:7px 18px;background:transparent;border:1px solid var(--bord);border-radius:10px;color:var(--ts);font-family:'Tajawal',sans-serif;font-size:11px;cursor:pointer;transition:all .2s" onmouseover="this.style.color='var(--gold)';this.style.borderColor='var(--gold)'" onmouseout="this.style.color='var(--ts)';this.style.borderColor='var(--bord)'">🔄 إعادة إرسال</button>
          <button onclick="cancelOTP()" style="padding:7px 18px;background:transparent;border:1px solid var(--bord);border-radius:10px;color:var(--ts);font-family:'Tajawal',sans-serif;font-size:11px;cursor:pointer;transition:all .2s" onmouseover="this.style.color='var(--err)'" onmouseout="this.style.color='var(--ts)'">← رجوع</button>
        </div>
        <div id="otp-err" style="color:var(--err);font-size:12px;margin-top:10px;text-align:center;min-height:18px"></div>
      </div>
    </div>

    <div class="auth-panel-right">
      <div class="auth-right-card"><div class="auth-right-card-ico">📊</div><div><div class="auth-right-card-label">التقارير الشهرية</div><div class="auth-right-card-sub">يومي • أسبوعي • شهري</div></div></div>
      <div class="auth-right-card"><div class="auth-right-card-ico">✏️</div><div><div class="auth-right-card-label">تعديل البيانات</div><div class="auth-right-card-sub">قيود • فواتير • سندات</div></div></div>
      <div class="auth-right-card"><div class="auth-right-card-ico">🧾</div><div><div class="auth-right-card-label">الفواتير والسندات</div><div class="auth-right-card-sub">قبض • صرف • مبيعات</div></div></div>
      <div class="auth-right-card"><div class="auth-right-card-ico">🤖</div><div><div class="auth-right-card-label">ريبورت AI</div><div class="auth-right-card-sub">تحليل ذكي • أوامر صوتية</div></div></div>
    </div>
  </div>

  <div class="auth-footer">
    <div class="auth-footer-logo">شرق الاحتراف © 2026</div>
    <div class="auth-footer-links">
      <span class="auth-footer-link">الرياض - الشفاء</span>
      <span class="auth-footer-link">0502283783</span>
      <span class="auth-footer-link">arido0409@gmail.com</span>
    </div>
    <div class="lang-sw"><div class="lang-dot"></div><span>العربية</span></div>
  </div>
</div>

<div class="udrop" id="udrop">
  <div class="udrop-head">
    <div class="uavatar" id="udrop-av" style="width:40px;height:40px;font-size:16px"></div>
    <div><div class="udrop-name" id="udrop-name">—</div><div class="udrop-role" id="udrop-role">—</div></div>
  </div>
  <div class="udrop-item" onclick="SP('tasks');closeUD()">📋 مهامي</div>
  <div class="udrop-item" onclick="SP('team');closeUD()" id="nav-team-drop">👥 إدارة الفريق</div>
  <div class="udrop-item" onclick="SP('approvals');closeUD()" id="nav-appr-drop">✅ الموافقات</div>
  <div class="udrop-item" onclick="openAccessRequests();closeUD()">📩 طلبات الدخول <span id="req-count-badge" style="background:rgba(231,76,60,0.2);color:var(--err);border-radius:10px;padding:1px 7px;font-size:10px;margin-right:4px"></span></div>
  <div class="udrop-item" onclick="SP('schedule');closeUD()">⏰ جدول الإرسال</div>
  <div class="udrop-item" onclick="doSignOut()">🚪 تسجيل الخروج</div>
</div>

<div id="app">
<header class="hdr">
  <div class="hdr-in">
    <div class="hlogo">
      <div class="lemb">ش</div>
      <div><div class="ltm">شرق الاحتراف</div><div class="lts">النظام المحاسبي الذكي <span style="background:rgba(201,168,76,0.2);border:1px solid rgba(201,168,76,0.4);border-radius:10px;padding:1px 7px;font-size:9px;color:var(--gold);margin-right:4px">v3.1</span></div></div>
    </div>
    <nav class="hnav">
      <button class="nb on" onclick="SP('dashboard')">📊 لوحة التحكم</button>
      <button class="nb" onclick="SP('add-entry')">➕ إضافة قيد</button>
      <button class="nb" onclick="SP('ocr')">📷 قراءة ذكية</button>
      <button class="nb" onclick="SP('transactions')">📋 السجلات</button>
      <button class="nb" onclick="SP('reports')">📈 التقارير</button>
      <button class="nb" onclick="SP('chat')">🤖 ريبورت AI</button>
      <button class="nb" onclick="SP('schedule')">⏰ جدول إرسال</button>
      <button class="nb" id="nav-tasks-btn" onclick="SP('tasks')">📋 المهام</button>
      <button class="nb" id="nav-team-btn"  onclick="SP('team')"  style="display:none">👥 الفريق</button>
      <button class="nb" id="nav-appr-btn"  onclick="SP('approvals')" style="display:none">✅ الموافقات</button>
      <button class="nb" onclick="SP('wabot')">💬 بوت واتساب</button>
      <button class="nb" onclick="SP('invoices')">🧾 الفواتير</button>
      <button class="nb" onclick="SP('receipts')">📄 سندات القبض</button>
      <button class="nb" onclick="SP('payments')">📝 سندات الصرف</button>
      <button class="nb" onclick="SP('comsettings')">⚙️ إعدادات</button>
    </nav>
    <div class="hacts">
      <button class="voice-btn" id="hdr-voice" onclick="toggleVoice()" title="صوت">🎤</button>
      <button class="bg" onclick="openEM()">📧 إرسال تقرير</button>
      <div class="uavatar" id="hdr-avatar" onclick="toggleUD()" title="حسابي" style="display:none"></div>
    </div>
  </div>
</header>

<main class="main">
<div id="page-dashboard" class="page on">
  <div class="pbanner">
    <div>
      <div class="plbl">صافي الربح الإجمالي</div>
      <div class="pval" id="net-profit">0 ريال</div>
      <div class="psub">اليوم — <span id="today-label"></span></div>
    </div>
    <div style="text-align:left">
      <div style="font-size:12px;color:var(--ts);margin-bottom:4px">آخر تحديث</div>
      <div style="font-size:17px;font-weight:700;color:var(--gold)" id="last-update">—</div>
    </div>
  </div>
  <div class="kgrid">
    <div class="kcard s"><div class="kico">💰</div><div class="klbl">إجمالي المبيعات</div><div class="kval" id="kts">0</div><div class="ksub">ريال سعودي</div></div>
    <div class="kcard p"><div class="kico">🛒</div><div class="klbl">إجمالي المشتريات</div><div class="kval" id="ktp">0</div><div class="ksub">ريال سعودي</div></div>
    <div class="kcard w"><div class="kico">💳</div><div class="klbl">المسحوبات والرواتب</div><div class="kval" id="ktw">0</div><div class="ksub">ريال سعودي</div></div>
    <div class="kcard e"><div class="kico">⚙️</div><div class="klbl">المصروفات التشغيلية</div><div class="kval" id="kte">0</div><div class="ksub">ريال سعودي</div></div>
  </div>
  <div class="card">
    <div class="ch"><div><div class="ct">آخر العمليات</div><div class="cst">أحدث القيود المحاسبية</div></div><button class="btn bto" onclick="SP('transactions')">عرض الكل</button></div>
    <div class="tw"><table><thead><tr><th>التاريخ</th><th>اليوم</th><th>النوع</th><th>المحول</th><th>المستفيد</th><th>المبلغ</th><th>التصنيف</th></tr></thead><tbody id="dtable"><tr><td colspan="7" style="text-align:center;color:var(--tm);padding:20px">لا توجد بيانات</td></tr></tbody></table></div>
  </div>
</div>
<div id="page-add-entry" class="page">
  <div class="card">
    <div class="ch"><div><div class="ct">إضافة قيد يدوي</div></div></div>
    <div class="tabs">
      <button class="tb on" onclick="setTab('s')">💰 مبيعات</button>
      <button class="tb" onclick="setTab('p')">🛒 مشتريات</button>
      <button class="tb" onclick="setTab('w')">💳 مسحوبات / رواتب</button>
      <button class="tb" onclick="setTab('e')">⚙️ مصروفات</button>
    </div>
    <div class="fg">
      <div class="fi"><label>التاريخ</label><input type="date" id="ed"></div>
      <div class="fi"><label>اسم المحول</label><input type="text" id="esend" placeholder="اسم الشخص أو الجهة"></div>
      <div class="fi"><label>اسم المستفيد</label><input type="text" id="erec" placeholder="اسم المستفيد"></div>
      <div class="fi"><label>المبلغ (ريال)</label><input type="number" id="eamt" placeholder="0.00" min="0" step="0.01"></div>
      <div class="fi"><label>التصنيف</label><select id="ecat"><option value="">— اختر —</option></select></div>
      <div class="fi"><label>طريقة الدفع</label><select id="epay"><option value="تحويل بنكي">تحويل بنكي</option><option value="نقدي">نقدي</option><option value="شيك">شيك</option><option value="بطاقة">بطاقة</option></select></div>
      <div class="fi full"><label>ملاحظات</label><textarea id="enotes" placeholder="تفاصيل إضافية..."></textarea></div>
    </div>
    <div class="br"><button class="btn btp" onclick="saveEntry()">💾 حفظ القيد</button><button class="btn bto" onclick="clearForm()">🗑️ مسح</button></div>
  </div>
</div>
<div id="page-ocr" class="page">
  <div class="card">
    <div class="ch"><div><div class="ct">📷 القراءة الذكية</div><div class="cst">تحليل رسائل البنك • PDF • الإイصالات • Excel • الصور</div></div></div>
    <div class="uzone" id="uzone" onclick="document.getElementById('fi2').click()">
      <div class="uico">📄</div>
      <div class="utxt"><strong>اسحب وأفلت الملفات هنا أو انقر للرفع</strong><br>
        <div style="margin-top:8px;display:flex;gap:6px;justify-content:center;flex-wrap:wrap">
          <span class="ftag">📄 PDF</span><span class="ftag">🖼️ صور</span><span class="ftag">📊 Excel</span><span class="ftag">📝 نصوص</span>
        </div></div>
      <input type="file" id="fi2" style="display:none" accept=".pdf,.jpg,.jpeg,.png,.txt,.xlsx,.xls,.csv" multiple onchange="handleFiles(this.files)">
    </div>
    <div id="fprev-list"></div>
    <div id="ocrload" class="ld" style="display:none"><div class="sp"></div><span>جاري التحليل...</span></div>
    <div id="ocrres" class="ares"><div style="font-size:13px;font-weight:700;color:var(--gold);margin-bottom:10px">نتائج التحليل</div><div id="ocritems"></div><div class="br"><button class="btn btp" onclick="confirmOCR()">✅ حفظ تلقائياً</button><button class="btn bto" onclick="editOCR()">✏️ تعديل</button></div></div>
    <div style="margin-top:18px">
      <div style="font-size:12px;color:var(--ts);margin-bottom:10px;font-weight:600">أو أدخل نص رسالة البنك / الفاتورة:</div>
      <textarea id="bmsg" placeholder="الصق النص هنا..." style="min-height:100px"></textarea>
      <div class="br"><button class="btn btp" onclick="analyzeMsg()">🔍 تحليل ذكي</button><button class="btn bto" onclick="document.getElementById('bmsg').value=''">مسح</button></div>
    </div>
  </div>
</div>
<div id="page-transactions" class="page">
  <div class="card">
    <div class="ch">
      <div><div class="ct">سجل العمليات</div></div>
      <div style="display:flex;gap:7px;flex-wrap:wrap">
        <select id="ftype" onchange="renderT()" style="width:auto;min-width:120px"><option value="">كل الأنواع</option><option>مبيعات</option><option>مشتريات</option><option>مسحوبات</option><option>مصروفات</option></select>
        <input type="date" id="fdf" onchange="renderT()" style="width:140px">
        <input type="date" id="fdt" onchange="renderT()" style="width:140px">
        <button class="btn bto" onclick="['ftype','fdf','fdt'].forEach(x=>document.getElementById(x).value='');renderT()">مسح</button>
      </div>
    </div>
    <div class="bulk-toolbar" id="bulk-bar">
      <input type="checkbox" class="bulk-check" onchange="toggleAll(this)"><span id="bulk-cnt" style="font-size:12px;color:var(--gold);font-weight:700">0 محدد</span>
      <button onclick="bulkDel()" style="padding:5px 12px;background:rgba(231,76,60,0.1);border:1px solid rgba(231,76,60,0.3);border-radius:20px;color:var(--err);cursor:pointer;font-size:11px;font-family:inherit;font-weight:600">🗑️ حذف</button>
      <button onclick="bulkExport()" style="padding:5px 12px;background:rgba(201,168,76,0.1);border:1px solid var(--bord);border-radius:20px;color:var(--gold);cursor:pointer;font-size:11px;font-family:inherit;font-weight:600">📥 تصدير</button>
      <button onclick="clearBulk()" style="padding:5px 12px;background:transparent;border:1px solid var(--bord);border-radius:20px;color:var(--ts);cursor:pointer;font-size:11px;font-family:inherit">✕</button>
    </div>
    <div class="tw"><table><thead><tr><th><input type="checkbox" class="bulk-check" onchange="toggleAll(this)"></th><th>#</th><th>التاريخ</th><th>اليوم</th><th>النوع</th><th>المحول</th><th>المستفيد</th><th>المبلغ</th><th>التصنيف</th><th>الدفع</th><th>ملاحظات</th><th>إجراءات</th></tr></thead><tbody id="ttable"></tbody></table></div>
  </div>
</div>
<div id="page-reports" class="page">
  <div class="card">
    <div class="ch">
      <div><div class="ct">📈 التقارير المالية</div></div>
      <div style="display:flex;gap:7px;flex-wrap:wrap">
        <button class="btn bto" onclick="expPDF()">📄 PDF</button>
        <button class="btn bto" onclick="expXLS()">📊 Excel</button>
        <button class="btn bto" onclick="window.print()">🖨️ طباعة</button>
      </div>
    </div>
    <div class="rctrl">
      <select id="rperiod" onchange="genReport()"><option value="daily">يومي</option><option value="weekly">أسبوعي</option><option value="monthly">شهري</option><option value="custom">مخصص</option></select>
      <input type="date" id="rfrom" onchange="genReport()" style="width:150px">
      <input type="date" id="rto"   onchange="genReport()" style="width:150px">
    </div>
    <div id="rcontent"><div style="text-align:center;padding:36px;color:var(--tm)">اختر فترة التقرير</div></div>
  </div>
</div>
<div id="page-chat" class="page">
  <div class="cwrap">
    <div class="cmain">
      <div class="ch" style="padding:14px 18px;margin:0;border-radius:0;border:none;border-bottom:1px solid var(--bord)">
        <div style="display:flex;align-items:center;gap:10px">
          <div style="width:36px;height:36px;border-radius:50%;background:linear-gradient(135deg,var(--gold-dark),var(--gold));display:flex;align-items:center;justify-content:center;font-size:16px">🤖</div>
          <div><div class="ct">ريبورت AI</div><div class="cst">مساعد مالي ذكي — صلاحيات كاملة</div></div>
        </div>
        <button class="voice-btn" id="chat-vbtn" onclick="toggleVoice()">🎤</button>
      </div>
      <div class="wave-wrap" id="wave-wrap">
        <div class="wave-lbl" style="font-size:12px;color:var(--err);margin-right:8px">🎤 جاري الاستماع...</div>
        <div class="wave-bar"></div><div class="wave-bar"></div><div class="wave-bar"></div><div class="wave-bar"></div><div class="wave-bar"></div>
      </div>
      <div class="cmsgs" id="cmsgs">
        <div class="msg mb"><div class="mbub">مرحباً! أنا <strong style="color:var(--gold)">ريبورت</strong> 🤖<br><br>📄 أحلل الملفات والفواتير<br>➕ أضيف مبيعات • مشتريات • مصروفات • رواتب<br>📊 أُصدر تقارير فورية<br>📧 أُرسل للمدير بريد وواتساب<br>🎤 أفهم الأوامر الصوتية</div><div class="mmeta">ريبورت • الآن</div></div>
      </div>
      <div id="chat-fz" style="margin:0 14px 8px;border:1px dashed var(--bord);border-radius:var(--rs);padding:9px 14px;text-align:center;cursor:pointer;font-size:12px;color:var(--ts);transition:all .2s" onclick="document.getElementById('chat-fi').click()">📎 ارفع ملف PDF • صورة • Excel
        <input type="file" id="chat-fi" style="display:none" accept=".pdf,.jpg,.jpeg,.png,.txt,.xlsx,.xls,.csv" multiple onchange="chatFile(this.files)">
      </div>
      <div class="cinput-row">
        <textarea id="cinput" placeholder="اكتب أمراً..." onkeydown="ck(event)"></textarea>
        <button class="voice-btn" onclick="toggleVoice()">🎤</button>
        <button class="btn btp" onclick="sendChat()" style="height:42px;padding:0 16px">إرسال ↗</button>
      </div>
    </div>
    <div class="csidebar">
      <div class="cside-card">
        <div class="cside-title">⚡ أوامر سريعة</div>
        <button class="qbtn" onclick="qc('أضف مبيعات 5000 ريال من مساحات دافئة اليوم')">💰 إضافة مبيعات</button>
        <button class="qbtn" onclick="qc('أضف مشتريات 3000 ريال من المورد')">🛒 إضافة مشتريات</button>
        <button class="qbtn" onclick="qc('صرفنا راتب محمد 4500 ريال')">💳 تسجيل راتب</button>
        <button class="qbtn" onclick="qc('أضف مصروفات كهرباء 800 ريال')">⚙️ إضافة مصروفات</button>
        <button class="qbtn" onclick="qc('أعطني تقرير اليوم الكامل')">📊 تقرير اليوم</button>
        <button class="qbtn" onclick="qc('تقرير هذا الشهر مع التحليل')">📆 تقرير الشهر</button>
        <button class="qbtn" onclick="qc('أرسل تقرير اليوم للمدير على البريد arido0409@gmail.com')">📧 إرسال بريد</button>
        <button class="qbtn" onclick="sendRptWA('daily')">💬 إرسال واتساب للمدير</button>
        <button class="qbtn" onclick="toggleVoice()">🎤 تشغيل الصوت</button>
      </div>
      <div class="cside-card">
        <div class="cside-title">📧 إعدادات الإرسال</div>
        <div class="fi" style="margin-bottom:8px"><label>بريد المدير</label><input type="email" id="mgr-email" value="arido0409@gmail.com"></div>
        <div class="fi" style="margin-bottom:8px"><label>واتساب المدير</label><input type="text" id="mgr-wa" value="+966502283783"></div>
        <div style="display:flex;gap:5px;margin-bottom:10px">
          <button class="tb on" id="sch-daily"   onclick="setSched('daily')"   style="flex:1;font-size:11px;padding:5px">يومي</button>
          <button class="tb"   id="sch-weekly"  onclick="setSched('weekly')"  style="flex:1;font-size:11px;padding:5px">أسبوعي</button>
          <button class="tb"   id="sch-monthly" onclick="setSched('monthly')" style="flex:1;font-size:11px;padding:5px">شهري</button>
        </div>
        <div class="fi" style="margin-bottom:10px"><label>وقت الإرسال</label><input type="time" id="sched-time" value="08:00"></div>
        <button class="btn btp" onclick="saveSchedSide()" style="width:100%;font-size:12px">💾 حفظ وتفعيل</button>
      </div>
    </div>
  </div>
</div>
<div id="page-schedule" class="page">
  <div class="card">
    <div class="ch"><div><div class="ct">⏰ جدول الإرسال التلقائي</div></div><div id="sched-badge"></div></div>
    <div class="sgrid">
      <div class="scard on" id="scard-daily"   onclick="setSched('daily')"><div class="scard-ico">📅</div><div class="scard-lbl">يومي</div></div>
      <div class="scard"    id="scard-weekly"  onclick="setSched('weekly')"><div class="scard-ico">📆</div><div class="scard-lbl">أسبوعي</div></div>
      <div class="scard"    id="scard-monthly" onclick="setSched('monthly')"><div class="scard-ico">🗓️</div><div class="scard-lbl">شهري</div></div>
    </div>
    <div class="fg" style="margin-bottom:16px">
      <div class="fi"><label>بريد المدير</label><input type="email" id="sp-email" value="arido0409@gmail.com"></div>
      <div class="fi"><label>واتساب المدير</label><input type="text" id="sp-wa" value="+966502283783"></div>
      <div class="fi"><label>وقت الإرسال</label><input type="time" id="sp-time" value="08:00"></div>
    </div>
    <div class="br">
      <button class="btn btp" onclick="saveSchedFull()">💾 حفظ وتفعيل</button>
      <button class="btn bto" onclick="sendNow('email')">📧 إرسال الآن بالبريد</button>
      <button class="btn bto" onclick="sendNow('wa')">💬 إرسال الآن بواتساب</button>
    </div>
    <div id="sched-log" style="margin-top:16px"></div>
  </div>
</div>
<div id="page-tasks" class="page">
  <div class="card">
    <div class="ch"><div><div class="ct">📋 إدارة المهام</div><div class="cst" id="tasks-sub">مهامي</div></div><button class="btn btp" id="add-task-btn" onclick="openTaskForm()" style="display:none">➕ مهمة جديدة</button></div>
    <div class="task-form" id="task-form" style="display:none;background:var(--card2);border:1px solid var(--bord);border-radius:var(--r);padding:16px;margin-bottom:16px">
      <div style="font-size:13px;font-weight:700;color:var(--gold);margin-bottom:12px">➕ مهمة جديدة</div>
      <div class="fg">
        <div class="fi half"><label>عنوان المهمة</label><input type="text" id="t-title" placeholder="عنوان المهمة"></div>
        <div class="fi"><label>الأولوية</label><select id="t-pri"><option value="high">🔴 عالية</option><option value="med" selected>🟡 متوسطة</option><option value="low">🟢 منخفضة</option></select></div>
        <div class="fi"><label>تعيين إلى</label><select id="t-assign"><option value="">— اختر —</option></select></div>
        <div class="fi"><label>الاستحقاق</label><input type="date" id="t-due"></div>
        <div class="fi"><label>القسم</label><select id="t-dept"><option>محاسبة</option><option>مبيعات</option><option>مشتريات</option><option>تشغيل</option><option>إدارة</option></select></div>
        <div class="fi full"><label>التفاصيل</label><textarea id="t-desc" style="min-height:50px"></textarea></div>
      </div>
      <div class="br"><button class="btn btp" onclick="saveTask()">💾 حفظ</button><button class="btn bto" onclick="document.getElementById('task-form').style.display='none'">إلغاء</button></div>
    </div>
    <div class="task-board">
      <div class="task-col todo" id="col-todo" ondragover="event.preventDefault()" ondrop="dropTask(event,'todo')"><div class="task-col-title">⏳ قيد التنفيذ <span id="cnt-todo" style="background:rgba(201,168,76,0.15);border-radius:10px;padding:1px 7px;font-size:11px">0</span></div><div id="tasks-todo"></div></div>
      <div class="task-col doing" id="col-doing" ondragover="event.preventDefault()" ondrop="dropTask(event,'doing')"><div class="task-col-title">🔄 جاري <span id="cnt-doing" style="background:rgba(243,156,18,0.15);border-radius:10px;padding:1px 7px;font-size:11px">0</span></div><div id="tasks-doing"></div></div>
      <div class="task-col done" id="col-done" ondragover="event.preventDefault()" ondrop="dropTask(event,'done')"><div class="task-col-title">✅ منجز <span id="cnt-done" style="background:rgba(46,204,113,0.1);border-radius:10px;padding:1px 7px;font-size:11px">0</span></div><div id="tasks-done"></div></div>
    </div>
  </div>
</div>
<div id="page-team" class="page"><div class="card"><div class="ch"><div><div class="ct">👥 إدارة الفريق</div></div><button class="btn btp" onclick="openInviteModal()">✉️ دعوة عضو</button></div><div class="team-grid" id="team-grid"></div></div></div>
<div id="page-approvals" class="page"><div class="card"><div class="ch"><div><div class="ct">✅ الموافقات</div></div><span id="appr-cnt" style="background:rgba(231,76,60,0.15);color:var(--err);border:1px solid rgba(231,76,60,0.3);border-radius:20px;padding:3px 12px;font-size:12px;font-weight:700"></span></div><div id="approvals-list"><div style="text-align:center;padding:36px;color:var(--tm)">✅ لا توجد موافقات معلقة</div></div></div></div>
<div id="page-wabot" class="page">
  <div style="display:grid;grid-template-columns:1fr 360px;gap:20px;align-items:start">
    <div>
      <div class="card" style="margin-bottom:16px">
        <div class="ch"><div><div class="ct">💬 بوت واتساب — ريبورت AI</div><div class="cst">أرسل أوامر من واتساب مباشرة</div></div>
          <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
            <div id="wa-status-badge" class="webhook-status inactive"><div class="pulse"></div><span id="wa-status-text">غير متصل</span></div>
          </div>
        </div>
        <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:14px">
          <div style="background:var(--card2);border:1px solid var(--bord);border-radius:var(--rs);padding:12px;text-align:center"><div style="font-size:20px;margin-bottom:3px">💬</div><div style="font-size:18px;font-weight:800;color:var(--gold)" id="wa-mc">0</div><div style="font-size:11px;color:var(--ts)">رسائل</div></div>
          <div style="background:var(--card2);border:1px solid var(--bord);border-radius:var(--rs);padding:12px;text-align:center"><div style="font-size:20px;margin-bottom:3px">✅</div><div style="font-size:18px;font-weight:800;color:var(--ok)" id="wa-cc">0</div><div style="font-size:11px;color:var(--ts)">أوامر</div></div>
          <div style="background:var(--card2);border:1px solid var(--bord);border-radius:var(--rs);padding:12px;text-align:center"><div style="font-size:20px;margin-bottom:3px">📊</div><div style="font-size:18px;font-weight:800;color:var(--inf)" id="wa-rc">0</div><div style="font-size:11px;color:var(--ts)">تقارير</div></div>
        </div>
        <div class="tabs" style="margin-bottom:14px">
          <button class="tb on" onclick="waTab('setup')">⚙️ الإعداد</button>
          <button class="tb" onclick="waTab('commands')">📋 الأوامر</button>
          <button class="tb" onclick="waTab('history')">📜 السجل</button>
          <button class="tb" onclick="waTab('webhook')">🔗 Webhook</button>
        </div>
        <div id="wa-tab-setup">
          <div class="sgrid"><div class="scard on" id="method-twilio" onclick="setWAM('twilio')"><div class="scard-ico">📱</div><div class="scard-lbl">Twilio</div></div><div class="scard" id="method-360" onclick="setWAM('360')"><div class="scard-ico">🔷</div><div class="scard-lbl">360dialog</div></div><div class="scard" id="method-baileys" onclick="setWAM('baileys')"><div class="scard-ico">⚡</div><div class="scard-lbl">Baileys</div></div></div>
          <div id="wa-method-content"></div>
          <div class="fg" style="grid-template-columns:1fr 1fr;margin-top:14px">
            <div class="fi"><label>رقم المدير</label><input type="text" id="wa-admin-num" value="+966502283783"></div>
            <div class="fi"><label>API Key</label><input type="text" id="wa-api-key" placeholder="ACxxxxxxxx"></div>
          </div>
          <div class="br"><button class="btn btp" onclick="saveWASet()">💾 حفظ</button><button class="btn bto" onclick="testWA()">🔍 اختبار</button></div>
        </div>
        <div id="wa-tab-commands" style="display:none">
          <div class="cmd-grid">
            <div class="cmd-card" onclick="simWA('أضف مبيعات 5000 ريال من مساحات دافئة')"><div class="cmd-icon">💰</div><div class="cmd-name">إضافة مبيعات</div><div class="cmd-example">أضف مبيعات 5000 من ...</div></div>
            <div class="cmd-card" onclick="simWA('أضف مشتريات 3000 من المورد')"><div class="cmd-icon">🛒</div><div class="cmd-name">إضافة مشتريات</div><div class="cmd-example">أضف مشتريات 3000 من ...</div></div>
            <div class="cmd-card" onclick="simWA('صرفنا راتب محمد 4500 ريال')"><div class="cmd-icon">💳</div><div class="cmd-name">تسجيل راتب</div><div class="cmd-example">صرفنا راتب محمد 4500</div></div>
            <div class="cmd-card" onclick="simWA('أضف مصروفات إيجار 8000 ريال')"><div class="cmd-icon">⚙️</div><div class="cmd-name">إضافة مصروفات</div><div class="cmd-example">أضف مصروفات إيجار 8000</div></div>
            <div class="cmd-card" onclick="simWA('تقرير اليوم')"><div class="cmd-icon">📊</div><div class="cmd-name">تقرير اليوم</div><div class="cmd-example">تقرير اليوم</div></div>
            <div class="cmd-card" onclick="simWA('تقرير الشهر')"><div class="cmd-icon">📆</div><div class="cmd-name">تقرير الشهر</div><div class="cmd-example">تقرير الشهر</div></div>
            <div class="cmd-card" onclick="simWA('صافي الربح')"><div class="cmd-icon">💰</div><div class="cmd-name">صافي الربح</div><div class="cmd-example">صافي الربح</div></div>
            <div class="cmd-card" onclick="simWA('مساعدة')"><div class="cmd-icon">❓</div><div class="cmd-name">مساعدة</div><div class="cmd-example">مساعدة</div></div>
          </div>
        </div>
        <div id="wa-tab-history" style="display:none"><div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px"><div style="font-size:13px;font-weight:700;color:var(--gold)">📜 سجل الأوامر</div><button onclick="clearWAHist()" style="padding:3px 10px;font-size:11px;background:rgba(231,76,60,0.08);border:1px solid rgba(231,76,60,0.3);border-radius:20px;color:var(--err);cursor:pointer">مسح</button></div><div id="wa-hist-list"><div style="text-align:center;padding:24px;color:var(--tm);font-size:12px">لا توجد أوامر بعد</div></div></div>
        <div id="wa-tab-webhook" style="display:none">

          <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:16px">
            <div id="wh-meta" onclick="setWHMethod('meta')" style="padding:12px;background:rgba(201,168,76,0.08);border:2px solid var(--gold);border-radius:10px;cursor:pointer;text-align:center">
              <div style="font-size:20px;margin-bottom:4px">📘</div>
              <div style="font-size:12px;font-weight:700;color:var(--gold)">Meta WhatsApp API</div>
              <div style="font-size:10px;color:var(--ts)">رسمي — موصى به</div>
            </div>
            <div id="wh-baileys" onclick="setWHMethod('baileys')" style="padding:12px;background:rgba(201,168,76,0.04);border:1px solid var(--bord);border-radius:10px;cursor:pointer;text-align:center">
              <div style="font-size:20px;margin-bottom:4px">⚡</div>
              <div style="font-size:12px;font-weight:700;color:var(--ts)">Baileys (مجاني)</div>
              <div style="font-size:10px;color:var(--tm)">Node.js بدون API</div>
            </div>
          </div>

          <div id="wh-meta-content">
            <div style="background:rgba(24,119,242,0.08);border:1px solid rgba(24,119,242,0.3);border-radius:10px;padding:12px 14px;margin-bottom:16px;display:flex;align-items:center;gap:10px">
              <span style="font-size:22px">📘</span>
              <div>
                <div style="font-size:12px;font-weight:700;color:#4b9eff">Meta WhatsApp Business API</div>
                <div style="font-size:11px;color:var(--ts)">الطريقة الرسمية — تدعم الإرسال التلقائي للمدير</div>
              </div>
            </div>

            <div class="setup-step">
              <div class="step-num">1</div>
              <div class="step-content">
                <div class="step-title">إعداد Meta Business</div>
                <div class="step-desc">اذهب إلى <strong style="color:var(--gold)">developers.facebook.com</strong><br>أنشئ تطبيقاً جديداً ← Business ← WhatsApp ← احصل على Access Token ورقم Phone Number ID</div>
              </div>
            </div>

            <div class="setup-step">
              <div class="step-num">2</div>
              <div class="step-content">
                <div class="step-title">ملف الإرسال PHP — <code style="color:var(--ok);font-size:11px">send_wa.php</code></div>
                <div class="step-desc">الصق هذا الكود في سيرفرك (Hostinger, cPanel, أي استضافة PHP)</div>
                <div class="code-block">&lt;?php
// ========== إعدادات شرق الاحتراف ==========
$token          = "YOUR_META_ACCESS_TOKEN";   // من Meta Developers
$phone_number_id= "YOUR_PHONE_NUMBER_ID";     // من Meta Developers
$admin_number   = "966502283783";              // رقمك بدون +
$anthropic_key  = "YOUR_ANTHROPIC_KEY";        // من console.anthropic.com

// ========== استقبال الأمر ==========
$input   = json_decode(file_get_contents('php://input'), true);
$command = $input['command'] ?? $_GET['cmd'] ?? 'تقرير اليوم';

// ========== توليد الرد بـ AI ==========
$ai_response = callReportBot($command, $anthropic_key);

// ========== إرسال عبر واتساب ==========
$result = sendWhatsApp($token, $phone_number_id, $admin_number, $ai_response);
echo json_encode(['status' => 'ok', 'sent' => $result]);

// ========== دالة ريبورت AI ==========
function callReportBot($command, $key) {
    $ch = curl_init("https://api.anthropic.com/v1/messages");
    curl_setopt_array($ch, [
        CURLOPT_POST        => true,
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_HTTPHEADER  => [
            "x-api-key: $key",
            "anthropic-version: 2023-06-01",
            "Content-Type: application/json"
        ],
        CURLOPT_POSTFIELDS  => json_encode([
            "model"      => "claude-sonnet-4-20250514",
            "max_tokens" => 500,
            "system"     => "أنت ريبورت، مساعد مالي ذكي لشرق الاحتراف. ردود قصيرة للواتساب.",
            "messages"   => [["role"=>"user","content"=>$command]]
        ])
    ]);
    $res  = curl_exec($ch); curl_close($ch);
    $data = json_decode($res, true);
    return $data['content'][0]['text'] ?? 'حدث خطأ';
}

// ========== دالة إرسال واتساب ==========
function sendWhatsApp($token, $phone_id, $to, $message) {
    $url  = "https://graph.facebook.com/v19.0/$phone_id/messages";
    $data = [
        "messaging_product" => "whatsapp",
        "to"                => $to,
        "type"              => "text",
        "text"              => ["body" => $message]
    ];
    $ch = curl_init($url);
    curl_setopt_array($ch, [
        CURLOPT_POST           => true,
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_POSTFIELDS     => json_encode($data),
        CURLOPT_HTTPHEADER     => [
            "Authorization: Bearer $token",
            "Content-Type: application/json"
        ]
    ]);
    $res = curl_exec($ch); curl_close($ch);
    return json_decode($res, true);
}
?&gt;</div>
                <button onclick="copyPHP()" style="display:inline-block;margin-top:6px;padding:4px 14px;background:rgba(201,168,76,0.1);border:1px solid var(--bord);border-radius:20px;color:var(--gold);font-size:10px;cursor:pointer;font-family:inherit">📋 نسخ الكود</button>
              </div>
            </div>

            <div class="setup-step">
              <div class="step-num">3</div>
              <div class="step-content">
                <div class="step-title">Webhook لاستقبال الأوامر من المدير</div>
                <div class="step-desc">عندما يرسل المدير رسالة على واتساب تصل تلقائياً لهذا الـ Webhook</div>
                <div class="code-block">&lt;?php
// webhook.php — يستقبل رسائل المدير من واتساب

$verify_token = "sharq_secret_2026"; // اختر كلمة سر

if ($_SERVER['REQUEST_METHOD'] === 'GET') {
    if ($_GET['hub_verify_token'] === $verify_token) {
        echo $_GET['hub_challenge'];
        exit;
    }
    exit('Unauthorized');
}

$body   = json_decode(file_get_contents('php://input'), true);
$entry  = $body['entry'][0]['changes'][0]['value'] ?? null;
if (!$entry || !isset($entry['messages'])) exit;

$from   = $entry['messages'][0]['from'];      
$text   = $entry['messages'][0]['text']['body'] ?? '';
$admin  = "966502283783";

if ($from === $admin && $text) {
    $reply = callReportBot($text, "YOUR_ANTHROPIC_KEY");
    sendWhatsApp("YOUR_TOKEN", "YOUR_PHONE_ID", $from, $reply);
}
?&gt;</div>
              </div>
            </div>

            <div class="setup-step">
              <div class="step-num">4</div>
              <div class="step-content">
                <div class="step-title">ربط الـ Webhook بـ Meta</div>
                <div class="step-desc">في لوحة Meta Developers ← WhatsApp ← Configuration ← Webhook:<br>
                أدخل URL: <code style="color:var(--ok)">https://yoursite.com/webhook.php</code><br>
                Verify Token: <code style="color:var(--ok)">sharq_secret_2026</code></div>
              </div>
            </div>

            <div style="background:rgba(46,204,113,0.08);border:1px solid rgba(46,204,113,0.25);border-radius:10px;padding:14px;margin-top:6px">
              <div style="font-size:12px;font-weight:700;color:var(--ok);margin-bottom:8px">⚡ اختبار سريع — أرسل الآن</div>
              <div class="fg" style="grid-template-columns:1fr 1fr;gap:8px">
                <div class="fi"><label>رابط سيرفرك</label><input type="text" id="php-url" placeholder="https://yoursite.com/send_wa.php" style="font-size:11px"></div>
                <div class="fi"><label>أمر الاختبار</label><input type="text" id="php-cmd" value="تقرير اليوم" style="font-size:11px"></div>
              </div>
              <button onclick="testPHPSend()" style="margin-top:10px;padding:8px 18px;background:linear-gradient(135deg,var(--gold-dark),var(--gold));border:none;border-radius:var(--rs);color:#000;font-family:inherit;font-size:12px;font-weight:700;cursor:pointer">📤 إرسال اختبار</button>
              <div id="php-test-result" style="margin-top:8px;font-size:11px;color:var(--ts)"></div>
            </div>
          </div>

          <div id="wh-baileys-content" style="display:none">
            <div class="setup-step"><div class="step-num">1</div><div class="step-content"><div class="step-title">تثبيت Baileys</div><div class="code-block">npm init -y
npm install @whiskeysockets/baileys express node-fetch</div></div></div>
            <div class="setup-step"><div class="step-num">2</div><div class="step-content"><div class="step-title">متغيرات البيئة</div><div class="code-block">ANTHROPIC_KEY=sk-ant-xxxxxxxxxxxxxxxx
ADMIN_WA=966502283783
PORT=3000</div></div></div>
            <div class="setup-step"><div class="step-num">3</div><div class="step-content"><div class="step-title">مسح QR لربط الهاتف</div><div class="step-desc">شغّل <code style="color:var(--ok)">node server.js</code> وامسح الـ QR الظاهر من واتساب ← الأجهزة المرتبطة</div></div></div>
          </div>

        </div>
      </div>
    </div>
    <div style="position:sticky;top:88px">
      <div style="font-size:12px;color:var(--ts);text-align:center;margin-bottom:8px;font-weight:600">📱 محاكي واتساب</div>
      <div class="wa-phone-frame">
        <div class="wa-status-bar"><span>9:41</span><span>▓▓▓ 🔋</span></div>
        <div class="wa-header"><div class="wa-av">🤖</div><div><div class="wa-hname">ريبورت AI</div><div class="wa-hstatus" id="wa-sim-st">متصل ●</div></div></div>
        <div class="wa-body" id="wa-sim-body">
          <div style="text-align:center;font-size:10px;color:#4a4a4a;background:rgba(0,0,0,0.3);border-radius:10px;padding:4px 10px;align-self:center;margin:4px 0">اليوم</div>
          <div class="wa-msg-in">🤖 *ريبورت AI*<br>أهلاً! أرسل أمراً 👋<br><br>مثال:<br>• أضف مبيعات 5000 من أحمد<br>• صرفنا راتب سعد 3500<br>• تقرير اليوم<br><br>أو اكتب: <strong>مساعدة</strong><div class="wa-time">9:00 ص</div></div>
        </div>
        <div class="wa-input-bar">
          <input class="wa-input" id="wa-sim-in" placeholder="اكتب أمراً..." onkeydown="if(event.key==='Enter')sendWASim()">
          <button class="wa-send-btn" onclick="sendWASim()">➤</button>
        </div>
      </div>
      <div style="margin-top:10px;display:grid;grid-template-columns:1fr 1fr;gap:5px">
        <button class="qbtn" onclick="simWA('أضف مبيعات 3000 من مساحات')">💰 مبيعات</button>
        <button class="qbtn" onclick="simWA('صرفنا راتب علي 4000')">💳 راتب</button>
        <button class="qbtn" onclick="simWA('أضف مصروفات كهرباء 500')">⚙️ مصروفات</button>
        <button class="qbtn" onclick="simWA('تقرير اليوم')">📊 تقرير</button>
        <button class="qbtn" onclick="simWA('صافي الربح')">💰 الربح</button>
        <button class="qbtn" onclick="simWA('مساعدة')">❓ مساعدة</button>
        <button class="qbtn" onclick="openAdminWA()" style="grid-column:span 2;background:rgba(37,211,102,0.1);border-color:rgba(37,211,102,0.35);color:#25D366;font-weight:700">💬 فتح محادثة المدير</button>
      </div>
    </div>
  </div>
</div>
<div id="page-invoices" class="page"><div class="card"><div class="ch"><div><div class="ct">🧾 الفواتير</div><div class="cst">فواتير المبيعات الصادرة</div></div><div style="display:flex;gap:8px;align-items:center"><span class="counter-pill">آخر: <span id="inv-last-num">INV-0000</span></span><button class="btn btp" onclick="openNewDoc('invoice')">➕ فاتورة جديدة</button></div></div>
<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px">
  <div style="background:rgba(201,168,76,0.06);border:1px solid rgba(201,168,76,0.15);border-radius:var(--rs);padding:12px;text-align:center"><div style="font-size:11px;color:var(--ts)">الإجمالي</div><div style="font-size:16px;font-weight:800;color:var(--gold)" id="inv-total">0</div></div>
  <div style="background:rgba(46,204,113,0.06);border:1px solid rgba(46,204,113,0.15);border-radius:var(--rs);padding:12px;text-align:center"><div style="font-size:11px;color:var(--ts)">مدفوعة</div><div style="font-size:16px;font-weight:800;color:var(--ok)" id="inv-paid">0</div></div>
  <div style="background:rgba(231,76,60,0.06);border:1px solid rgba(231,76,60,0.15);border-radius:var(--rs);padding:12px;text-align:center"><div style="font-size:11px;color:var(--ts)">غير مدفوعة</div><div style="font-size:16px;font-weight:800;color:var(--err)" id="inv-unpaid">0</div></div>
  <div style="background:rgba(243,156,18,0.06);border:1px solid rgba(243,156,18,0.15);border-radius:var(--rs);padding:12px;text-align:center"><div style="font-size:11px;color:var(--ts)">جزئية</div><div style="font-size:16px;font-weight:800;color:var(--warn)" id="inv-partial">0</div></div>
</div>
<div id="invoices-list"><div style="text-align:center;padding:36px;color:var(--tm)">لا توجد فواتير</div></div></div></div>
<div id="page-receipts" class="page"><div class="card"><div class="ch"><div><div class="ct">📄 سندات القبض</div></div><div style="display:flex;gap:8px;align-items:center"><span class="counter-pill">آخر: <span id="rcpt-last-num">REC-0000</span></span><button class="btn btp" onclick="openNewDoc('receipt')">➕ سند قبض</button></div></div>
<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:14px">
  <div style="background:rgba(46,204,113,0.06);border:1px solid rgba(46,204,113,0.15);border-radius:var(--rs);padding:12px;text-align:center"><div style="font-size:11px;color:var(--ts)">الإجمالي</div><div style="font-size:16px;font-weight:800;color:var(--ok)" id="rcpt-total">0 ر.س</div></div>
  <div style="background:rgba(46,204,113,0.06);border:1px solid rgba(46,204,113,0.15);border-radius:var(--rs);padding:12px;text-align:center"><div style="font-size:11px;color:var(--ts)">العدد</div><div style="font-size:16px;font-weight:800;color:var(--ok)" id="rcpt-count">0</div></div>
  <div style="background:rgba(52,152,219,0.06);border:1px solid rgba(52,152,219,0.15);border-radius:var(--rs);padding:12px;text-align:center"><div style="font-size:11px;color:var(--ts)">هذا الشهر</div><div style="font-size:16px;font-weight:800;color:var(--inf)" id="rcpt-month">0 ر.س</div></div>
</div>
<div id="receipts-list"><div style="text-align:center;padding:36px;color:var(--tm)">لا توجد سندات</div></div></div></div>
<div id="page-payments" class="page"><div class="card"><div class="ch"><div><div class="ct">📝 سندات الصرف</div></div><div style="display:flex;gap:8px;align-items:center"><span class="counter-pill">آخر: <span id="pay-last-num">PAY-0000</span></span><button class="btn btp" onclick="openNewDoc('payment')">➕ سند صرف</button></div></div>
<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:14px">
  <div style="background:rgba(231,76,60,0.06);border:1px solid rgba(231,76,60,0.15);border-radius:var(--rs);padding:12px;text-align:center"><div style="font-size:11px;color:var(--ts)">الإجمالي</div><div style="font-size:16px;font-weight:800;color:var(--err)" id="pay-total">0 ر.س</div></div>
  <div style="background:rgba(231,76,60,0.06);border:1px solid rgba(231,76,60,0.15);border-radius:var(--rs);padding:12px;text-align:center"><div style="font-size:11px;color:var(--ts)">العدد</div><div style="font-size:16px;font-weight:800;color:var(--err)" id="pay-count">0</div></div>
  <div style="background:rgba(243,156,18,0.06);border:1px solid rgba(243,156,18,0.15);border-radius:var(--rs);padding:12px;text-align:center"><div style="font-size:11px;color:var(--ts)">هذا الشهر</div><div style="font-size:16px;font-weight:800;color:var(--warn)" id="pay-month">0 ر.س</div></div>
</div>
<div id="payments-list"><div style="text-align:center;padding:36px;color:var(--tm)">لا توجد سندات</div></div></div></div>
<div id="page-comsettings" class="page"><div class="card"><div class="ch"><div><div class="ct">⚙️ إعدادات الشركة</div></div><div style="display:flex;gap:8px"><button class="btn btp" onclick="saveCoSettings()">💾 حفظ</button><button class="btn" onclick="resetCoSettings()" style="background:rgba(231,76,60,0.08);border:1px solid rgba(231,76,60,0.3);color:var(--err)">🗑️ مسح</button></div></div>
<div style="display:grid;grid-template-columns:1fr 1fr;gap:20px">
  <div>
    <div style="font-size:13px;font-weight:700;color:var(--gold);margin-bottom:12px">الشعار والختم</div>
    <div style="display:flex;gap:18px;align-items:center;margin-bottom:14px">
      <div><div style="font-size:11px;color:var(--ts);margin-bottom:6px">الشعار</div><div class="logo-upload" id="logo-btn" onclick="document.getElementById('logo-fi').click()"><img id="logo-img" src="" style="display:none"><span id="logo-ph">🏢</span></div><input type="file" id="logo-fi" style="display:none" accept="image/*" onchange="uploadLogo(this)"></div>
      <div><div style="font-size:11px;color:var(--ts);margin-bottom:6px">الختم</div><div class="stamp-preview" id="stamp-el"><span id="stamp-pr" style="font-size:9px;text-align:center;padding:4px">ختم<br>الشركة</span></div></div>
    </div>
    <div class="fg" style="grid-template-columns:1fr">
      <div class="fi"><label>سطر الختم 1</label><input type="text" id="stamp-l1" oninput="updStamp()"></div>
      <div class="fi"><label>سطر الختم 2</label><input type="text" id="stamp-l2" oninput="updStamp()"></div>
      <div class="fi"><label>لون الختم</label><input type="color" id="stamp-col" value="#C9A84C" style="height:36px" oninput="updStamp()"></div>
    </div>
  </div>
  <div>
    <div style="font-size:13px;font-weight:700;color:var(--gold);margin-bottom:12px">بيانات الشركة</div>
    <div class="fg" style="grid-template-columns:1fr">
      <div class="fi"><label>الاسم (عربي)</label><input type="text" id="co-nar" placeholder="شرق الاحتراف للتجارة"></div>
      <div class="fi"><label>الاسم (إنجليزي)</label><input type="text" id="co-nen" placeholder="Sharq Al-Ihtraf Co."></div>
      <div class="fi"><label>السجل التجاري</label><input type="text" id="co-cr" placeholder="1234567890"></div>
      <div class="fi"><label>الرقم الضريبي (VAT)</label>
        <div style="display:flex;gap:5px"><input type="text" id="co-vat" placeholder="300XXXXXXXXXXXXX" style="flex:1">
        <button onclick="saveVAT()" style="padding:8px;border:1px solid rgba(46,204,113,.3);border-radius:var(--rs);background:rgba(46,204,113,.08);color:var(--ok);cursor:pointer;font-size:12px">💾</button>
        <button onclick="clearVAT()" style="padding:8px;border:1px solid rgba(231,76,60,.3);border-radius:var(--rs);background:rgba(231,76,60,.06);color:var(--err);cursor:pointer;font-size:12px">✕</button></div>
        <div id="vat-disp" style="font-size:11px;color:var(--ts);margin-top:3px"></div></div>
      <div class="fi"><label>العنوان</label><input type="text" id="co-addr" placeholder="الرياض - حي الشفا"></div>
      <div class="fi"><label>الهاتف</label><input type="text" id="co-ph" placeholder="+966 50 228 3783"></div>
      <div class="fi"><label>البريد</label><input type="email" id="co-em" placeholder="info@sharq.sa"></div>
    </div>
  </div>
</div>
<div style="margin-top:16px;padding-top:14px;border-top:1px solid var(--bord)">
  <div style="font-size:13px;font-weight:700;color:var(--gold);margin-bottom:12px">تسلسل الأرقام</div>
  <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:14px">
    <div class="fi"><label>بادئة الفاتورة</label><input type="text" id="seq-ip" value="INV-" maxlength="6"><label style="margin-top:6px">بداية من</label><input type="number" id="seq-is" value="1001" min="1"></div>
    <div class="fi"><label>بادئة القبض</label><input type="text" id="seq-rp" value="REC-" maxlength="6"><label style="margin-top:6px">بداية من</label><input type="number" id="seq-rs" value="1001" min="1"></div>
    <div class="fi"><label>بادئة الصرف</label><input type="text" id="seq-pp" value="PAY-" maxlength="6"><label style="margin-top:6px">بداية من</label><input type="number" id="seq-ps" value="1001" min="1"></div>
  </div>
  <div style="margin-top:14px;display:flex;gap:14px;flex-wrap:wrap">
    <label style="display:flex;align-items:center;gap:6px;font-size:13px;color:var(--tp)"><input type="checkbox" id="vat-en" checked style="width:auto"> تفعيل ضريبة القيمة المضافة (15%)</label>
    <label style="display:flex;align-items:center;gap:6px;font-size:13px;color:var(--tp)"><input type="checkbox" id="wm-en" style="width:auto"> علامة مائية للمسودات</label>
  </div>
</div></div></div>
</main>
</div>

<div class="mo" id="doc-modal"><div class="modal" style="width:760px;max-width:96vw;padding:0;overflow:hidden"><div style="padding:14px 20px;border-bottom:1px solid var(--bord);display:flex;align-items:center;justify-content:space-between;background:var(--card2)"><div class="ct" id="doc-modal-title">مستند جديد</div><button class="mclose" onclick="closeDocModal()">✕</button></div><div style="padding:20px;max-height:75vh;overflow-y:auto"><div class="fg" id="doc-form-fields"></div><div id="inv-items-section" style="display:none;margin-top:14px"><div style="font-size:12px;font-weight:700;color:var(--gold);margin-bottom:8px">بنود الفاتورة</div><div id="inv-items-list"></div><button class="btn bto" onclick="addInvItem()" style="margin-top:8px;font-size:11px">➕ إضافة بند</button><div style="display:flex;justify-content:flex-end;margin-top:12px"><div style="background:rgba(201,168,76,0.08);border:1px solid var(--bord);border-radius:var(--rs);padding:12px 16px;min-width:200px"><div style="display:flex;justify-content:space-between;font-size:12px;padding:3px 0"><span style="color:var(--ts)">قبل الضريبة</span><span id="inv-sub">0.00</span></div><div style="display:flex;justify-content:space-between;font-size:12px;padding:3px 0"><span style="color:var(--ts)">VAT 15%</span><span id="inv-vat">0.00</span></div><div style="display:flex;justify-content:space-between;font-size:13px;font-weight:800;padding:5px 0;border-top:1px solid var(--bord);margin-top:4px;color:var(--gold)"><span>الإجمالي</span><span id="inv-gt">0.00</span></div></div></div></div><div class="br"><button class="btn btp" onclick="saveDoc()">💾 حفظ</button><button class="btn bto" onclick="previewDoc()">👁️ معاينة</button><button class="btn bto" onclick="closeDocModal()">إلغاء</button></div></div></div></div>
<div class="mo" id="doc-preview-modal"><div class="modal" style="width:880px;max-width:97vw;padding:0"><div style="padding:14px 20px;border-bottom:1px solid var(--bord);display:flex;align-items:center;justify-content:space-between;background:var(--card2)"><div class="ct" id="prev-title">معاينة</div><button class="mclose" onclick="closePrev()">✕</button></div><div class="preview-body" id="preview-body"></div><div class="preview-actions"><button class="btn btp" onclick="printDoc()">🖨️ طباعة</button><button class="btn bto" onclick="dlDocPDF()">📄 PDF</button><button class="btn bto" onclick="dlDocXLS()">📊 Excel</button><button class="btn bto" onclick="shareDocWA()">💬 واتساب</button><button class="btn bto" onclick="closePrev()">إغلاق</button></div></div></div>
<div class="mo" id="edit-entry-modal"><div class="modal" style="width:680px;max-width:96vw;padding:0;overflow:hidden"><div style="padding:14px 20px;border-bottom:1px solid var(--bord);display:flex;align-items:center;justify-content:space-between;background:var(--card2)"><div class="ct">✏️ تعديل القيد</div><button class="mclose" onclick="closeEditE()">✕</button></div><div style="padding:20px;max-height:78vh;overflow-y:auto"><div style="margin-bottom:14px;display:flex;gap:6px;flex-wrap:wrap"><button class="edit-type-btn on" data-t="مبيعات" onclick="setET(this,'مبيعات')">💰 مبيعات</button><button class="edit-type-btn" data-t="مشتريات" onclick="setET(this,'مشتريات')">🛒 مشتريات</button><button class="edit-type-btn" data-t="مسحوبات" onclick="setET(this,'مسحوبات')">💳 مسحوبات</button><button class="edit-type-btn" data-t="مصروفات" onclick="setET(this,'مصروفات')">⚙️ مصروفات</button></div><div class="fg" style="grid-template-columns:repeat(2,1fr)"><div class="fi"><label>التاريخ</label><input type="date" id="ee-date"></div><div class="fi"><label>المبلغ</label><input type="number" id="ee-amt" oninput="liveEC()"></div><div class="fi"><label>المحول</label><input type="text" id="ee-snd"></div><div class="fi"><label>المستفيد</label><input type="text" id="ee-rec"></div><div class="fi"><label>التصنيف</label><select id="ee-cat"></select></div><div class="fi"><label>الدفع</label><select id="ee-pay"><option>تحويل بنكي</option><option>نقدي</option><option>شيك</option><option>بطاقة</option></select></div><div class="fi full"><label>ملاحظات</label><textarea id="ee-notes" style="min-height:50px"></textarea></div></div><div id="ee-preview" style="margin-top:12px;padding:12px 16px;background:rgba(201,168,76,0.06);border:1px solid var(--bord);border-radius:var(--rs);display:flex;align-items:center;justify-content:space-between"><span style="color:var(--ts);font-size:12px">المبلغ المعدَّل:</span><span id="ee-amt-prev" style="font-size:18px;font-weight:800;color:var(--gold)">0.00 ر.س</span></div><div class="br" style="margin-top:14px"><button class="btn btp" onclick="saveEditE()">💾 حفظ</button><button class="btn bto" onclick="closeEditE()">إلغاء</button><button class="btn" onclick="delFromEdit()" style="background:rgba(231,76,60,0.08);border:1px solid rgba(231,76,60,0.3);color:var(--err);margin-right:auto">🗑️ حذف</button></div></div></div></div>
<div class="mo" id="invite-modal"><div class="modal"><div class="mtitle">✉️ دعوة عضو <button class="mclose" onclick="closeInvite()">✕</button></div><div class="fi" style="margin-bottom:10px"><label>الاسم</label><input type="text" id="inv-nm"></div><div class="fi" style="margin-bottom:10px"><label>البريد</label><input type="email" id="inv-em"></div><div class="fi" style="margin-bottom:10px"><label>الدور</label><select id="inv-rl"><option value="manager">📊 محاسب رئيسي</option><option value="accountant">🧾 محاسب ثانوي</option><option value="employee">👤 موظف</option></select></div><div style="margin-bottom:14px" id="inv-perms"><div class="perm-toggle"><span style="font-size:13px">📊 عرض التقارير</span><div class="toggle-sw on" data-p="view_reports" onclick="togglePerm(this)"></div></div><div class="perm-toggle"><span style="font-size:13px">➕ إضافة قيود</span><div class="toggle-sw on" data-p="add_entries" onclick="togglePerm(this)"></div></div><div class="perm-toggle"><span style="font-size:13px">🗑️ حذف قيود</span><div class="toggle-sw" data-p="del_entries" onclick="togglePerm(this)"></div></div><div class="perm-toggle"><span style="font-size:13px">📥 تصدير</span><div class="toggle-sw on" data-p="export" onclick="togglePerm(this)"></div></div><div class="perm-toggle"><span style="font-size:13px">📧 إرسال تقارير</span><div class="toggle-sw" data-p="send_rpts" onclick="togglePerm(this)"></div></div></div><div class="br"><button class="btn btp" onclick="saveInvite()">📤 إرسال</button><button class="btn bto" onclick="closeInvite()">إلغاء</button></div></div></div>
<div class="mo" id="em"><div class="modal"><div class="mtitle">📧 إرسال تقرير <button class="mclose" onclick="closeEM()">✕</button></div><div class="fi" style="margin-bottom:10px"><label>البريد</label><input type="email" id="em-to" value="arido0409@gmail.com"></div><div class="fi" style="margin-bottom:10px"><label>النوع</label><select id="em-type"><option value="daily">يومي</option><option value="weekly">أسبوعي</option><option value="monthly">شهري</option></select></div><div class="fi" style="margin-bottom:10px"><label>واتساب (اختياري)</label><input type="text" id="em-wa" value="+966502283783"></div><div class="fi" style="margin-bottom:16px"><label>ملاحظات</label><textarea id="em-notes" style="min-height:60px"></textarea></div><div class="br"><button class="btn btp" onclick="doSendReport()">📤 إرسال</button><button class="btn bto" onclick="closeEM()">إلغاء</button></div></div></div>
<div id="tc"></div>

<script>
// ══════════ CONSTANTS ══════════
const ADMIN_EMAIL = 'arido0409@gmail.com';
const ADMIN_WA    = '+966502283783';
const DAYS = ['الأحد','الإثنين','الثلاثاء','الأربعاء','الخميس','الجمعة','السبت'];
const ROLES = {
  admin:     {label:'مدير عام',      badge:'role-admin',     icon:'👑', perms:['all']},
  manager:   {label:'محاسب رئيسي',   badge:'role-manager',   icon:'📊', perms:['view_reports','add_entries','del_entries','export','send_rpts','assign_tasks']},
  accountant:{label:'محاسب ثانوي',   badge:'role-accountant',icon:'🧾', perms:['view_reports','add_entries','export']},
  employee:  {label:'موظف',           badge:'role-employee',  icon:'👤', perms:['view_reports','add_entries']}
};
const CATS = {مبيعات:['مبيعات نقدية','مبيعات آجلة','صب وتشكيل','مبيعات صادرة','أخرى'],مشتريات:['مواد خام','بضاعة','آلات','أخرى'],مسحوبات:['رواتب','مسحوبات شخصية','مكافآت','سلف','أخرى'],مصروفات:['كهرباء','إيجار','صيانة','مواصلات','اتصالات','تشغيل عام','أخرى']};

// ══════════ STATE ══════════
let entries    = JSON.parse(localStorage.getItem('se')||'[]');
let teamMembers= JSON.parse(localStorage.getItem('team')||'[]');
let tasks      = JSON.parse(localStorage.getItem('tasks')||'[]');
let pending    = JSON.parse(localStorage.getItem('pending')||'[]');
let invoices   = JSON.parse(localStorage.getItem('docs_inv')||'[]');
let receipts   = JSON.parse(localStorage.getItem('docs_rec')||'[]');
let payments   = JSON.parse(localStorage.getItem('docs_pay')||'[]');
let coSettings = JSON.parse(localStorage.getItem('co_settings')||'{}');
let docCtrs    = JSON.parse(localStorage.getItem('doc_ctrs')||'{"inv":1001,"rec":1001,"pay":1001}');
let schedCfg   = JSON.parse(localStorage.getItem('sched')||'{"mode":"daily","time":"08:00","email":"arido0409@gmail.com","wa":"+966502283783"}');
let currentUser= JSON.parse(localStorage.getItem('cu')||'null');
let curType='مبيعات', ocrData=null, chatHist=[], voiceActive=false, recog=null;
let selectedE=new Set(), editingEId=null, editEOrig=null, editEType='مبيعات';
let currentDocType='invoice', currentDocData=null, editingDocId=null, invItems=[];
let dragTaskId=null;
let waHist=JSON.parse(localStorage.getItem('wa_hist')||'[]');
let waMC=parseInt(localStorage.getItem('wa_mc')||'0'),waCC=parseInt(localStorage.getItem('wa_cc')||'0'),waRC=parseInt(localStorage.getItem('wa_rc')||'0');
let waMethod='twilio';
let schedTimer=null;
let otp=null, otpExpiry=null, otpTries=0, otpInterval=null;

// ══════════ INIT ══════════
window.addEventListener('load',()=>{
  setTimeout(()=>{
    document.getElementById('splash').classList.add('out');
    setTimeout(()=>{
      document.getElementById('splash').style.display='none';
      updateAuthStats();
      if(currentUser){ showApp(); applyRoleUI(); }
    },800);
  },2800);
});

function updateAuthStats(){
  const ic=document.getElementById('auth-inv-count');
  const ec=document.getElementById('auth-entry-count');
  if(ic)ic.textContent=invoices.length;
  if(ec)ec.textContent=entries.length;
}

function showApp(){
  document.getElementById('auth-screen').classList.add('hidden');
  const app=document.getElementById('app');
  app.classList.add('on'); app.style.display='block';
  const t=new Date();
  document.getElementById('today-label').textContent=t.toLocaleDateString('ar-SA',{weekday:'long',year:'numeric',month:'long',day:'numeric'});
  document.getElementById('ed').value=t.toISOString().split('T')[0];
  document.getElementById('rfrom').value=t.toISOString().split('T')[0];
  document.getElementById('rto').value=t.toISOString().split('T')[0];
  updateKPIs(); renderDT(); initVoice(); startSchedTimer();
  setTimeout(refreshDocLists,500);
}

function loadMayData() {}

// ══════════ OTP SYSTEM ══════════
function genOTP(){return Math.floor(100000+Math.random()*900000).toString();}

window.initAdminLogin = function initAdminLogin(){
  otp=genOTP(); otpExpiry=Date.now()+5*60*1000; otpTries=0;
  document.getElementById('main-login-box').style.display='none';
  const ob=document.getElementById('otp-box');
  ob.style.display='block'; ob.style.opacity='0';
  setTimeout(()=>ob.style.opacity='1',50);
  for(let i=0;i<6;i++){const d=document.getElementById('d'+i);d.value='';d.style.borderColor='var(--bord)';}
  document.getElementById('otp-err').textContent='';
  document.getElementById('d0').focus();
  startOTPTimer(); sendOTPNotif();
}

function sendOTPNotif(){
  sendWALink(ADMIN_WA,`🔐 *كود دخول شرق الاحتراف*\n━━━━━━━━━━━\n📟 الكود: *${otp}*\n⏱ صالح 5 دقائق\n⚠️ لا تشاركه مع أحد`);
  console.log('%c🔐 OTP: '+otp,'color:#C9A84C;font-size:24px;font-weight:900;background:#000;padding:8px');
  const tc=document.getElementById('tc');
  const el=document.createElement('div');
  el.className='toast'; el.id='otp-toast';
  el.style.cssText='background:#1a1200;border:2px solid var(--gold);border-radius:12px;padding:14px 20px;font-size:13px;min-width:240px;direction:rtl';
  el.innerHTML=`<div style="color:var(--ts);font-size:11px;margin-bottom:6px">📧 كود التحقق أُرسل إلى ${ADMIN_EMAIL}</div><div style="font-size:28px;font-weight:900;color:var(--gold);letter-spacing:8px;text-align:center">${otp}</div>`;
  tc.appendChild(el);
  setTimeout(()=>el.remove(),20000);
}

function startOTPTimer(){
  if(otpInterval)clearInterval(otpInterval);
  otpInterval=setInterval(()=>{
    const rem=otpExpiry-Date.now();
    if(rem<=0){
      clearInterval(otpInterval);
      document.getElementById('otp-timer').textContent='انتهى الكود';
      otp=null; document.getElementById('otp-btn').disabled=true;
      return;
    }
    const m=Math.floor(rem/60000),s=Math.floor((rem%60000)/1000);
    document.getElementById('otp-timer').textContent=m+':'+(s<10?'0':'')+s;
  },1000);
}

window.otpIn = function otpIn(i){
  const d=document.getElementById('d'+i);
  d.value=d.value.replace(/[^0-9]/g,'');
  d.style.borderColor=d.value?'var(--gold)':'var(--bord)';
  if(d.value){const n=document.getElementById('d'+(i+1));if(n)n.focus();}
  checkOTPFull();
}

window.otpKey = function otpKey(e,i){
  if(e.key==='Backspace'){
    const d=document.getElementById('d'+i);
    if(!d.value&&i>0){const p=document.getElementById('d'+(i-1));p.value='';p.focus();}
  }
}

function checkOTPFull(){
  const full=[0,1,2,3,4,5].every(i=>document.getElementById('d'+i)?.value);
  const btn=document.getElementById('otp-btn');
  btn.disabled=!full; btn.style.opacity=full?'1':'.4';
}

window.verifyOTP = function verifyOTP(){
  const entered=[0,1,2,3,4,5].map(i=>document.getElementById('d'+i)?.value||'').join('');
  if(entered!==otp){ toast('كود التحقق غير صحيح','err'); return; }
  clearInterval(otpInterval);
  loginUser({id:'admin_main',name:'المدير العام',email:ADMIN_EMAIL,role:'admin',perms:['all']});
}

window.cancelOTP = function cancelOTP(){
  if(otpInterval)clearInterval(otpInterval);
  document.getElementById('otp-box').style.display='none';
  document.getElementById('main-login-box').style.display='block';
}

window.resendOTP = function resendOTP(){ resendOTP(); }

// ══════════ CORE BUSINESS & AUTH ══════════
window.requestAccess = function requestAccess(){
  const name=document.getElementById('req-name').value.trim();
  const email=document.getElementById('req-email').value.trim();
  const role=document.getElementById('req-role').value;
  if(!name||!email){toast('يرجى ملا الحقول','err');return;}
  toast('تم إرسال طلب الدخول بنجاح للمدير');
}

function loginUser(user){
  currentUser=user; localStorage.setItem('cu',JSON.stringify(user));
  showApp(); applyRoleUI();
}

window.doSignOut = function doSignOut(){
  currentUser=null; localStorage.removeItem('cu'); window.location.reload();
}

// ══════════ APP FUNCTIONS ══════════
window.SP = function SP(id){
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('on'));
  const pg=document.getElementById('page-'+id); if(pg) pg.classList.add('on');
  if(id==='transactions')renderT();
  if(id==='reports')genReport();
}

window.setTab = function setTab(t){
  const m={s:'مبيعات',p:'مشتريات',w:'مسحوبات',e:'مصروفات'}; curType=m[t];
  document.querySelectorAll('.tabs .tb').forEach(b=>b.classList.remove('on'));
  event.target.classList.add('on');
  buildCatSelect();
}

function buildCatSelect(){
  const s=document.getElementById('ecat'); s.innerHTML='<option value="">— اختر —</option>';
  (CATS[curType]||[]).forEach(c=>s.innerHTML+=`<option value="${c}">${c}</option>`);
}

window.saveEntry = function saveEntry(){
  const amt=parseFloat(document.getElementById('eamt').value);
  const date=document.getElementById('ed').value;
  if(!amt||!date){toast('يرجى كتابة المبلغ والتاريخ','err');return;}
  entries.unshift({id:Date.now(),type:curType,amount:amt,date,sender:document.getElementById('esend').value,recipient:document.getElementById('erec').value,category:document.getElementById('ecat').value,payment:document.getElementById('epay').value,notes:document.getElementById('enotes').value});
  localStorage.setItem('se',JSON.stringify(entries)); updateKPIs(); renderDT(); clearForm(); toast('تم حفظ القيد بنجاح');
}

window.clearForm = function clearForm(){
  ['esend','erec','eamt','enotes'].forEach(id=>document.getElementById(id).value='');
}

function updateKPIs(){
  let s=0, p=0, w=0, e=0;
  entries.forEach(x=>{
    if(x.type==='مبيعات')s+=x.amount; if(x.type==='مشتريات')p+=x.amount;
    if(x.type==='مسحوبات')w+=x.amount; if(x.type==='مصروفات')e+=x.amount;
  });
  document.getElementById('kts').textContent=fmt(s); document.getElementById('ktp').textContent=fmt(p);
  document.getElementById('ktw').textContent=fmt(w); document.getElementById('kte').textContent=fmt(e);
  const net=s-p-w-e; document.getElementById('net-profit').textContent=fmt(net)+' ريال';
}

function renderDT(){
  const tb=document.getElementById('dtable'); tb.innerHTML='';
  entries.slice(0,5).forEach(x=>{
    tb.innerHTML+=`<tr><td>${x.date}</td><td>-</td><td>${x.type}</td><td>${x.sender||'—'}</td><td>${x.recipient||'—'}</td><td>${fmt(x.amount)}</td><td>${x.category}</td></tr>`;
  });
}

function renderT(){
  const tb=document.getElementById('ttable'); tb.innerHTML='';
  entries.forEach((x,i)=>{
    tb.innerHTML+=`<tr><td><input type="checkbox" class="rc" data-id="${x.id}"></td><td>${i+1}</td><td>${x.date}</td><td>-</td><td>${x.type}</td><td>${x.sender}</td><td>${x.recipient}</td><td>${fmt(x.amount)}</td><td>${x.category}</td><td>${x.payment}</td><td>${x.notes}</td><td><button onclick="delEntry(${x.id})">🗑</button></td></tr>`;
  });
}

window.delEntry = function delEntry(id){
  entries=entries.filter(x=>x.id!==id); localStorage.setItem('se',JSON.stringify(entries)); updateKPIs(); renderDT(); renderT();
}

// ══════════ REVENUE DOCUMENTS (INVOICES & RECEIPTS) ══════════
window.refreshDocLists = function refreshDocLists(){
  const il=document.getElementById('invoices-list'); if(!il) return;
  il.innerHTML=invoices.length===0?'<div style="text-align:center;padding:36px;color:var(--tm)">لا توجد فواتير</div>':
  invoices.map(x=>`<div class="doc-list-item" onclick="viewDocDetails('${x.id}','invoice')"><span class="doc-num">${x.num}</span><div class="doc-info"><div class="doc-title">${x.client}</div><div class="doc-sub">${x.date} | الإجمالي: ${fmt(x.grandTotal)} ر.س</div></div><span class="st-${x.status}">${x.status==='paid'?'مدفوعة':'معلقة'}</span></div>`).join('');

  const rl=document.getElementById('receipts-list'); if(!rl) return;
  rl.innerHTML=receipts.length===0?'<div style="text-align:center;padding:36px;color:var(--tm)">لا توجد سندات</div>':
  receipts.map(x=>`<div class="doc-list-item" onclick="viewDocDetails('${x.id}','receipt')"><span class="doc-num">${x.num}</span><div class="doc-info"><div class="doc-title">${x.from}</div><div class="doc-sub">${x.date} | المبلغ: ${fmt(x.amount)} ر.س</div></div></div>`).join('');

  const pl=document.getElementById('payments-list'); if(!pl) return;
  pl.innerHTML=payments.length===0?'<div style="text-align:center;padding:36px;color:var(--tm)">لا توجد سندات</div>':
  payments.map(x=>`<div class="doc-list-item" onclick="viewDocDetails('${x.id}','payment')"><span class="doc-num">${x.num}</span><div class="doc-info"><div class="doc-title">${x.to}</div><div class="doc-sub">${x.date} | المبلغ: ${fmt(x.amount)} ر.س</div></div></div>`).join('');
}

window.openNewDoc = function openNewDoc(type){
  currentDocType=type; editingDocId=null;
  const fields=document.getElementById('doc-form-fields');
  const num=getNum(type==='invoice'?'inv':type==='receipt'?'rec':'pay');
  document.getElementById('doc-modal-title').textContent=DOC_D[type].label;
  
  if(type==='invoice'){
    document.getElementById('inv-items-section').style.display='block'; invItems=[{desc:'',qty:1,price:0}]; renderInvItems();
    fields.innerHTML=`<div class="fi"><label>رقم الفاتورة</label><input type="text" id="df-num" value="${num}" readonly></div><div class="fi"><label>التاريخ</label><input type="date" id="df-date" value="${new Date().toISOString().split('T')[0]}"></div><div class="fi"><label>اسم العميل</label><input type="text" id="df-cl" placeholder="الزبون"></div><div class="fi"><label>حالة الدفع</label><select id="df-st"><option value="unpaid">غير مدفوعة</option><option value="paid">مدفوعة</option></select></div>`;
  } else {
    document.getElementById('inv-items-section').style.display='none';
    fields.innerHTML=`<div class="fi"><label>رقم السند</label><input type="text" id="df-num" value="${num}" readonly></div><div class="fi"><label>التاريخ</label><input type="date" id="df-date" value="${new Date().toISOString().split('T')[0]}"></div><div class="fi full"><label>${type==='receipt'?'استلمنا من':'صرفنا إلى'}</label><input type="text" id="df-person" placeholder="الاسم الكامل"></div><div class="fi"><label>المبلغ</label><input type="number" id="df-amount" placeholder="0.00"></div><div class="fi"><label>طريقة الدفع</label><input type="text" id="df-method" value="تحويل بنكي"></div><div class="fi full"><label>وذلك مقابل</label><textarea id="df-reason" placeholder="البيان الإضافي للشرح"></textarea></div>`;
  }
  document.getElementById('doc-modal').classList.add('open');
}

window.addInvItem = function addInvItem(){ invItems.push({desc:'',qty:1,price:0}); renderInvItems(); }
window.delInvItem = function delInvItem(i){ invItems.splice(i,1); renderInvItems(); }

function renderInvItems(){
  const container=document.getElementById('inv-items-list'); container.innerHTML='';
  invItems.forEach((x,i)=> {
    container.innerHTML+=`<div style="display:flex;gap:8px;margin-bottom:6px"><input type="text" placeholder="الوصف والبيان" value="${x.desc}" oninput="invItems(${i}).desc=this.value" style="flex:2"> <input type="number" placeholder="الكمية" value="${x.qty}" oninput="invItems[${i}].qty=parseInt(this.value)||0;calcDocTotals()" style="width:70px"> <input type="number" placeholder="السعر" value="${x.price}" oninput="invItems[${i}].price=parseFloat(this.value)||0;calcDocTotals()" style="width:100px"> <button onclick="delInvItem(${i})" style="color:var(--err);background:none;border:none;cursor:pointer">✕</button></div>`;
  });
  calcDocTotals();
}

window.calcDocTotals = function calcDocTotals(){
  let sub=0; invItems.forEach(x=>sub+=(x.qty*x.price));
  let vat=sub*0.15; let gt=sub+vat;
  document.getElementById('inv-sub').textContent=sub.toFixed(2);
  document.getElementById('inv-vat').textContent=vat.toFixed(2);
  document.getElementById('inv-gt').textContent=gt.toFixed(2);
}

window.saveDoc = function saveDoc(){
  const date=document.getElementById('df-date').value;
  const num=document.getElementById('df-num').value;
  if(currentDocType==='invoice'){
    const client=document.getElementById('df-cl').value;
    let sub=0; invItems.forEach(x=>sub+=(x.qty*x.price));
    invoices.push({id:'i_'+Date.now(),num,date,client,items:invItems,subTotal:sub,vatTotal:sub*0.15,grandTotal:sub*1.15,status:document.getElementById('df-st').value});
    bumpCtr('inv'); saveDocs();
  } else if(currentDocType==='receipt'){
    receipts.push({id:'r_'+Date.now(),num,date,from:document.getElementById('df-person').value,amount:parseFloat(document.getElementById('df-amount').value),method:document.getElementById('df-method').value,reason:document.getElementById('df-reason').value});
    bumpCtr('rec'); saveDocs();
  } else {
    payments.push({id:'p_'+Date.now(),num,date,to:document.getElementById('df-person').value,amount:parseFloat(document.getElementById('df-amount').value),method:document.getElementById('df-method').value,reason:document.getElementById('df-reason').value});
    bumpCtr('pay'); saveDocs();
  }
  closeDocModal(); refreshDocLists(); toast('تم حفظ المستند وإصدار الرقم');
}

window.closeDocModal = function closeDocModal(){ document.getElementById('doc-modal').classList.remove('open'); }

window.viewDocDetails = function viewDocDetails(id,type){
  let doc=null; if(type==='invoice') doc=invoices.find(x=>x.id===id);
  if(type==='receipt') doc=receipts.find(x=>x.id===id);
  if(type==='payment') doc=payments.find(x=>x.id===id);
  if(!doc) return; currentDocData=doc; currentDocType=type;
  
  const pb=document.getElementById('preview-body');
  if(type==='invoice'){
    pb.innerHTML=`<div style="background:#fff;color:#000;padding:30px;font-family:sans-serif;direction:rtl"><div style="display:flex;justify-content:space-between"><div><h2>${coSettings.co_nar||'شرق الاحتراف للتجارة'}</h2><p>${coSettings.co_addr||'الرياض'}</p><p>الرقم الضريبي: ${coSettings.co_vat||'—'}</p></div><div style="text-align:left"><h1>فاتورة مبيعات</h1><p>الرقم: ${doc.num}</p><p>التاريخ: ${doc.date}</p></div></div><hr style="margin:20px 0"><p>المطلوب من المكرم: <strong>${doc.client}</strong></p><table style="width:100%;border-collapse:collapse;margin-top:20px;color:#000"><thead style="background:#eee"><tr><th style="padding:8px;border:1px solid #ccc">البيان</th><th style="width:60px;padding:8px;border:1px solid #ccc">الكمية</th><th style="width:100px;padding:8px;border:1px solid #ccc">السعر</th><th style="width:100px;padding:8px;border:1px solid #ccc">الإجمالي</th></tr></thead><tbody>${doc.items.map(x=>`<tr><td style="padding:8px;border:1px solid #ccc">${x.desc}</td><td style="padding:8px;border:1px solid #ccc;text-align:center">${x.qty}</td><td style="padding:8px;border:1px solid #ccc;text-align:center">${fmt(x.price)}</td><td style="padding:8px;border:1px solid #ccc;text-align:center">${fmt(x.qty*x.price)}</td></tr>`).join('')}</tbody></table><div style="margin-top:20px;text-align:left;float:left;width:250px"><p>المجموع: ${fmt(doc.subTotal)} ر.س</p><p>الضريبة (15%): ${fmt(doc.vatTotal)} ر.س</p><h3>الإجمالي النهائي: ${fmt(doc.grandTotal)} ر.س</h3></div><div style="clear:both"></div></div>`;
  } else {
    pb.innerHTML=`<div style="background:#fff;color:#000;padding:40px;border:2px dashed #999;font-family:sans-serif;direction:rtl;max-width:700px;margin:0 auto"><h2 style="text-align:center;margin-bottom:20px">${type==='receipt'?'سند قبض مالي':'سند صرف مالي'}</h2><div style="display:flex;justify-content:space-between"><p>الرقم: ${doc.num}</p><p>التاريخ: ${doc.date}</p></div><p style="margin-top:20px;line-height:2">${type==='receipt'?'إستلمنا من المكرم':'صرفنا إلى المكرم'}: <strong>${type==='receipt'?doc.from:doc.to}</strong></p><p style="line-height:2">مبلغ وقدره: <strong style="font-size:16px;background:#eee;padding:2px 10px">${fmt(doc.amount)} ريال سعودي</strong></p><p style="line-height:2">طريقة الدفع: ${doc.method}</p><p style="line-height:2">وذلك مقابل: ${doc.reason||'—'}</p><div style="margin-top:4px;display:flex;justify-content:space-between"><p>المحاسب المسؤول: ____________</p><p>التوقيع والختم: ____________</p></div></div>`;
  }
  document.getElementById('doc-preview-modal').classList.add('open');
}

window.closePrev = function closePrev(){ document.getElementById('doc-preview-modal').classList.remove('open'); }
window.printDoc = function printDoc(){ window.print(); }

// ══════════ COMPANY SETTINGS ══════════
function loadCoUI(){
  document.getElementById('co-nar').value=coSettings.co_nar||'';
  document.getElementById('co-nen').value=coSettings.co_nen||'';
  document.getElementById('co-cr').value=coSettings.co_cr||'';
  document.getElementById('co-vat').value=coSettings.co_vat||'';
  document.getElementById('co-addr').value=coSettings.co_addr||'';
  document.getElementById('co-ph').value=coSettings.co_ph||'';
  document.getElementById('co-em').value=coSettings.co_em||'';
}

window.saveCoSettings = function saveCoSettings(){
  coSettings.co_nar=document.getElementById('co-nar').value;
  coSettings.co_nen=document.getElementById('co-nen').value;
  coSettings.co_cr=document.getElementById('co-cr').value;
  coSettings.co_vat=document.getElementById('co-vat').value;
  coSettings.co_addr=document.getElementById('co-addr').value;
  coSettings.co_ph=document.getElementById('co-ph').value;
  coSettings.co_em=document.getElementById('co-em').value;
  localStorage.setItem('co_settings',JSON.stringify(coSettings)); toast('تم حفظ ملف إعدادات الشركة بنجاح');
}

// ══════════ AUXILIARIES ══════════
function toast(m,t='ok'){
  const container=document.getElementById('tc'); if(!container) return;
  const el=document.createElement('div'); el.className='toast '+t; el.textContent=m;
  container.appendChild(el); setTimeout(()=>el.remove(),4000);
}

// Global UI trigger events
document.addEventListener('DOMContentLoaded', () => {
  setTab('s');
});
</script>
</body>
</html>
