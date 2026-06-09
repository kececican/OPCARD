# SETUP_CLAUDE.md — Modül Kurulum Talimatı

> Bu dosya yalnızca ilk kurulum içindir. Entegrasyon tamamlandıktan sonra sil; sadece CLAUDE.md kalsın.

---

## Tetikleyici

Kullanıcı **"projeyi başlat"** veya **"modülleri entegre et"** dediğinde aşağıdaki adımları sırayla uygula.

---

## Kurulum Adımları

### 1. Feedback Widget — Code.gs
`Code.gs` dosyasını aç. `submitFeedback` fonksiyonu var mı kontrol et.
- **Yoksa:** Modül 1 GAS bloğunu dosyanın sonuna ekle.
- **Varsa:** Dokunma.

### 2. Feedback Widget — Index.html
`Index.html` dosyasını aç. `id="fb-btn"` elementi var mı kontrol et.
- **Yoksa:** Modül 1 HTML bloğunu `</body>` etiketinin hemen öncesine ekle.
- **Varsa:** Dokunma.

### 3. Giriş/Çıkış Loglama — Code.gs
`Code.gs` dosyasını aç. `createSession` fonksiyonu var mı kontrol et.

- **Yoksa ve `doGet()` de tanımlı değilse:** Modül 2 GAS bloğunun tamamını dosyanın sonuna ekle.
- **Yoksa ama `doGet()` zaten mevcutsa:** Sadece `createSession`, `logExit`, `_getOrCreateLogSheet` fonksiyonlarını ekle. Mevcut `doGet()` içinde `template.evaluate()` çağrısından önce şu 4 satırı entegre et:
  ```javascript
  var session = createSession();
  template.sessionId  = session.sessionId;
  template.userEmail  = session.email;
  template.timeoutMin = TIMEOUT_MIN;
  ```
- **Varsa:** Dokunma.

### 4. Giriş/Çıkış Loglama — Index.html
`Index.html` dosyasını aç. `SESSION_ID` değişkeni var mı kontrol et.
- **Yoksa:** Modül 2 HTML bloğunu `</body>` etiketinin hemen öncesine ekle.
- **Varsa:** Dokunma.

### 5. Bitti — Kullanıcıya bildir

> "Entegrasyon tamamlandı. Şu değişkenleri kendi bilginle güncelle:
> - `Code.gs` → `FB_SHEET_ID`, `PROJECT_NAME`
> - `Index.html` → `FB_APP_NAME`
>
> Güncellemeler bittikten sonra `SETUP_CLAUDE.md` dosyasını silebilirsin."

---

## Modül 1 — Feedback Widget

### Code.gs'e eklenecek

```javascript
// ─── Feedback Widget ───────────────────────────────────────────────────────────
var FB_SHEET_ID = '1ZdakLmkO8s57T-WU1MGuyMFitzRQxO5XxhuIL7UoMwI';  // ← güncelle

function submitFeedback(payload) {
  var ss      = SpreadsheetApp.openById(FB_SHEET_ID);
  var tabName = payload.appName || 'Genel';
  var sheet   = ss.getSheetByName(tabName);

  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.getRange(1, 1, 1, 9).setValues([[
      'FeedbackID','Email','Feedback_Type','Priority','Message',
      'CreatedAt','Comments','Status','Standardization Y/N'
    ]]).setBackground('#4A6CF7').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 110); sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 120); sheet.setColumnWidth(4, 90);
    sheet.setColumnWidth(5, 320); sheet.setColumnWidth(6, 160);
    sheet.setColumnWidth(7, 200);
  }

  var id    = Utilities.getUuid().substring(0, 8).toUpperCase();
  var now   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var email = Session.getActiveUser().getEmail();

  sheet.appendRow([id, email,
    payload.feedbackType || '', payload.priority || '',
    payload.message || '', now, '', 'Open', '']);

  return { success: true, id: id };
}
```

### Index.html — `</body>` öncesine eklenecek

```html
<style>
  #fb-btn{position:fixed;bottom:24px;left:20px;display:flex;align-items:center;gap:8px;
    padding:10px 18px;background:#1E293B;color:#fff;border:none;border-radius:99px;
    font-size:13px;font-weight:600;cursor:pointer;box-shadow:0 4px 16px rgba(0,0,0,.25);
    z-index:8000;transition:all .18s ease;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;}
  #fb-btn:hover{background:#0F172A;transform:translateY(-2px);}
  #fb-overlay{display:none;position:fixed;inset:0;background:rgba(15,23,42,.5);
    z-index:8001;align-items:flex-end;justify-content:flex-start;padding:0 0 80px 20px;}
  #fb-overlay.fb-open{display:flex;}
  #fb-box{background:#fff;border-radius:16px;width:340px;box-shadow:0 8px 32px rgba(0,0,0,.2);
    overflow:hidden;animation:fbSlide .2s ease;
    font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;}
  @keyframes fbSlide{from{transform:translateY(16px);opacity:0}to{transform:translateY(0);opacity:1}}
  #fb-head{display:flex;align-items:center;justify-content:space-between;
    padding:14px 16px;border-bottom:1px solid #E2E8F0;}
  #fb-head span{font-size:14px;font-weight:700;color:#1E293B;}
  #fb-head button{background:none;border:none;font-size:17px;cursor:pointer;color:#94A3B8;
    padding:2px 5px;border-radius:4px;line-height:1;}
  #fb-head button:hover{background:#F1F5F9;color:#1E293B;}
  #fb-body{padding:14px 16px;display:flex;flex-direction:column;gap:12px;}
  .fb-label{font-size:11px;font-weight:700;color:#94A3B8;text-transform:uppercase;
    letter-spacing:.06em;display:block;margin-bottom:5px;}
  .fb-required{color:#EF4444;}
  #fb-type-wrap,#fb-pri-wrap{display:flex;gap:6px;}
  .fb-type,.fb-pri{flex:1;padding:8px 4px;border:2px solid #E2E8F0;border-radius:8px;
    background:#fff;font-size:12px;font-weight:600;color:#64748B;cursor:pointer;
    transition:all .14s;text-align:center;font-family:inherit;}
  .fb-type:hover,.fb-pri:hover{border-color:#94A3B8;}
  .fb-type[data-v="Bug"].fb-sel        {border-color:#EF4444;background:#FEE2E2;color:#991B1B;}
  .fb-type[data-v="Improvement"].fb-sel{border-color:#22C55E;background:#DCFCE7;color:#15803D;}
  .fb-pri[data-v="Low"].fb-sel   {border-color:#3B82F6;background:#EFF6FF;color:#1D4ED8;}
  .fb-pri[data-v="Medium"].fb-sel{border-color:#F59E0B;background:#FEF3C7;color:#92400E;}
  .fb-pri[data-v="Urgent"].fb-sel{border-color:#EF4444;background:#FEE2E2;color:#991B1B;}
  #fb-type-err,#fb-pri-err,#fb-msg-err{color:#EF4444;font-size:11px;display:none;margin-top:4px;}
  #fb-message{width:100%;padding:9px 11px;border:1px solid #E2E8F0;border-radius:8px;
    font-size:13px;font-family:inherit;resize:vertical;min-height:90px;
    outline:none;color:#1E293B;transition:border-color .14s;}
  #fb-message:focus{border-color:#4A6CF7;}
  #fb-message.fb-err{border-color:#EF4444;}
  #fb-foot{display:flex;justify-content:flex-end;gap:8px;padding:12px 16px;
    border-top:1px solid #E2E8F0;}
  .fb-cancel{padding:8px 14px;border:1px solid #E2E8F0;border-radius:6px;background:#fff;
    font-size:13px;font-weight:600;color:#64748B;cursor:pointer;font-family:inherit;}
  .fb-cancel:hover{background:#F1F5F9;}
  .fb-submit{padding:8px 16px;border:none;border-radius:6px;background:#4A6CF7;
    color:#fff;font-size:13px;font-weight:600;cursor:pointer;font-family:inherit;transition:background .14s;}
  .fb-submit:hover{background:#3A5AE0;}
  .fb-submit:disabled{background:#94A3B8;cursor:not-allowed;}
  #fb-success{display:none;padding:28px 16px;text-align:center;}
  #fb-success .fb-ok-icon{font-size:38px;margin-bottom:8px;}
  #fb-success strong{font-size:15px;color:#1E293B;display:block;}
  #fb-success p{color:#64748B;font-size:13px;margin-top:5px;}
  @media(max-width:600px){
    #fb-btn{bottom:16px;left:12px;padding:9px 14px;font-size:12px;}
    #fb-overlay{padding:0 0 68px 12px;}
    #fb-box{width:calc(100vw - 24px);}
  }
</style>

<button id="fb-btn" onclick="fbToggle()">
  <svg width="15" height="15" viewBox="0 0 24 24" fill="none"
       stroke="#4ADE80" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
    <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
  </svg>
  Geri Bildirim
</button>

<div id="fb-overlay" onclick="fbOverlayClick(event)">
  <div id="fb-box">
    <div id="fb-head">
      <span>Geri Bildirim</span>
      <button onclick="fbClose()">&#x2715;</button>
    </div>
    <div id="fb-form-wrap">
      <div id="fb-body">
        <div>
          <span class="fb-label">Tür <span class="fb-required">*</span></span>
          <div id="fb-type-wrap">
            <button type="button" class="fb-type" data-v="Bug"         onclick="fbType(this)">Bug</button>
            <button type="button" class="fb-type" data-v="Improvement" onclick="fbType(this)">İyileştirme</button>
          </div>
          <div id="fb-type-err">Lütfen bir tür seçin.</div>
        </div>
        <div>
          <span class="fb-label">Öncelik <span class="fb-required">*</span></span>
          <div id="fb-pri-wrap">
            <button type="button" class="fb-pri" data-v="Low"    onclick="fbPri(this)">Düşük</button>
            <button type="button" class="fb-pri" data-v="Medium" onclick="fbPri(this)">Orta</button>
            <button type="button" class="fb-pri" data-v="Urgent" onclick="fbPri(this)">Acil</button>
          </div>
          <div id="fb-pri-err">Lütfen bir öncelik seçin.</div>
        </div>
        <div>
          <span class="fb-label">Mesaj <span class="fb-required">*</span></span>
          <textarea id="fb-message" placeholder="Geri bildiriminizi yazın..."></textarea>
          <div id="fb-msg-err">Lütfen mesajınızı girin.</div>
        </div>
      </div>
      <div id="fb-foot">
        <button class="fb-cancel" onclick="fbClose()">İptal</button>
        <button class="fb-submit" id="fb-submit-btn" onclick="fbSubmit()">Gönder</button>
      </div>
    </div>
    <div id="fb-success">
      <div class="fb-ok-icon">&#x2705;</div>
      <strong>Teşekkürler!</strong>
      <p>Geri bildiriminiz kaydedildi.</p>
    </div>
  </div>
</div>

<script>
  var FB_APP_NAME = 'OPCARD';  // ← güncelle

  var _fbPriority = '', _fbType = '';

  function fbToggle(){document.getElementById('fb-overlay').classList.contains('fb-open')?fbClose():fbOpen();}
  function fbOpen(){fbReset();document.getElementById('fb-overlay').classList.add('fb-open');}
  function fbClose(){document.getElementById('fb-overlay').classList.remove('fb-open');}
  function fbOverlayClick(e){if(e.target===document.getElementById('fb-overlay'))fbClose();}
  function fbType(btn){
    document.querySelectorAll('.fb-type').forEach(function(b){b.classList.remove('fb-sel');});
    btn.classList.add('fb-sel'); _fbType=btn.getAttribute('data-v');
    document.getElementById('fb-type-err').style.display='none';
  }
  function fbPri(btn){
    document.querySelectorAll('.fb-pri').forEach(function(b){b.classList.remove('fb-sel');});
    btn.classList.add('fb-sel'); _fbPriority=btn.getAttribute('data-v');
    document.getElementById('fb-pri-err').style.display='none';
  }
  function fbReset(){
    _fbPriority=''; _fbType='';
    document.querySelectorAll('.fb-pri,.fb-type').forEach(function(b){b.classList.remove('fb-sel');});
    document.getElementById('fb-message').value='';
    document.getElementById('fb-message').classList.remove('fb-err');
    ['fb-type-err','fb-pri-err','fb-msg-err'].forEach(function(id){
      document.getElementById(id).style.display='none';
    });
    document.getElementById('fb-form-wrap').style.display='';
    document.getElementById('fb-success').style.display='none';
    var btn=document.getElementById('fb-submit-btn');
    btn.disabled=false; btn.textContent='Gönder';
  }
  function fbSubmit(){
    var msg=document.getElementById('fb-message').value.trim();
    var ok=true;
    if(!_fbType){document.getElementById('fb-type-err').style.display='block';ok=false;}
    if(!_fbPriority){document.getElementById('fb-pri-err').style.display='block';ok=false;}
    if(!msg){
      document.getElementById('fb-message').classList.add('fb-err');
      document.getElementById('fb-msg-err').style.display='block'; ok=false;
    } else {
      document.getElementById('fb-message').classList.remove('fb-err');
      document.getElementById('fb-msg-err').style.display='none';
    }
    if(!ok) return;
    var btn=document.getElementById('fb-submit-btn');
    btn.disabled=true; btn.textContent='Gönderiliyor...';
    google.script.run
      .withSuccessHandler(function(){
        document.getElementById('fb-form-wrap').style.display='none';
        document.getElementById('fb-success').style.display='block';
        setTimeout(fbClose, 2200);
      })
      .withFailureHandler(function(err){
        btn.disabled=false; btn.textContent='Gönder';
        alert('Hata: '+(err.message||err));
      })
      .submitFeedback({appName:FB_APP_NAME, feedbackType:_fbType, priority:_fbPriority, message:msg});
  }
  document.addEventListener('keydown', function(e){if(e.key==='Escape') fbClose();});
</script>
```

---

## Modül 2 — Giriş/Çıkış Loglama

### Code.gs'e eklenecek

> Adım 3'teki kurala göre: `doGet()` yoksa tüm bloğu ekle. `doGet()` varsa yalnızca `createSession`, `logExit`, `_getOrCreateLogSheet` fonksiyonlarını ekle ve mevcut `doGet()` içine session satırlarını entegre et.

```javascript
// ─── Giriş/Çıkış Loglama ──────────────────────────────────────────────────────
// FB_SHEET_ID Modül 1'de tanımlı — aynı değişkeni kullanır
var PROJECT_NAME = 'OPCARD';  // ← güncelle
var LOG_SHEET    = PROJECT_NAME + '_GirisLoglari';
var TIMEOUT_MIN  = 10;

// Bu doGet() bloğunu SADECE projede henüz doGet() yoksa ekle.
// doGet() zaten varsa bu fonksiyonu atlayıp createSession/logExit'i ekle.
function doGet(e) {
  var session  = createSession();
  var template = HtmlService.createTemplateFromFile('Index');
  template.sessionId  = session.sessionId;
  template.userEmail  = session.email;
  template.timeoutMin = TIMEOUT_MIN;
  return template.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function createSession() {
  var sheet     = _getOrCreateLogSheet();
  var sessionId = Utilities.getUuid();
  var email     = Session.getActiveUser().getEmail() || 'Anonim';
  sheet.appendRow([sessionId, email, new Date(), '', '', 'Açık']);
  return { sessionId: sessionId, email: email };
}

function logExit(sessionId, exitTimestamp, durationSec) {
  var sheet = _getOrCreateLogSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === sessionId && data[i][5] === 'Açık') {
      sheet.getRange(i + 1, 4).setValue(new Date(exitTimestamp));
      sheet.getRange(i + 1, 5).setValue(Math.round(durationSec / 60 * 10) / 10);
      sheet.getRange(i + 1, 6).setValue('Tamamlandı');
      break;
    }
  }
}

function _getOrCreateLogSheet() {
  var ss    = SpreadsheetApp.openById(FB_SHEET_ID);
  var sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet.appendRow(['Oturum ID', 'E-posta', 'Giriş Zamanı', 'Çıkış Zamanı', 'Süre (dk)', 'Durum']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  return sheet;
}
```

### Index.html — `</body>` öncesine eklenecek

```html
<script>
  var SESSION_ID = "<?= sessionId ?>";
  var TIMEOUT_MS = <?= timeoutMin ?> * 60 * 1000;

  var sessionStart = Date.now();
  var hiddenAt     = null;
  var exitLogged   = false;

  function _logExit(exitTimestamp, durationMs) {
    if (exitLogged) return;
    exitLogged = true;
    google.script.run.logExit(SESSION_ID, exitTimestamp, Math.round(durationMs / 1000));
  }

  function _newSession() {
    exitLogged = false;
    google.script.run
      .withSuccessHandler(function(result) {
        SESSION_ID   = result.sessionId;
        sessionStart = Date.now();
        hiddenAt     = null;
      })
      .createSession();
  }

  document.addEventListener('visibilitychange', function() {
    if (document.visibilityState === 'hidden') {
      hiddenAt = Date.now();
    } else {
      if (hiddenAt === null) return;
      var elapsed = Date.now() - hiddenAt;
      if (elapsed >= TIMEOUT_MS) {
        _logExit(hiddenAt, hiddenAt - sessionStart);
        _newSession();
      } else {
        hiddenAt = null;
      }
    }
  });

  window.addEventListener('beforeunload', function() {
    if (exitLogged) return;
    var exitAt = hiddenAt !== null ? hiddenAt : Date.now();
    _logExit(exitAt, exitAt - sessionStart);
  });
</script>
```


# Modül 3 — Versiyon Yönetimi

---

## Tetikleyici

Kullanıcı **"versiyon sistemini ekle"** veya **"beta banner ekle"** dediğinde aşağıdaki adımları sırayla uygula.

---

## Kurulum Adımları

### 1. Versiyon Motoru — kod.gs

`kod.gs` dosyasını aç. `getAppVersion` fonksiyonu var mı kontrol et.
- **Yoksa:** Modül 3 GAS bloğunu dosyanın en üstüne, ilk `function` tanımından önce ekle.
- **Varsa:** Dokunma.

### 2. kurulumTrigger Entegrasyonu — kod.gs

`kod.gs` dosyasında `kurulumTrigger` fonksiyonunu bul.
- **`_kurulumYap` çağrısı yoksa:** Mevcut `kurulumTrigger` fonksiyonunun gövdesini Modül 3 GAS bloğundaki `_kurulumYap` / `kurulumTrigger` / `majorDeploy` yapısıyla değiştir.
- **Varsa:** Dokunma.

### 3. Beta Banner — Index.html

`Index.html` dosyasını aç. `id="beta-banner"` elementi var mı kontrol et.
- **Yoksa:** Modül 3 HTML bloğunu `<body>` etiketinin hemen ardına, `<!-- SIDEBAR -->` yorumundan önce ekle.
- **Varsa:** Dokunma.

### 4. Sidebar Versiyon Rozeti — Index.html

`Index.html` dosyasında `id="version-badge"` elementi var mı kontrol et.
- **Yoksa:** Sidebar'ın `mt-auto` bölümündeki son `</div>` etiketinden önce Modül 3 Rozet bloğunu ekle. Ardından `</style>` etiketinden önce Modül 3 CSS bloğunu ekle.
- **Varsa:** Dokunma.

### 5. Bitti — Kullanıcıya bildir

> "Versiyon sistemi kuruldu. Şimdi GAS editöründe:
> - `kurulumTrigger` fonksiyonunu **bir kez çalıştır** — bu versiyonu `1.0` olarak kaydeder.
>
> Bundan sonra:
> - Küçük değişiklik deploy'larında → `kurulumTrigger` çalıştır
> - Büyük değişiklik deploy'larında → `majorDeploy` çalıştır
> - Beta bitti, yayına geçiyorsun → `setBeta` fonksiyonunu `false` parametresiyle çalıştır"

---

## Modül 3 — Versiyon Yönetimi

### kod.gs'e eklenecek

```javascript
// ─── Versiyon Yönetimi ────────────────────────────────────────────────────────

// Frontend'in sayfa açılışında çağırdığı fonksiyon
function getAppVersion() {
  var props = PropertiesService.getScriptProperties();
  return {
    version:   props.getProperty('APP_VERSION')    || '1.0',
    buildDate: props.getProperty('APP_BUILD_DATE') || '',
    isBeta:    props.getProperty('APP_IS_BETA')    !== 'false'
  };
}

// Versiyon numarasını artırır — doğrudan çağrılmaz, kurulumTrigger/majorDeploy üzerinden çalışır
function bumpVersion(type) {
  var props   = PropertiesService.getScriptProperties();
  var current = props.getProperty('APP_VERSION') || '1.0';
  var parts   = current.split('.').map(Number);
  if (type === 'major') { parts[0]++; parts[1] = 0; }
  else                  { parts[1]++; }
  var next      = parts[0] + '.' + parts[1];
  var buildDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  props.setProperty('APP_VERSION',    next);
  props.setProperty('APP_BUILD_DATE', buildDate);
  Logger.log('Versiyon: ' + next + ' (' + buildDate + ')');
  return { version: next, buildDate: buildDate };
}

// Beta modunu açar veya kapatır — bir kez çalıştırılır, kalıcıdır
function setBeta(flag) {
  PropertiesService.getScriptProperties().setProperty('APP_IS_BETA', flag ? 'true' : 'false');
  Logger.log('Beta modu: ' + (flag ? 'AÇIK' : 'KAPALI'));
}

// kurulumTrigger ve majorDeploy'un ortak kurulum mantığı
function _kurulumYap(bumpType) {
  bumpVersion(bumpType);
  // ... mevcut kurulumTrigger içeriği buraya taşınır (props set, trigger kurulum vb.)
}

// Küçük değişiklik deploy'unda çalıştır → örn. 1.0 → 1.1
function kurulumTrigger() { _kurulumYap('minor'); }

// Büyük değişiklik deploy'unda çalıştır → örn. 1.1 → 2.0
function majorDeploy()    { _kurulumYap('major'); }
```

> **Not:** `_kurulumYap` içine mevcut `kurulumTrigger` gövdesini (PropertiesService set, trigger kurulumu, vb.) taşı. `kurulumTrigger` ve `majorDeploy` bu gövdeyi yalnızca bump türünü değiştirerek çalıştırır.

---

### Index.html — `<body>` hemen ardına eklenecek (Beta Banner)

```html
<!-- BETA BANNER -->
<div id="beta-banner" class="fixed top-0 left-0 right-0 z-[9999] items-center justify-between px-6 py-2.5 bg-gradient-to-r from-[#041b3c] via-[#0052cc] to-[#041b3c] border-b border-[#76ff03]/30 shadow-lg" style="display:none;">
  <div class="flex items-center gap-3">
    <span class="text-[10px] font-extrabold tracking-widest text-[#041b3c] bg-[#76ff03] px-2 py-0.5 rounded-sm">BETA</span>
    <span class="text-white text-xs font-semibold">
      Task Tracking Portal
      <span id="bb-version" class="text-[#76ff03] font-bold"></span>
      <span class="text-slate-300 font-normal">— aktif geliştirme aşamasındadır. Geri bildiriminiz için</span>
      <button onclick="fbToggle();betaBannerKapat();" class="underline text-[#76ff03] hover:text-white transition-colors" style="background:none;border:none;cursor:pointer;font-size:inherit;font-family:inherit;padding:0;">Geri Bildirim</button>
      <span class="text-slate-300 font-normal">butonunu kullanın.</span>
    </span>
  </div>
  <button onclick="betaBannerKapat();" class="text-slate-400 hover:text-white transition-colors ml-4 flex-shrink-0" style="background:none;border:none;cursor:pointer;">
    <span class="material-symbols-outlined text-[18px]">close</span>
  </button>
</div>
<script>
  function betaBannerKapat() {
    var b = document.getElementById('beta-banner');
    var h = document.querySelector('header');
    var s = document.getElementById('sidebar');
    var m = document.getElementById('main-content');
    b.style.display = 'none';
    if(h) h.style.top = '';
    if(s) s.style.top = '';
    if(m) m.style.paddingTop = '';
  }

  function _applyVersionUI(info) {
    var b  = document.getElementById('beta-banner');
    var bv = document.getElementById('bb-version');
    if(bv) bv.textContent = 'v' + info.version;
    if(info.isBeta) {
      b.style.display = 'flex';
      var bh = b.offsetHeight || 40;
      var h = document.querySelector('header');
      var s = document.getElementById('sidebar');
      var m = document.getElementById('main-content');
      if(h) h.style.top = bh + 'px';
      if(s) s.style.top = bh + 'px';
      if(m) m.style.paddingTop = bh + 'px';
    }
    var vbv = document.getElementById('vb-version');
    var vbl = document.getElementById('vb-beta-label');
    if(vbv) vbv.textContent = 'v' + info.version + ' · ' + info.buildDate;
    if(vbl) { if(info.isBeta) vbl.classList.remove('hidden'); else vbl.classList.add('hidden'); }
  }

  document.addEventListener('DOMContentLoaded', function(){
    google.script.run
      .withSuccessHandler(_applyVersionUI)
      .withFailureHandler(function(){
        _applyVersionUI({ version:'1.0', buildDate:'', isBeta:true });
      })
      .getAppVersion();
  });
</script>
```

---

### Index.html — Sidebar `mt-auto` bölümüne eklenecek (Versiyon Rozeti)

```html
<!-- Version Badge -->
<div id="version-badge" class="sidebar-label mx-2 mt-2 px-3 py-2 rounded-lg bg-white/5 border border-white/10 flex items-center justify-between overflow-hidden transition-all duration-300">
  <div class="flex items-center gap-2">
    <span id="vb-beta-label" class="text-[9px] font-extrabold tracking-widest text-[#76ff03] uppercase bg-[#76ff03]/10 border border-[#76ff03]/30 px-1.5 py-0.5 rounded hidden">BETA</span>
    <span id="vb-version" class="text-[10px] text-slate-400 font-semibold whitespace-nowrap">...</span>
  </div>
  <span class="material-symbols-outlined text-[14px] text-slate-600">science</span>
</div>
```

---

### Index.html — `</style>` öncesine eklenecek (CSS)

```css
/* Sidebar collapsed: version badge gizle */
#sidebar.collapsed #version-badge { width:0; opacity:0; height:0; padding:0; margin:0; }
```


### Valeo Logosu eklemek.
<!-- Valeo Logo -->
<div id="sidebar-logo" class="px-4 py-2 mb-1 overflow-hidden transition-all duration-300">
  <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/2/2b/Valeo_Logo.svg/960px-Valeo_Logo.svg.png" 
       alt="Valeo" 
       class="w-full max-w-[160px] object-contain mx-auto block">
</div>

