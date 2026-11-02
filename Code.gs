/**
 * بيت الصنعة - الوسيط (Google Apps Script)
 * 1) أنشئ Google Sheet، افتح Extensions > Apps Script والصق هذا الكود
 * 2) شغّل الدالة setup() مرة واحدة (وافق على الصلاحيات)
 * 3) Project Settings > Script Properties: أضف ADMIN_KEY بكلمة سر قوية
 * 4) Deploy > New deployment > Web app: Execute as Me، Access: Anyone
 * 5) انسخ رابط /exec وضعه في CFG.API داخل index.html
 */
const SHEETS = {
  Products: ['id','name','cat','desc','price','old','sold','hidden','sizes','colors'],
  Inbox_Orders: ['id','date','name','phone','gov','address','notes','items','subtotal','shipping','total','source','sync'],
  Inbox_Messages: ['date','phone','text','sync'],
  Users: ['phone','name','salt','hash','token','fails','lockUntil','created'],
  Status: ['id','phone','status'],
  OTP: ['phone','hash','salt','exp','tries','email','reqs','reqStart']
};
// تُرسل تقارير الطلبات والعملاء والحالات والمخزون تلقائياً إلى هذين العنوانين.
const REPORT_EMAILS = ['mg8599287@gmail.com', 'nsr09083@gmail.com'];
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(SHEETS).forEach(n => {
    const sh = ss.getSheetByName(n) || ss.insertSheet(n);
    if (sh.getLastRow() === 0) sh.appendRow(SHEETS[n]);
    const pc = SHEETS[n].indexOf('phone');
    if (pc > -1) sh.getRange(1, pc + 1, sh.getMaxRows(), 1).setNumberFormat('@'); // نص حتى لا يُحذف الصفر
  });
}
// الهاتف قد يتحول إلى رقم فيفقد الصفر الأول، هذه الدالة تعيده
const ph_ = v => { v = String(v); return v.length === 10 ? '0' + v : v; };
const sh_ = n => SpreadsheetApp.getActive().getSheetByName(n);
const rows_ = n => { const v = sh_(n).getDataRange().getValues(), h = v.shift(); return v.map((r,i) => { const o = {_row:i+2}; h.forEach((k,j) => o[k] = r[j]); return o; }); };
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const prop_ = k => PropertiesService.getScriptProperties().getProperty(k);
function hash_(pass, salt) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + pass));
}

function doGet(e) {
  if (e.parameter.action === 'products') {
    const p = rows_('Products').filter(r => r.hidden !== true && r.hidden !== 'TRUE')
      .map(r => ({id:+r.id, name:r.name, cat:r.cat, desc:r.desc, price:+r.price, old:+r.old, sold:+r.sold,
                  sizes:String(r.sizes).split('|'), colors:JSON.parse(r.colors)}));
    return out_({ok:true, products:p});
  }
  return out_({ok:false});
}

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const fn = {register, login, createOrder, myOrders, adminPull, adminAck, adminPublish, adminStatus, adminResetPass, adminUpload, requestOtp, track}[d.action];
    return out_(fn ? fn(d) : {ok:false, error:'طلب غير معروف'});
  } catch (err) { return out_({ok:false, error: err.message === 'denied' ? 'مفتاح الأدمن غير صحيح' : 'خطأ في الخادم: ' + err.message}); }
  finally { lock.releaseLock(); }
}

function register(d) {
  if (!/^01[0125]\d{8}$/.test(d.phone) || !d.pass || d.pass.length < 6 || !d.name) return {ok:false, error:'بيانات غير صحيحة'};
  if (rows_('Users').some(u => ph_(u.phone) === d.phone)) return {ok:false, error:'الرقم مسجّل من قبل'};
  const salt = Utilities.getUuid(), token = Utilities.getUuid();
  sh_('Users').appendRow([d.phone, String(d.name).slice(0,60), salt, hash_(d.pass, salt), token, 0, 0, new Date()]);
  sendReport_('تسجيل عميل جديد', 'تم إنشاء حساب عميل جديد في متجر دفيّ', [
    ['الاسم', String(d.name).slice(0,60)],
    ['رقم الهاتف', d.phone],
    ['تاريخ التسجيل', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm')]
  ]);
  return {ok:true, user:{name:d.name, phone:d.phone, token}};
}
function login(d) {
  const u = rows_('Users').find(r => ph_(r.phone) === d.phone);
  if (!u) return {ok:false, error:'الرقم أو كلمة السر غير صحيحة'};
  if (Number(u.lockUntil) > Date.now()) return {ok:false, error:'الحساب مقفل مؤقتاً، حاول بعد قليل'};
  const s = sh_('Users');
  if (hash_(d.pass, u.salt) !== u.hash) {
    const f = Number(u.fails) + 1;
    s.getRange(u._row, 6, 1, 2).setValues([[f >= 5 ? 0 : f, f >= 5 ? Date.now() + 15*60000 : 0]]);
    return {ok:false, error:'الرقم أو كلمة السر غير صحيحة'};
  }
  const token = Utilities.getUuid();
  s.getRange(u._row, 5, 1, 3).setValues([[token, 0, 0]]);
  return {ok:true, user:{name:u.name, phone:ph_(u.phone), token}};
}
function createOrder(d) {
  const o = d.order || {};
  if (!o.name || !/^01[0125]\d{8}$/.test(o.phone) || !o.address || !o.items || !o.items.length) return {ok:false, error:'بيانات الطلب ناقصة'};
  if (!checkOtp_(o.phone, d.otp)) return {ok:false, error:'رمز التحقق غير صحيح أو منتهي'};
  const recent = rows_('Inbox_Orders').filter(r => ph_(r.phone) === o.phone && new Date() - new Date(r.date) < 3600000);
  if (recent.length >= 3) return {ok:false, error:'تم تجاوز الحد الأقصى للطلبات، تواصل معنا واتساب'};
  // إعادة حساب الأسعار من الشيت لمنع التلاعب
  const prods = rows_('Products'); let sub = 0;
  const items = o.items.map(i => { const p = prods.find(x => +x.id === +i.id); if (!p) throw 1; const q = Math.min(99, Math.max(1, +i.qty|0)); sub += p.price * q; return {id:+p.id, name:p.name, price:+p.price, qty:q, ci:+i.ci, color:String(i.color||''), size:String(i.size||'')}; });
  const ship = Math.max(0, +o.shipping || 0), total = sub + ship;
  const id = 'A' + Utilities.formatDate(new Date(), 'GMT+2', 'yyMMddHHmmss').slice(-7);
  sh_('Inbox_Orders').appendRow([id, new Date(), o.name, o.phone, o.gov, o.address, o.notes || '', JSON.stringify(items), sub, ship, total, o.source || 'web', 'pending']);
  sh_('Status').appendRow([id, o.phone, 'جديد']);
  sendReport_('طلب جديد ' + id, 'تم استلام طلب جديد في متجر دفيّ', [
    ['رقم الطلب', id],
    ['العميل', o.name],
    ['رقم الهاتف', o.phone],
    ['المحافظة', o.gov || 'غير محددة'],
    ['العنوان', o.address],
    ['المنتجات', items.map(i => i.name + ' (' + i.color + ' / ' + i.size + ') × ' + i.qty).join('\n')],
    ['قيمة المنتجات', sub + ' ج'],
    ['الشحن', ship + ' ج'],
    ['الإجمالي', total + ' ج'],
    ['طريقة الطلب', o.source || 'web']
  ]);
  notify_('طلب جديد ' + id + ' - ' + o.name + ' - ' + total + ' ج');
  return {ok:true, id};
}
function myOrders(d) {
  const u = rows_('Users').find(r => r.token && r.token === d.token);
  if (!u) return {ok:false, error:'سجّل الدخول من جديد'};
  const st = rows_('Status').filter(r => ph_(r.phone) === ph_(u.phone));
  const orders = rows_('Inbox_Orders').filter(r => ph_(r.phone) === ph_(u.phone)).reverse().map(r => ({
    id:r.id, date:Utilities.formatDate(new Date(r.date), 'GMT+2', 'yyyy/MM/dd'), total:r.total,
    items:JSON.parse(r.items), status:(st.find(s => s.id === r.id) || {}).status || 'جديد'}));
  return {ok:true, orders};
}

/* ===== واجهة الأدمن (تتطلب ADMIN_KEY) ===== */
function admin_(d) { if (!prop_('ADMIN_KEY') || d.key !== prop_('ADMIN_KEY')) throw new Error('denied'); }
function adminPull(d) { admin_(d);
  return {ok:true, orders:rows_('Inbox_Orders').filter(r => r.sync === 'pending'),
          users:rows_('Users').map(u => ({phone:ph_(u.phone), name:u.name, created:u.created}))}; }
function adminAck(d) { admin_(d); const s = sh_('Inbox_Orders');
  rows_('Inbox_Orders').filter(r => (d.ids||[]).indexOf(r.id) > -1).forEach(r => s.getRange(r._row, 13).setValue('synced')); return {ok:true}; }
function adminPublish(d) { admin_(d);
  const products = d.products || [];
  const stockLines = products.map(p => {
    const colors = typeof p.colors === 'string' ? JSON.parse(p.colors || '[]') : (p.colors || []);
    const stock = colors.reduce((sum, c) => sum + (c.stock || []).reduce((n, q) => n + (Number(q) || 0), 0), 0);
    return p.name + ': ' + (p.hidden ? 'مخفي' : 'منشور') + ' — ' + stock + ' قطعة متاحة';
  });
  const s = sh_('Products'); s.clearContents(); s.appendRow(SHEETS.Products);
  products.forEach(p => s.appendRow(SHEETS.Products.map(k => p[k] === undefined ? '' : p[k])));
  sendReport_('تحديث حالة المنتجات', 'تم نشر تحديثات المنتجات والمخزون', [
    ['عدد المنتجات', products.length],
    ['المنتجات والمخزون', stockLines.join('\n') || 'لا توجد منتجات']
  ]);
  return {ok:true}; }
function adminStatus(d) { admin_(d);
  const statusRow = rows_('Status').find(r => r.id === d.id);
  if (!statusRow) return {ok:false, error:'الطلب غير موجود'};
  const previous = statusRow.status;
  sh_('Status').getRange(statusRow._row, 3).setValue(d.status);
  if (previous !== d.status) {
    const order = rows_('Inbox_Orders').find(r => r.id === d.id);
    sendReport_('تحديث حالة الطلب ' + d.id, 'تم تغيير حالة طلب في متجر دفيّ', [
      ['رقم الطلب', d.id],
      ['العميل', order ? order.name : 'غير متاح'],
      ['رقم الهاتف', order ? ph_(order.phone) : ph_(statusRow.phone)],
      ['الحالة السابقة', previous],
      ['الحالة الجديدة', d.status]
    ]);
  }
  return {ok:true}; }
function adminResetPass(d) { admin_(d); const u = rows_('Users').find(r => ph_(r.phone) === d.phone); if (!u) return {ok:false};
  const salt = Utilities.getUuid(); sh_('Users').getRange(u._row, 3, 1, 5).setValues([[salt, hash_(d.newPass, salt), '', 0, 0]]); return {ok:true}; }

function notify_(msg) {
  const mail = prop_('NOTIFY_EMAIL');
  if (mail) {
    try { MailApp.sendEmail(mail, 'بيت الصنعة', msg); }
    catch (err) { Logger.log('تعذر إرسال الإشعار إلى ' + mail + ': ' + err.message); }
  }
  const t = prop_('TG_TOKEN'), c = prop_('TG_CHAT');
  if (t && c) {
    try { UrlFetchApp.fetch('https://api.telegram.org/bot' + t + '/sendMessage', {method:'post', payload:{chat_id:c, text:msg}, muteHttpExceptions:true}); }
    catch (err) { Logger.log('تعذر إرسال إشعار تيليجرام: ' + err.message); }
  }
}
const emailHtml_ = value => String(value == null ? '' : value)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
function sendReport_(subject, heading, details) {
  const rows = details.map(([label, value]) => {
    const text = String(value == null ? '' : value);
    return `<tr><td style="padding:10px 12px;border-bottom:1px solid #E1E8F1;color:#69758A;vertical-align:top;white-space:nowrap">${emailHtml_(label)}</td><td style="padding:10px 12px;border-bottom:1px solid #E1E8F1;color:#1B2433;line-height:1.8;white-space:pre-line">${emailHtml_(text)}</td></tr>`;
  }).join('');
  const body = heading + '\n\n' + details.map(([label, value]) => label + ': ' + String(value == null ? '' : value)).join('\n');
  const htmlBody = `<div dir="rtl" lang="ar" style="margin:0;padding:28px 12px;background:#F5F8FC;font-family:Tahoma,Arial,sans-serif;color:#1B2433">
    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="max-width:620px;margin:0 auto;background:#FFFFFF;border:1px solid #E1E8F1;border-radius:18px;overflow:hidden">
      <tr><td style="height:6px;background:#E8758F;font-size:0;line-height:0">&nbsp;</td></tr>
      <tr><td style="padding:24px 26px 8px;text-align:center"><div style="font-size:28px;font-weight:bold;color:#24406B">دفيّ</div><div style="margin-top:4px;font-size:12px;color:#69758A">تقارير المتجر</div></td></tr>
      <tr><td style="padding:12px 26px 24px"><h1 style="margin:0 0 16px;text-align:center;font-size:20px;line-height:1.6;color:#16294A">${emailHtml_(heading)}</h1>
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border:1px solid #E1E8F1;border-radius:12px;border-spacing:0">${rows}</table>
      </td></tr>
      <tr><td style="padding:14px 20px;border-top:1px solid #E1E8F1;text-align:center;color:#69758A;font-size:12px">إشعار آلي من متجر دفيّ</td></tr>
    </table>
  </div>`;
  try {
    MailApp.sendEmail({
      to: REPORT_EMAILS.join(','),
      subject: 'دفيّ | ' + subject,
      body,
      htmlBody,
      name: 'دفيّ'
    });
  } catch (err) {
    Logger.log('تعذر إرسال التقرير "' + subject + '" إلى ' + REPORT_EMAILS.join(',') + ': ' + err.message);
  }
}

/* رفع صور المنتجات إلى Google Drive (مجلد dafi-images) وإرجاع رابط عام */
function adminUpload(d) { admin_(d);
  const it = DriveApp.getFoldersByName('dafi-images');
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder('dafi-images');
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(d.data), 'image/jpeg', 'p' + Date.now() + '.jpg'));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {ok:true, url:'https://lh3.googleusercontent.com/d/' + file.getId()};
}
/* شغّلها مرة واحدة من المحرر لمنح صلاحية Drive */
function authorizeDrive() { DriveApp.getRootFolder(); }

/* ===== رمز التحقق OTP (يُرسل بالبريد ويُحفظ مشفّراً في ورقة OTP) ===== */
function otpEmailHtml_(code) {
  return `<div dir="rtl" lang="ar" style="margin:0;padding:32px 12px;background:#F5F8FC;font-family:Tahoma,Arial,sans-serif;color:#1B2433">
    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="max-width:560px;margin:0 auto;background:#FFFFFF;border:1px solid #E1E8F1;border-radius:20px;overflow:hidden">
      <tr><td style="height:6px;background:#E8758F;font-size:0;line-height:0">&nbsp;</td></tr>
      <tr><td style="padding:28px 30px 12px;text-align:center">
        <div style="font-size:30px;font-weight:bold;line-height:1.4;color:#24406B">دفيّ</div>
        <div style="margin-top:4px;font-size:13px;color:#69758A">دفء ناعم لكل خطوة صغيرة</div>
      </td></tr>
      <tr><td style="padding:16px 30px 30px;text-align:center">
        <div style="display:inline-block;padding:6px 14px;border-radius:20px;background:#FCE6EB;color:#B94D68;font-size:12px;font-weight:bold">تأكيد الطلب</div>
        <h1 style="margin:16px 0 8px;font-size:22px;line-height:1.5;color:#16294A">رمز التحقق الخاص بك</h1>
        <p style="margin:0;color:#69758A;font-size:15px;line-height:1.9">استخدم الرمز التالي لتأكيد طلبك. هذا الرمز صالح لمدة <strong style="color:#24406B">5 دقائق</strong>.</p>
        <div dir="ltr" style="margin:22px auto;padding:16px 12px;max-width:280px;border:1px dashed #E8758F;border-radius:14px;background:#FFF8FA;color:#24406B;font-family:Arial,sans-serif;font-size:34px;font-weight:bold;letter-spacing:10px">${code}</div>
        <p style="margin:0;color:#69758A;font-size:13px;line-height:1.9">لا تشارك هذا الرمز مع أي شخص. فريق دفيّ لن يطلبه منك.</p>
      </td></tr>
      <tr><td style="padding:16px 24px;border-top:1px solid #E1E8F1;text-align:center;color:#69758A;font-size:12px;line-height:1.8">
        إذا لم تطلب هذا الرمز، يمكنك تجاهل الرسالة بأمان.<br>
        © دفيّ · شباشب مخملية من المصنع إلى بيتك
      </td></tr>
    </table>
  </div>`;
}
function requestOtp(d) {
  if (!/^01[0125]\d{8}$/.test(d.phone) || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(d.email || '')) return {ok:false, error:'أدخل رقم هاتف وبريداً صحيحين'};
  if (prop_('OTP') === 'off') return {ok:true, skip:true};
  const s = sh_('OTP'), row = rows_('OTP').find(r => ph_(r.phone) === d.phone), now = Date.now();
  let reqs = 1, start = now;
  if (row) {
    if (now - row.exp + 300000 < 60000 && row.exp > 0) return {ok:false, error:'انتظر دقيقة قبل طلب رمز جديد'};
    if (now - row.reqStart < 3600000) { if (row.reqs >= 4) return {ok:false, error:'تجاوزت عدد المحاولات، حاول بعد ساعة'}; reqs = row.reqs + 1; start = row.reqStart; }
  }
  const code = String(parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 900000 + 100000), salt = Utilities.getUuid();
  const vals = [d.phone, hash_(code, salt), salt, now + 300000, 0, d.email, reqs, start];
  if (row) s.getRange(row._row, 1, 1, 8).setValues([vals]); else s.appendRow(vals);
  MailApp.sendEmail({
    to: d.email,
    subject: 'رمز تأكيد طلبك من دفيّ',
    body: 'رمز تأكيد طلبك من دفيّ: ' + code + '\nصالح لمدة 5 دقائق. لا تشاركه مع أي شخص.',
    htmlBody: otpEmailHtml_(code),
    name: 'دفيّ'
  });
  return {ok:true};
}
function checkOtp_(phone, code) {
  if (prop_('OTP') === 'off') return true;
  const row = rows_('OTP').find(r => ph_(r.phone) === phone);
  if (!row || Date.now() > row.exp || row.tries >= 5) return false;
  const s = sh_('OTP');
  if (hash_(String(code || ''), row.salt) !== row.hash) { s.getRange(row._row, 5).setValue(row.tries + 1); return false; }
  s.getRange(row._row, 4, 1, 2).setValues([[0, 0]]); // الرمز يُستخدم مرة واحدة
  return true;
}
/* تتبّع الطلبات: يتطلب تطابق رقم الهاتف مع رقم الطلب */
function track(d) {
  const ids = (d.ids || []).slice(0, 30), phone = String(d.phone || ''), st = rows_('Status');
  return {ok:true, orders:rows_('Inbox_Orders').filter(r => ids.indexOf(r.id) > -1 && ph_(r.phone) === phone).map(r => ({
    id:r.id, total:r.total, date:Utilities.formatDate(new Date(r.date), 'GMT+2', 'yyyy/MM/dd'), items:JSON.parse(r.items),
    status:(st.find(x => x.id === r.id) || {}).status || 'جديد'}))};
}
/* شغّلها مرة واحدة لمنح صلاحيتي Drive والبريد */
function authorizeAll() { DriveApp.getRootFolder(); MailApp.getRemainingDailyQuota(); }
