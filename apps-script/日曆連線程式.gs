/**
 * 嘉特利日曆連線（Google Apps Script 網頁應用程式）　版本 3：可改標題、寫回姓名／地址／品項明細
 *
 * 讓桌面「嘉特利業績」資料夾：
 *   - 即時讀取「黃金單身坤-業務排程」日曆的行程
 *   - 把業績表裡「編輯」的修改寫回日曆（寫在行程說明欄的【嘉特利訂單】區塊）
 *
 * 只有帶著下面這把金鑰的請求才能讀寫，金鑰和桌面資料夾裡的 app\token.txt 相同。
 * 請不要把這段程式或金鑰傳給別人。
 */

const TOKEN = '請改成你自己的金鑰（一長串英數字，不要外流）';
const CAL_ID = '0978f50fe0c3a8293598bed4f4c3143db7c45aeb6d6e8ed3d50fc0b1843bc8c1@group.calendar.google.com';
const TZ = 'Asia/Taipei';
const TAG = '【嘉特利訂單】';
const OLD_TAGS = ['【嘉特利結帳】'];
const VERSION = 3;
const RESERVED = ['實收金額', '量測', '服務項目', '地區', '客戶類型', '不列入業績', '結帳備註', '最後修改'];

function cal_() {
  const c = CalendarApp.getCalendarById(CAL_ID);
  if (!c) throw new Error('找不到業務排程日曆，請用擁有這個日曆的帳號部署。');
  return c;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function day_(s, fallbackDays) {
  if (/^\d{4}-\d{2}-\d{2}$/.test(String(s || ''))) return new Date(s + 'T00:00:00+08:00');
  return new Date(Date.now() + fallbackDays * 86400000);
}

/** 讀取行程：?token=…&from=YYYY-MM-DD&to=YYYY-MM-DD，或 ?token=…&action=ping 測試連線 */
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.token !== TOKEN) return json_({ ok: false, error: '金鑰不正確' });
    const cal = cal_();
    if (p.action === 'ping') return json_({ ok: true, version: VERSION, calendar: cal.getName() });

    const from = day_(p.from, -730);
    const to = day_(p.to, 400);
    const events = cal.getEvents(from, to).map((ev) => {
      const allDay = ev.isAllDayEvent();
      const start = allDay ? ev.getAllDayStartDate() : ev.getStartTime();
      const stamp = Utilities.formatDate(start, TZ, "yyyy-MM-dd'T'HH:mm");
      return {
        id: ev.getId() + (ev.isRecurringEvent() ? '|' + stamp : ''),
        summary: ev.getTitle() || '',
        description: ev.getDescription() || '',
        day: stamp.slice(0, 10),
        time: allDay ? '' : stamp.slice(11, 16),
        allDay: allDay,
      };
    });
    return json_({ ok: true, version: VERSION, at: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), events: events });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/** 寫回修改：POST {token, id, summary?, changes:{items, area, old, amount, len, rate, memo, excluded, fields:{姓名, 電話, 縣市, 詳細地址, 明細}}} 或 {token, id, clear:true}
 *  summary 有給時會改行程標題。 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.token !== TOKEN) return json_({ ok: false, error: '金鑰不正確' });
    const id = String(body.id || '').split('|')[0];
    if (!id) return json_({ ok: false, error: '缺少訂單編號' });
    const ev = cal_().getEventById(id);
    if (!ev) return json_({ ok: false, error: '日曆上找不到這筆行程，可能已被刪除。' });

    const desc = ev.getDescription() || '';
    let cut = -1;
    [TAG].concat(OLD_TAGS).forEach((t) => { const i = desc.indexOf(t); if (i >= 0 && (cut < 0 || i < cut)) cut = i; });
    const kept = (cut >= 0 ? desc.slice(0, cut) : desc).replace(/\s+$/, '');

    if (body.clear) {
      ev.setDescription(kept);
      return json_({ ok: true, cleared: true });
    }

    const c = body.changes || {};
    const lines = [TAG];
    if (c.amount !== null && c.amount !== undefined && String(c.amount).trim() !== '') {
      const a = Math.round(Number(c.amount));
      if (!(a >= 0)) return json_({ ok: false, error: '實收金額不是數字' });
      lines.push('實收金額：' + a);
    }
    if (Number(c.len) > 0 && Number(c.rate) > 0) lines.push('量測：' + Number(c.len) + ' 公分 × ' + Number(c.rate) + ' 元');
    const items = (c.items || []).filter((x) => /^[\u4e00-\u9fff]{1,8}$/.test(String(x)));
    if (items.length) lines.push('服務項目：' + items.join('、'));
    if (c.area) lines.push('地區：' + String(c.area).slice(0, 20));
    if (c.old === '舊客戶' || c.old === '新客戶') lines.push('客戶類型：' + c.old);
    if (c.excluded) lines.push('不列入業績：是');
    if (c.memo) lines.push('結帳備註：' + String(c.memo).replace(/\s+/g, ' ').slice(0, 300));
    const f = c.fields || {};
    Object.keys(f).forEach((k) => {
      if (!/^[\u4e00-\u9fff]{1,8}$/.test(k) || RESERVED.indexOf(k) >= 0) return;
      const v = String(f[k] == null ? '' : f[k]).replace(/\s+/g, ' ').trim().slice(0, 500);
      if (v) lines.push(k + '：' + v);
    });
    lines.push('最後修改：' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') + '（桌面業績表）');

    if (typeof body.summary === 'string' && body.summary.trim()) ev.setTitle(body.summary.slice(0, 1000));
    ev.setDescription((kept ? kept + '\n\n' : '') + lines.join('\n'));
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/** 部署前可以先按「執行」這個函式，確認帳號看得到日曆並完成授權。 */
function testConnection() {
  const cal = cal_();
  const n = cal.getEvents(new Date(Date.now() - 30 * 86400000), new Date(Date.now() + 30 * 86400000)).length;
  Logger.log('連線成功：' + cal.getName() + '，前後 30 天共 ' + n + ' 筆行程。');
}
