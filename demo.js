/**
 * デモモード（ローカル確認用）
 * config.js で API_URL = 'demo' のときだけ動きます。
 * Code.gs を同じフォルダから読み込み、Google のサービスをブラウザ内で模擬して実行します。
 * - データ：このブラウザの localStorage に保存（他の人・他のPCとは共有されません）
 * - メール：実際には送信せず、画面右下の「送信メール」に表示
 * - 管理者パスワード：demo1234
 */
(function () {
  if (window.API_URL !== 'demo') return;

  const TZ = 'Asia/Tokyo';
  const K = { sheet: 'demo_sheet', props: 'demo_props', mails: 'demo_mails', cache: 'demo_cache' };
  const LS = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (_) { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (_) {} },
  };

  let data = [], props = {}, cache = {};
  const load = () => { data = LS.get(K.sheet, []); props = LS.get(K.props, {}); cache = LS.get(K.cache, {}); };
  const save = () => { LS.set(K.sheet, data); LS.set(K.props, props); LS.set(K.cache, cache); };

  function formatDate(d, tz, f) {
    const p = {};
    new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', hourCycle: 'h23', weekday: 'short' })
      .formatToParts(d).forEach(x => p[x.type] = x.value);
    const u = { Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6, Sun: 7 }[p.weekday];
    return f.replace('yyyy', p.year).replace('MM', p.month).replace('dd', p.day)
            .replace('HH', p.hour).replace('mm', p.minute).replace('u', u);
  }
  function uuid() {
    const b = new Uint8Array(16); crypto.getRandomValues(b);
    b[6] = (b[6] & 0x0f) | 0x40; b[8] = (b[8] & 0x3f) | 0x80;
    const h = [...b].map(x => x.toString(16).padStart(2, '0')).join('');
    return `${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}`;
  }

  // ---- スプレッドシートの模擬 ----
  const range = (r, c, nr = 1, nc = 1) => ({
    setValues(v) {
      for (let i = 0; i < nr; i++) {
        data[r - 1 + i] = data[r - 1 + i] || [];
        for (let j = 0; j < nc; j++) {
          const x = v[i][j];
          data[r - 1 + i][c - 1 + j] = x instanceof Date ? formatDate(x, TZ, 'yyyy/MM/dd HH:mm') : x;
        }
      }
      return this;
    },
    getValues() {
      const o = [];
      for (let i = 0; i < nr; i++) {
        const row = [];
        for (let j = 0; j < nc; j++) {
          const x = (data[r - 1 + i] || [])[c - 1 + j];
          row.push(x == null ? '' : x);
        }
        o.push(row);
      }
      return o;
    },
    clearContent() {
      for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) if (data[r - 1 + i]) data[r - 1 + i][c - 1 + j] = '';
      return this;
    },
    setNumberFormat() { return this; }, setFontWeight() { return this; }, setBackground() { return this; },
  });
  const sheet = {
    getRange(r, c, nr, nc) { return typeof r === 'string' ? range(1, 1) : range(r, c, nr, nc); },
    getLastRow() { return data.length; },
    deleteRow(r) { data.splice(r - 1, 1); },
    setFrozenRows() {},
  };

  window.SpreadsheetApp = {
    getActive: () => ({ getSheetByName: () => (data.length ? sheet : null), insertSheet: () => sheet, getSpreadsheetTimeZone: () => TZ }),
    flush() {},
  };
  window.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) };
  window.PropertiesService = { getScriptProperties: () => ({
    getProperty: k => (k in props ? props[k] : null),
    setProperty: (k, v) => { props[k] = String(v); },
    getProperties: () => Object.assign({}, props),
    setProperties: o => { for (const k in o) props[k] = String(o[k]); },
  }) };
  window.CacheService = { getScriptCache: () => ({
    get: k => (cache[k] && cache[k].exp > Date.now() ? cache[k].v : null),
    put: (k, v, s) => { cache[k] = { v, exp: Date.now() + (s || 600) * 1000 }; },
  }) };
  window.MailApp = {
    sendEmail(o) {
      const m = LS.get(K.mails, []);
      m.unshift(Object.assign({ time: formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm') }, o));
      LS.set(K.mails, m.slice(0, 200));
      setTimeout(updateBadge, 0);
    },
    getRemainingDailyQuota: () => 100,
  };
  window.ContentService = { MimeType: { JSON: 'json' }, createTextOutput: t => ({ t, setMimeType() { return this; } }) };
  window.Utilities = { getUuid: uuid, sleep() {}, formatDate };

  // ---- Code.gs を読み込んで実行 ----
  const ready = new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = 'Code.gs';
    s.onload = () => { CONFIG.ADMIN_PASSWORD = 'demo1234'; resolve(); }; // eslint-disable-line no-undef
    s.onerror = () => reject(new Error('Code.gs を読み込めませんでした。index.html と同じフォルダに置いてください。'));
    document.head.appendChild(s);
  });

  window.DEMO = {
    async post(obj) {
      await ready;
      await new Promise(r => setTimeout(r, 250)); // 通信しているような待ち時間
      load();
      try { return JSON.parse(doPost({ postData: { contents: JSON.stringify(obj) } }).t); } // eslint-disable-line no-undef
      finally { save(); }
    },
  };

  // ---- 画面：デモ表示・送信メール一覧 ----
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const linkify = s => esc(s).replace(/((?:https?|file):\/\/[^\s<]+)/g, '<a href="$1">$1</a>');
  let badge, panel;

  function updateBadge() {
    if (badge) badge.textContent = `📧 送信メール（${LS.get(K.mails, []).length}）`;
    if (panel && !panel.hidden) renderPanel();
  }
  function renderPanel() {
    const mails = LS.get(K.mails, []);
    panel.querySelector('.dm-list').innerHTML = mails.length ? mails.map(m => `
      <div class="dm-mail">
        <div class="dm-meta">${esc(m.time)}　宛先：${esc(m.to)}${m.bcc ? `（BCC：${esc(m.bcc)}）` : ''}</div>
        <div class="dm-subj">${esc(m.subject)}</div>
        <pre>${linkify(m.body)}</pre>
      </div>`).join('') : '<p class="dm-empty">まだメールはありません。</p>';
  }

  function mountUi() {
    const css = document.createElement('style');
    css.textContent = `
      .dm-banner { background:#fff4ce; color:#5c4400; font-size:.85rem; padding:8px 16px; text-align:center; }
      @media (prefers-color-scheme: dark) { .dm-banner { background:#3a3115; color:#f3dc9a; } }
      .dm-badge { position:fixed; right:16px; bottom:16px; z-index:50; padding:10px 14px; border-radius:999px; border:0;
        background:var(--accent,#1f6feb); color:#fff; font:inherit; font-weight:600; box-shadow:0 2px 8px rgba(0,0,0,.25); cursor:pointer; }
      .dm-panel { position:fixed; right:16px; bottom:68px; z-index:50; width:min(440px, calc(100vw - 32px)); max-height:70vh;
        display:flex; flex-direction:column; background:var(--card,#fff); color:var(--text,#222); border:1px solid var(--line,#ddd);
        border-radius:12px; box-shadow:0 6px 24px rgba(0,0,0,.25); }
      .dm-head { display:flex; justify-content:space-between; align-items:center; padding:10px 14px; border-bottom:1px solid var(--line,#ddd); font-weight:600; }
      .dm-head button { background:none; border:0; color:var(--accent,#1f6feb); font:inherit; cursor:pointer; font-size:.85rem; }
      .dm-list { overflow-y:auto; padding:8px 14px; }
      .dm-mail { border-bottom:1px solid var(--line,#ddd); padding:10px 0; }
      .dm-meta { font-size:.75rem; color:var(--muted,#777); }
      .dm-subj { font-weight:600; margin:2px 0 6px; font-size:.9rem; }
      .dm-mail pre { white-space:pre-wrap; word-break:break-all; font:inherit; font-size:.82rem; margin:0; }
      .dm-empty { color:var(--muted,#777); font-size:.9rem; }
      [hidden] { display:none !important; }`;
    document.head.appendChild(css);

    const banner = document.createElement('div');
    banner.className = 'dm-banner';
    banner.textContent = 'デモモード：データはこのブラウザ内だけに保存され、メールは実際には送信されません（右下で文面を確認できます）';
    document.body.prepend(banner);

    badge = document.createElement('button');
    badge.className = 'dm-badge'; badge.type = 'button';
    document.body.appendChild(badge);

    panel = document.createElement('div');
    panel.className = 'dm-panel'; panel.hidden = true;
    panel.innerHTML = `<div class="dm-head"><span>送信されたメール（デモ）</span>
      <span><button type="button" class="dm-reset">デモデータを全消去</button>　<button type="button" class="dm-close">閉じる</button></span></div>
      <div class="dm-list"></div>`;
    document.body.appendChild(panel);

    badge.addEventListener('click', () => { panel.hidden = !panel.hidden; if (!panel.hidden) renderPanel(); });
    panel.querySelector('.dm-close').addEventListener('click', () => { panel.hidden = true; });
    panel.querySelector('.dm-reset').addEventListener('click', () => {
      if (!confirm('デモの枠・予約・設定・メールをすべて消去します。よろしいですか？')) return;
      Object.values(K).forEach(k => { try { localStorage.removeItem(k); } catch (_) {} });
      location.href = location.pathname;
    });
    window.addEventListener('storage', updateBadge); // 別タブで予約したときも件数を更新

    const lf = document.getElementById('loginForm');
    if (lf) {
      const p = document.createElement('p');
      p.style.cssText = 'font-size:.85rem;margin:12px 0 0;color:var(--muted)';
      p.textContent = 'デモモードの管理者パスワード：demo1234';
      lf.appendChild(p);
    }
    updateBadge();
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', mountUi);
  else mountUi();
})();
