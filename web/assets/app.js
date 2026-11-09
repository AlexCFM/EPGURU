(() => {
  const $ = (s) => document.querySelector(s);
  const state = { mode: 'epg', channel: '', busy: false, jobs: [] };
  const EXT = ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xls', '.xml', '.ods', '.csv', '.tsv', '.txt', '.htm', '.html'];

  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  async function api(path, opts = {}) {
    const r = await fetch(path, { credentials: 'same-origin', ...opts });
    if (r.status === 401) { location.href = '/login'; throw new Error('401'); }
    if (!r.ok) {
      let msg = r.statusText;
      try { msg = (await r.json()).detail || msg; } catch (_) {}
      throw new Error(msg);
    }
    return r.json();
  }

  // ---------- каналы ----------
  async function loadChannels() {
    const cfg = await api('/api/config');
    const list = [{ code: '', name: 'АВТО', template: 'ПО ИМЕНИ ФАЙЛА' }, ...cfg.channels];
    $('#chans').innerHTML = list.map((c) => `
      <button type="button" class="key ckey${c.code === state.channel ? ' on' : ''}" data-code="${esc(c.code)}"
        data-tpl="${esc(c.code ? c.template : 'АВТО')}" aria-pressed="${c.code === state.channel}">
        <span class="slot" aria-hidden="true"></span><b>${esc(c.name)}</b><small>${esc(c.template)}</small>
      </button>`).join('');
    $('#chans').querySelectorAll('.ckey').forEach((b) => b.addEventListener('click', () => {
      state.channel = b.dataset.code;
      $('#chans').querySelectorAll('.ckey').forEach((x) => {
        const on = x === b;
        x.classList.toggle('on', on);
        x.setAttribute('aria-pressed', on);
      });
      $('#lcd-tpl').textContent = $('#c-tpl').textContent = b.dataset.tpl;
    }));
  }

  // ---------- режимы ----------
  document.querySelectorAll('[data-mode]').forEach((b) => b.addEventListener('click', () => {
    state.mode = b.dataset.mode;
    document.querySelectorAll('[data-mode]').forEach((x) => {
      const on = x === b;
      x.classList.toggle('on', on);
      x.setAttribute('aria-pressed', on);
    });
    const dates = state.mode === 'dates';
    $('#c-mode').textContent = dates ? 'ДАТЫ' : 'EPG';
    $('#lcd-title').textContent = dates ? 'ПОЧИНКА ДАТ' : 'ПРИЁМ EPG';
    $('#drop-big').textContent = dates ? 'ПЕРЕТАЩИ XLSX С ДАТАМИ' : 'ПЕРЕТАЩИ EPG СЮДА';
    $('#drop-small').textContent = dates
      ? 'Excel, CSV, XML: даты текстом, числом, с перепутанными днём и месяцем → DD/MM/YYYY'
      : 'XLSX, XLS, CSV, XML, ODS — можно несколько сразу';
    $('#chans').closest('div').style.opacity = dates ? '.45' : '1';
  }));

  // ---------- каталоги ----------
  async function loadCatalogs() {
    const list = await api('/api/catalogs');
    $('#cat-list').innerHTML = list.map((c) => {
      const when = c.uploaded ? new Date(c.uploaded * 1000).toLocaleString('ru-RU', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '';
      const info = c.loaded
        ? `<span class="f">${esc(c.file || '')}</span><span class="label" style="letter-spacing:.04em">${esc(when)} · фильмов ${c.films ?? '—'} · сериалов ${c.series ?? '—'}${c.other ? ' · прочее ' + c.other : ''}</span>`
        : '<span class="label" style="letter-spacing:.04em">каталог не загружен</span>';
      return `<div class="cat">
          <span class="dot${c.loaded ? ' on' : ''}" aria-hidden="true"></span>
          <div class="cat-name"><b>${esc(c.name)}</b><span class="label">${esc(c.label || '')}</span></div>
          <div class="cat-info">${info}</div>
          <label class="key fkey cat-up">ЗАГРУЗИТЬ<input type="file" class="vh" data-code="${esc(c.code)}"
            accept=".xlsx,.xlsm,.xls,.xml,.ods,.csv,.tsv,.txt,.htm,.html"></label>
        </div>`;
    }).join('') || '<div class="empty">Для каналов не настроены каталоги.</div>';
    $('#cat-list').querySelectorAll('input[type=file]').forEach((inp) => inp.addEventListener('change', async () => {
      const f = inp.files[0];
      if (!f) return;
      const fd = new FormData();
      fd.append('file', f);
      foot('ЗАГРУЗКА КАТАЛОГА…');
      try {
        const r = await api(`/api/catalogs/${encodeURIComponent(inp.dataset.code)}`, { method: 'POST', body: fd });
        foot(`КАТАЛОГ ${inp.dataset.code}: ФИЛЬМОВ ${r.films}, СЕРИАЛОВ ${r.series}`);
      } catch (e) {
        if (e.message !== '401') { foot('КАТАЛОГ: ' + e.message.toUpperCase()); flash(); }
      }
      loadCatalogs();
    }));
  }
  $('#k-catalog').addEventListener('click', () => {
    const sec = $('#catalogs');
    sec.hidden = !sec.hidden;
    $('#k-catalog').classList.toggle('on', !sec.hidden);
    if (!sec.hidden) { loadCatalogs(); sec.scrollIntoView({ behavior: 'smooth' }); }
  });

  $('#k-history').addEventListener('click', () => $('#history').scrollIntoView({ behavior: 'smooth' }));
  $('#k-logout').addEventListener('click', async () => {
    await fetch('/api/logout', { method: 'POST' });
    location.href = '/login';
  });

  // ---------- загрузка ----------
  // файл можно бросить в любое место страницы
  const drop = $('#drop');
  let depth = 0;
  window.addEventListener('dragenter', (ev) => { ev.preventDefault(); depth++; drop.classList.add('over'); });
  window.addEventListener('dragover', (ev) => { ev.preventDefault(); });
  window.addEventListener('dragleave', (ev) => { ev.preventDefault(); if (--depth <= 0) { depth = 0; drop.classList.remove('over'); } });
  window.addEventListener('drop', (ev) => {
    ev.preventDefault(); depth = 0; drop.classList.remove('over');
    send([...(ev.dataTransfer?.files || [])]);
  });
  $('#file').addEventListener('change', (ev) => { send([...ev.target.files]); ev.target.value = ''; });

  function foot(text) { $('#lcd-foot').textContent = text; }
  function flash() {
    const f = $('#lcd-foot');
    f.classList.remove('alert'); void f.offsetWidth; f.classList.add('alert');
  }

  async function send(files) {
    files = files.filter((f) => EXT.some((e) => f.name.toLowerCase().endsWith(e)));
    if (!files.length) { foot('НЕ ТОТ ФОРМАТ: НУЖНЫ EXCEL, CSV, XML, ODS'); flash(); return; }
    if (state.busy) return;
    state.busy = true;
    drop.classList.add('busy');
    const bigWas = $('#drop-big').textContent;
    $('#drop-big').textContent = 'ОБРАБОТКА…';
    foot(`ФАЙЛОВ: ${files.length}`);
    const fd = new FormData();
    files.forEach((f) => fd.append('files', f));
    if (state.mode === 'epg') fd.append('channel', state.channel);
    try {
      const res = await api(state.mode === 'dates' ? '/api/fix-dates' : '/api/process', { method: 'POST', body: fd });
      const ok = res.filter((r) => r.status === 'done').length;
      foot(ok === res.length ? `ГОТОВО ${ok}/${res.length}` : `ГОТОВО ${ok}/${res.length} · ЕСТЬ ОШИБКИ`);
      if (ok === 1 && res.length === 1) download(res[0].id);
      await loadJobs(res.map((r) => r.id));
    } catch (e) {
      if (e.message !== '401') { foot('ОШИБКА: ' + e.message.toUpperCase()); flash(); }
    } finally {
      state.busy = false;
      drop.classList.remove('busy');
      $('#drop-big').textContent = bigWas;
    }
  }

  function download(id) {
    const a = document.createElement('a');
    a.href = `/api/jobs/${encodeURIComponent(id)}/download`;
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  // ---------- история ----------
  async function loadJobs(openIds = []) {
    state.jobs = await api('/api/jobs?limit=50');
    const today = new Date().toDateString();
    const n = state.jobs.filter((j) => j.status === 'done' && new Date(j.created * 1000).toDateString() === today).length;
    $('#c-today').textContent = `${n} ФАЙЛ${n % 10 === 1 && n % 100 !== 11 ? '' : (n % 10 >= 2 && n % 10 <= 4 && (n % 100 < 10 || n % 100 >= 20) ? 'А' : 'ОВ')}`;
    const rows = $('#rows');
    if (!state.jobs.length) { rows.innerHTML = '<div class="empty">Пока пусто. Закинь первый EPG.</div>'; return; }
    rows.innerHTML = state.jobs.map((j) => {
      const ok = j.status === 'done';
      const pill = !ok ? '<span class="pill err">ОШИБКА</span>'
        : j.kind === 'dates' ? '<span class="pill dates">ДАТЫ</span>' : '<span class="pill ok">ГОТОВО</span>';
      const time = new Date(j.created * 1000).toLocaleString('ru-RU', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' });
      const action = ok ? `<a class="act" href="/api/jobs/${encodeURIComponent(j.id)}/download">Скачать xlsx</a>` : '';
      const logText = ok ? j.log.join('\n') : j.error;
      const open = openIds.includes(j.id) ? ' open' : '';
      return `<div class="tr">
          <span class="f">${esc(j.src)}<br><span class="label" style="letter-spacing:.04em">${esc(time)}${j.template ? ' · ' + esc(j.template) : ''}</span></span>
          <span class="f f2">${esc(j.out || '—')}</span>
          ${pill}
          <span style="display:flex;flex-direction:column;gap:6px;align-items:flex-start">${action}
            <details${open}><summary class="act" style="list-style:none">${ok ? 'Лог' : 'Почему'}</summary></details></span>
          <pre class="log" hidden>${esc(logText)}</pre>
        </div>`;
    }).join('');
    rows.querySelectorAll('details').forEach((d) => {
      const pre = d.closest('.tr').querySelector('.log');
      const sync = () => { pre.hidden = !d.open; };
      d.addEventListener('toggle', sync);
      sync();
    });
  }

  loadChannels().then(() => loadJobs()).catch(() => {});
})();
