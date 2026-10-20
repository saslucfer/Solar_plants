(async () => {
  'use strict';

  let data = window.MDL_DATA;
  if (!data || !Array.isArray(data.documents) || !Array.isArray(data.plants)) {
    document.body.innerHTML = '<main style="padding:3rem;font-family:Tahoma;direction:rtl">فایل داده‌ها بارگذاری نشد. لطفاً فایل refresh-data.ps1 را اجرا کنید.</main>';
    return;
  }

  const STATUS = {
    approved: { label: 'تأیید شده', color: '#2eaa88' },
    pending: { label: 'در انتظار تأیید', color: '#e3aa4b' },
    'not-sent': { label: 'ارسال نشده', color: '#a8b8bb' },
    revision: { label: 'در انتظار اصلاح', color: '#d98575' },
    unknown: { label: 'نامشخص', color: '#a79ab8' }
  };
  const statusOrder = Object.keys(STATUS);
  let plantMap = new Map(data.plants.map(plant => [plant.id, plant]));
  let documentMap = new Map(data.documents.map(doc => [doc.id, doc]));
  const pageSize = 25;
  let currentPlant = 'all';
  let page = 1;
  let live = false;
  let currentVersion = '';
  let refreshing = false;
  const writing = new Set();
  let toastTimer;

  const $ = id => document.getElementById(id);
  const faNumber = value => new Intl.NumberFormat('fa-IR').format(value);
  const latinDigits = value => String(value).replace(/[۰-۹]/g, d => String(d.charCodeAt(0) - 1776)).replace(/[٠-٩]/g, d => String(d.charCodeAt(0) - 1632));
  const persianDigits = value => String(value).replace(/[0-9]/g, d => '۰۱۲۳۴۵۶۷۸۹'[Number(d)]);
  const collator = new Intl.Collator('fa');
  const persianCalendar = new Intl.DateTimeFormat('en-US-u-ca-persian', { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'Asia/Tehran' });
  const tehranGregorian = new Intl.DateTimeFormat('en-US-u-ca-gregory', { year: 'numeric', month: 'numeric', day: 'numeric', timeZone: 'Asia/Tehran' });
  const parsedDateCache = new Map();
  let currentOverdueCutoff = '';

  function effectiveStatus(doc) { return doc.status; }
  function effectiveDate(doc) { return latinDigits(doc.lastChange || ''); }
  function showToast(message) {
    const toast = $('toast');
    toast.textContent = message;
    toast.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => toast.classList.remove('show'), 3300);
  }
  function updateConnectionLabel() {
    document.documentElement.dataset.connection = live ? 'live' : 'offline';
    $('connection-status').textContent = live ? 'متصل به فایل‌های اکسل' : 'نمای آفلاین؛ بدون امکان ویرایش';
    if (live) {
      $('sync-note').textContent = 'تاریخ شمسی را وارد کنید و «ثبت» را بزنید؛ کلید Enter یا خروج از کادر هم تاریخ را در فایل اکسل ذخیره می‌کند.';
    } else {
      const launcher = el('bdi', 'launcher-name', 'start-dashboard.cmd'); launcher.dir = 'ltr';
      $('sync-note').replaceChildren('برای ویرایش و همگام‌سازی با اکسل، فایل ', launcher, ' را اجرا کنید.');
    }
    $('refresh-data').disabled = !live;
  }
  function updateSourceDate() {
    $('source-date').textContent = `خواندن از اکسل: ${new Intl.DateTimeFormat('fa-IR-u-ca-persian', { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit', timeZone: 'Asia/Tehran' }).format(new Date(data.generatedAt))}`;
  }
  function updateDashboardDate() {
    $('dashboard-date').textContent = `تاریخ امروز: ${persianDigits(todayPersian())}`;
  }
  async function readResponse(response) {
    const body = await response.json();
    if (!response.ok) throw new Error(body.error || 'ارتباط با فایل اکسل برقرار نشد.');
    return body;
  }
  async function refreshFromExcel(showMessage = false) {
    if (refreshing || writing.size) return;
    refreshing = true;
    try {
      const result = await readResponse(await fetch('/api/data', { cache: 'no-store' }));
      data = result.data;
      currentVersion = result.version;
      plantMap = new Map(data.plants.map(plant => [plant.id, plant]));
      documentMap = new Map(data.documents.map(doc => [doc.id, doc]));
      live = true;
      if (currentPlant !== 'all' && !plantMap.has(currentPlant)) currentPlant = 'all';
      updateConnectionLabel(); updateSourceDate(); renderAll();
      if (showMessage) showToast('آخرین تغییرات فایل‌های اکسل بارگذاری شد.');
    } catch (error) {
      if (showMessage) showToast(error.message);
      if (!live) { updateConnectionLabel(); renderAll(); }
    } finally { refreshing = false; }
  }
  async function checkExcelVersion() {
    if (!live || refreshing || writing.size || document.activeElement?.classList.contains('date-input')) return;
    try {
      const result = await readResponse(await fetch('/api/version', { cache: 'no-store' }));
      if (document.documentElement.dataset.connection === 'disconnected') updateConnectionLabel();
      if (result.version !== currentVersion) await refreshFromExcel();
    } catch {
      document.documentElement.dataset.connection = 'disconnected';
      $('connection-status').textContent = 'ارتباط با اکسل قطع شده است';
    }
  }
  async function saveDocument(doc, status, lastChange) {
    if (!live) { showToast('برای ویرایش، start-dashboard.cmd را اجرا کنید.'); return false; }
    if (writing.has(doc.id)) return false;
    writing.add(doc.id);
    try {
      const response = await fetch('/api/document', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: doc.id, status, lastChange, expectedStatus: doc.status, expectedLastChange: doc.lastChange || '' })
      });
      const result = await readResponse(response);
      Object.assign(doc, result.document);
      currentVersion = result.version;
      data.generatedAt = new Date().toISOString();
      updateSourceDate();
      return true;
    } catch (error) {
      showToast(error.message);
      if (/تغییر کرده|جابه‌جا شده/.test(error.message)) setTimeout(() => refreshFromExcel(), 0);
      return false;
    } finally { writing.delete(doc.id); }
  }
  function el(tag, className, text) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined) node.textContent = text;
    return node;
  }
  function scopeDocs() { return currentPlant === 'all' ? data.documents : data.documents.filter(doc => doc.plantId === currentPlant); }
  function countStatuses(docs) {
    const counts = { approved: 0, pending: 0, 'not-sent': 0, revision: 0, unknown: 0 };
    docs.forEach(doc => counts[effectiveStatus(doc)]++);
    return counts;
  }
  function percent(part, total) { return total ? Math.round(part / total * 100) : 0; }

  function renderMetrics() {
    const docs = scopeDocs();
    const counts = countStatuses(docs);
    const name = currentPlant === 'all' ? 'همه نیروگاه‌ها' : plantMap.get(currentPlant).name;
    $('scope-name').textContent = name;
    $('documents-scope').textContent = name;
    const cards = [
      { label: 'کل مدارک', value: docs.length, color: '#1b6970', suffix: 'مدرک' },
      ...statusOrder.slice(0, 4).map(status => ({ label: STATUS[status].label, value: counts[status], color: STATUS[status].color, suffix: `${percent(counts[status], docs.length)}٪` }))
    ];
    $('metric-cards').replaceChildren(...cards.map(card => {
      const root = el('div', 'metric-card');
      root.style.setProperty('--accent', card.color);
      root.append(el('div', 'metric-label', card.label));
      const value = el('div', 'metric-value', faNumber(card.value));
      value.append(el('span', '', card.suffix));
      root.append(value);
      return root;
    }));
    const bar = $('distribution');
    bar.replaceChildren(...statusOrder.map(status => {
      const segment = el('span');
      segment.style.width = `${docs.length ? counts[status] / docs.length * 100 : 0}%`;
      segment.style.backgroundColor = STATUS[status].color;
      return segment;
    }));
    bar.setAttribute('aria-label', statusOrder.map(s => `${STATUS[s].label}: ${counts[s]}`).join('، '));
    $('distribution-legend').replaceChildren(...statusOrder.map(status => {
      const item = el('span', 'legend-item');
      const dot = el('i'); dot.style.backgroundColor = STATUS[status].color;
      item.append(dot, el('span', '', `${STATUS[status].label} · ${faNumber(counts[status])}`));
      return item;
    }));
    $('approval-rate').textContent = `${faNumber(percent(counts.approved, docs.length))}٪ تأیید شده`;
  }

  function miniBar(docs) {
    const counts = countStatuses(docs);
    const bar = el('div', 'mini-bar');
    statusOrder.forEach(status => {
      const segment = el('span');
      segment.style.width = `${docs.length ? counts[status] / docs.length * 100 : 0}%`;
      segment.style.backgroundColor = STATUS[status].color;
      bar.append(segment);
    });
    return bar;
  }
  function plantCard(id, name, docs) {
    const counts = countStatuses(docs);
    const button = el('button', `plant-card${currentPlant === id ? ' active' : ''}`);
    button.type = 'button';
    button.dataset.plant = id;
    button.setAttribute('aria-pressed', String(currentPlant === id));
    const head = el('div', 'plant-head');
    head.append(el('span', 'plant-name', name), el('span', 'plant-count', `${faNumber(docs.length)} مدرک`));
    const foot = el('div', 'plant-foot');
    foot.append(el('span', '', `${faNumber(counts.approved)} تأیید شده`), el('strong', '', `${faNumber(percent(counts.approved, docs.length))}٪ پیشرفت`));
    button.append(head, foot, miniBar(docs));
    return button;
  }
  function renderPlants() {
    const cards = [plantCard('all', 'همه نیروگاه‌ها', data.documents)];
    data.plants.forEach(plant => cards.push(plantCard(plant.id, plant.name, data.documents.filter(doc => doc.plantId === plant.id))));
    $('plant-grid').replaceChildren(...cards);
  }
  function overdueCutoff() {
    const today = Object.fromEntries(tehranGregorian.formatToParts(new Date()).filter(part => ['year', 'month', 'day'].includes(part.type)).map(part => [part.type, Number(part.value)]));
    const sevenDaysAgo = new Date(Date.UTC(today.year, today.month - 1, today.day - 7));
    const parts = Object.fromEntries(persianCalendar.formatToParts(sevenDaysAgo).filter(part => ['year', 'month', 'day'].includes(part.type)).map(part => [part.type, Number(part.value)]));
    return `${parts.year}/${String(parts.month).padStart(2, '0')}/${String(parts.day).padStart(2, '0')}`;
  }
  function normalizedDocumentDate(doc) {
    const rawDate = effectiveDate(doc);
    if (!rawDate) return '';
    if (!parsedDateCache.has(rawDate)) parsedDateCache.set(rawDate, parsePersianDate(rawDate)?.date || '');
    return parsedDateCache.get(rawDate);
  }
  function renderOverdue() {
    const cutoff = overdueCutoff();
    currentOverdueCutoff = cutoff;
    let total = 0;
    const cards = data.plants.map(plant => {
      const unapproved = data.documents.filter(doc => doc.plantId === plant.id && effectiveStatus(doc) !== 'approved');
      const overdue = unapproved
        .map(doc => ({ doc, date: normalizedDocumentDate(doc) }))
        .filter(item => item.date && item.date < cutoff)
        .sort((a, b) => a.date.localeCompare(b.date) || collator.compare(a.doc.title, b.doc.title));
      total += overdue.length;
      const card = el('article', `overdue-card${overdue.length ? ' has-overdue' : ''}`);
      const heading = el('div', 'overdue-card-head');
      heading.append(el('h3', '', plant.name), el('span', 'overdue-count', `${faNumber(overdue.length)} مدرک`));
      card.append(heading);
      if (!overdue.length) {
        const undated = unapproved.filter(doc => !normalizedDocumentDate(doc)).length;
        if (undated) card.append(el('p', 'overdue-empty', `تاریخ ${faNumber(undated)} مدرک تأییدنشده ثبت نشده است؛ تأخیر آن‌ها قابل محاسبه نیست.`));
      }
      const groups = [
        { status: 'pending', label: 'در انتظار تأیید' },
        { status: 'revision', label: 'در انتظار اصلاح' },
        { status: 'other', label: 'سایر وضعیت‌های تأییدنشده' }
      ];
      groups.forEach(group => {
        const documents = overdue.filter(({ doc }) => group.status === 'other'
          ? !['pending', 'revision'].includes(effectiveStatus(doc))
          : effectiveStatus(doc) === group.status);
        if (group.status === 'other' && !documents.length) return;
        const section = el('section', `overdue-group overdue-group-${group.status}`);
        const groupHead = el('div', 'overdue-group-head');
        groupHead.append(el('h4', '', group.label), el('span', 'overdue-group-count', `${faNumber(documents.length)} مدرک`));
        section.append(groupHead);
        if (!documents.length) {
          section.append(el('p', 'overdue-group-empty', 'مدرکی ثبت نشده است.'));
        } else {
          const list = el('ul', 'overdue-list');
          documents.forEach(({ doc, date }) => {
            const item = el('li', 'overdue-item');
            item.append(el('strong', 'overdue-document', doc.title));
            const details = el('div', 'overdue-details');
            if (doc.code) details.append(el('span', 'overdue-code', doc.code));
            if (group.status === 'other') details.append(el('span', '', STATUS[effectiveStatus(doc)].label));
            details.append(el('span', 'overdue-date', `آخرین تغییر: ${persianDigits(date)}`));
            item.append(details);
            list.append(item);
          });
          section.append(list);
        }
        card.append(section);
      });
      return card;
    });
    $('overdue-total').textContent = `${faNumber(total)} مدرک در ${faNumber(data.plants.length)} نیروگاه`;
    $('overdue-grid').replaceChildren(...cards);
  }
  function renderCategories() {
    const select = $('category-filter');
    const oldValue = select.value;
    const categories = [...new Set(scopeDocs().map(doc => doc.category))].sort(collator.compare);
    select.replaceChildren(new Option('همه بخش‌ها', 'all'), ...categories.map(category => new Option(category, category)));
    select.value = categories.includes(oldValue) ? oldValue : 'all';
  }
  function filteredDocs() {
    const query = latinDigits($('search').value.trim()).toLocaleLowerCase();
    const status = $('status-filter').value;
    const category = $('category-filter').value;
    const missingDate = $('missing-date').checked;
    return scopeDocs().filter(doc => {
      if (status !== 'all' && effectiveStatus(doc) !== status) return false;
      if (category !== 'all' && doc.category !== category) return false;
      if (missingDate && effectiveDate(doc)) return false;
      if (query && !`${doc.title} ${doc.code} ${plantMap.get(doc.plantId)?.name || ''}`.toLocaleLowerCase().includes(query)) return false;
      return true;
    });
  }

  function parsePersianDate(value) {
    const normalized = latinDigits(value.trim()).replace(/[-.]/g, '/');
    if (!normalized) return { date: '' };
    const match = normalized.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
    if (!match) return null;
    const year = Number(match[1]), month = Number(match[2]), day = Number(match[3]);
    if (year < 1200 || year > 1600 || month < 1 || month > 12 || day < 1 || day > 31) return null;
    const date = `${year}/${String(month).padStart(2, '0')}/${String(day).padStart(2, '0')}`;
    const start = Date.UTC(year + 621, 2, 18);
    for (let offset = 0; offset < 375; offset++) {
      const candidate = new Date(start + offset * 86400000);
      const parts = Object.fromEntries(persianCalendar.formatToParts(candidate).filter(p => ['year', 'month', 'day'].includes(p.type)).map(p => [p.type, Number(p.value)]));
      if (parts.year === year && parts.month === month && parts.day === day) return { date };
    }
    return null;
  }
  function todayPersian() {
    const parts = Object.fromEntries(persianCalendar.formatToParts(new Date()).filter(p => ['year', 'month', 'day'].includes(p.type)).map(p => [p.type, p.value]));
    return `${parts.year}/${parts.month}/${parts.day}`;
  }

  function documentRow(doc, ordinal) {
    const status = effectiveStatus(doc);
    const row = el('tr'); row.dataset.id = doc.id;
    row.append(el('td', 'row-num', faNumber(ordinal)));
    const nameCell = el('td', 'doc-name', doc.title);
    nameCell.append(el('span', 'doc-plant', plantMap.get(doc.plantId)?.name || ''));
    row.append(nameCell);
    row.append(el('td', `doc-code${doc.code ? '' : ' empty'}`, doc.code || '—'));
    const category = el('td', 'doc-category', doc.category); category.title = doc.category; row.append(category);
    const statusCell = el('td');
    const statusWrap = el('span', `status-control status-${status}`);
    const statusSelect = el('select', 'status-select');
    statusSelect.setAttribute('aria-label', `وضعیت ${doc.title}`);
    statusOrder.forEach(key => statusSelect.append(new Option(STATUS[key].label, key)));
    statusSelect.value = status;
    statusSelect.disabled = !live;
    statusWrap.append(statusSelect); statusCell.append(statusWrap); row.append(statusCell);
    const dateCell = el('td', 'date-cell');
    const dateWrap = el('div', 'date-entry');
    const dateInput = el('input', 'date-input');
    dateInput.type = 'text'; dateInput.inputMode = 'numeric'; dateInput.placeholder = persianDigits(todayPersian());
    dateInput.maxLength = 10; dateInput.value = persianDigits(effectiveDate(doc));
    dateInput.disabled = !live;
    dateInput.setAttribute('aria-label', `تاریخ آخرین تغییر وضعیت ${doc.title}`);
    const saveDate = el('button', 'save-date-button', 'ثبت'); saveDate.type = 'button';
    saveDate.title = 'ذخیره تاریخ این مدرک در فایل اکسل'; saveDate.disabled = !live;
    const today = el('button', 'today-button', 'امروز'); today.type = 'button'; today.title = 'ثبت تاریخ امروز به تقویم شمسی'; today.disabled = !live;
    dateWrap.append(dateInput, saveDate, today); dateCell.append(dateWrap); row.append(dateCell);
    return row;
  }
  function renderDocuments() {
    const docs = filteredDocs();
    const pages = Math.max(1, Math.ceil(docs.length / pageSize));
    page = Math.min(page, pages);
    const start = (page - 1) * pageSize;
    $('document-rows').replaceChildren(...docs.slice(start, start + pageSize).map((doc, index) => documentRow(doc, start + index + 1)));
    $('empty-state').hidden = docs.length !== 0;
    $('result-count').textContent = `${faNumber(docs.length)} مدرک`;
    $('page-info').textContent = docs.length ? `${faNumber(start + 1)} تا ${faNumber(Math.min(start + pageSize, docs.length))} از ${faNumber(docs.length)}` : '۰ مدرک';
    $('prev-page').disabled = page <= 1;
    $('next-page').disabled = page >= pages;
  }
  function renderAll() { renderMetrics(); renderPlants(); renderOverdue(); renderCategories(); renderDocuments(); }

  function csvCell(value) {
    let text = String(value ?? '');
    if (/^[\s]*[=+@-]/.test(text)) text = `'${text}`;
    return `"${text.replace(/"/g, '""')}"`;
  }
  function download(filename, content, mime) {
    const url = URL.createObjectURL(new Blob([content], { type: mime }));
    const link = document.createElement('a'); link.href = url; link.download = filename;
    document.body.append(link); link.click(); link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
  function dateStamp() { return new Date().toISOString().slice(0, 10); }
  function exportCsv() {
    const columns = ['نیروگاه', 'بخش', 'نام مدرک', 'شماره مدرک', 'وضعیت', 'آخرین تغییر وضعیت (شمسی)', 'فایل منبع', 'ردیف منبع'];
    const rows = data.documents.map(doc => {
      const plant = plantMap.get(doc.plantId);
      return [plant.name, doc.category, doc.title, doc.code, STATUS[effectiveStatus(doc)].label, effectiveDate(doc), plant.file, doc.sourceRow];
    });
    download(`solar-mdl-${dateStamp()}.csv`, '\ufeff' + [columns, ...rows].map(row => row.map(csvCell).join(',')).join('\r\n'), 'text/csv;charset=utf-8');
  }
  updateConnectionLabel(); updateDashboardDate(); updateSourceDate(); renderAll();
  if (location.protocol === 'http:' && ['localhost', '127.0.0.1'].includes(location.hostname)) await refreshFromExcel();

  $('plant-grid').addEventListener('click', event => {
    const button = event.target.closest('[data-plant]');
    if (!button) return;
    currentPlant = button.dataset.plant; page = 1; renderAll();
  });
  $('search').addEventListener('input', () => { page = 1; renderDocuments(); });
  ['status-filter', 'category-filter', 'missing-date'].forEach(id => $(id).addEventListener('change', () => { page = 1; renderDocuments(); }));
  $('clear-filters').addEventListener('click', () => {
    currentPlant = 'all'; $('search').value = ''; $('status-filter').value = 'all'; $('missing-date').checked = false; page = 1; renderAll();
  });
  $('prev-page').addEventListener('click', () => { if (page > 1) { page--; renderDocuments(); } });
  $('next-page').addEventListener('click', () => { page++; renderDocuments(); });
  $('document-rows').addEventListener('change', async event => {
    if (!event.target.matches('.status-select')) return;
    const id = event.target.closest('tr').dataset.id;
    const doc = documentMap.get(id);
    if (event.target.value === doc.status) return;
    event.target.disabled = true;
    const saved = await saveDocument(doc, event.target.value, todayPersian());
    renderMetrics(); renderPlants(); renderOverdue(); renderDocuments();
    if (saved) {
      showToast(`وضعیت و تاریخ آخرین تغییر (${persianDigits(effectiveDate(doc))}) در اکسل ذخیره شد.`);
    }
  });
  async function saveDateInput(input) {
    if (input.disabled) return;
    const result = parsePersianDate(input.value);
    if (!result) { input.classList.add('invalid'); showToast(`تاریخ شمسی معتبر نیست. نمونه: ${persianDigits(todayPersian())}`); return; }
    input.classList.remove('invalid');
    const doc = documentMap.get(input.closest('tr').dataset.id);
    if (result.date === effectiveDate(doc)) { input.value = persianDigits(result.date); return; }
    input.disabled = true;
    const saved = await saveDocument(doc, doc.status, result.date);
    input.disabled = false;
    if (!saved) { input.value = persianDigits(effectiveDate(doc)); return; }
    input.value = persianDigits(result.date);
    renderOverdue();
    if ($('missing-date').checked) renderDocuments();
    showToast(result.date ? 'تاریخ آخرین تغییر در اکسل ذخیره شد.' : 'تاریخ در اکسل پاک شد.');
  }
  $('document-rows').addEventListener('focusout', event => { if (event.target.matches('.date-input') && !event.relatedTarget?.matches('.today-button, .save-date-button')) saveDateInput(event.target); });
  $('document-rows').addEventListener('keydown', event => { if (event.target.matches('.date-input') && event.key === 'Enter') event.target.blur(); });
  $('document-rows').addEventListener('click', event => {
    const row = event.target.closest('tr'); if (!row) return;
    if (event.target.matches('.today-button')) {
      const input = row.querySelector('.date-input'); input.value = persianDigits(todayPersian()); saveDateInput(input);
    }
    if (event.target.matches('.save-date-button')) saveDateInput(row.querySelector('.date-input'));
  });
  $('export-csv').addEventListener('click', exportCsv);
  $('refresh-data').addEventListener('click', () => refreshFromExcel(true));
  window.addEventListener('focus', checkExcelVersion);
  setInterval(checkExcelVersion, 7000);
  setInterval(() => { updateDashboardDate(); if (overdueCutoff() !== currentOverdueCutoff) renderOverdue(); }, 60000);
})();
