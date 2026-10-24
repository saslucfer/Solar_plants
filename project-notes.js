(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const kinds = { note: 'یادداشت', comment: 'کامنت', report: 'گزارش' };
  const dateFormat = new Intl.DateTimeFormat('fa-IR-u-ca-persian', { dateStyle: 'medium', timeStyle: 'short', timeZone: 'Asia/Tehran' });
  const number = value => new Intl.NumberFormat('fa-IR').format(value);
  const canConnect = location.protocol === 'http:' && ['localhost', '127.0.0.1'].includes(location.hostname);
  const drafts = new Map();
  let plants = [], notes = [], project = '', editing = null;
  let ready = false, saving = false, refreshing = false, generation = 0, fingerprint = '';
  function node(tag, className, text) {
    const element = document.createElement(tag); element.className = className;
    if (text !== undefined) element.textContent = text;
    return element;
  }
  function feedback(text, error = false) {
    $('note-feedback').textContent = text;
    $('note-feedback').classList.toggle('is-error', error);
  }
  function formValue() { return { kind: $('note-kind').value, title: $('note-title').value, body: $('note-body').value }; }
  function fillForm(value = {}) {
    $('note-kind').value = value.kind || 'note'; $('note-title').value = value.title || ''; $('note-body').value = value.body || '';
  }
  function controls() {
    $('note-project').disabled = saving || Boolean(editing) || !plants.length;
    ['note-kind', 'note-title', 'note-body'].forEach(id => { $(id).disabled = saving || !canConnect; });
    $('note-save').disabled = !ready || saving || !project;
    $('note-save').textContent = saving ? 'در حال ذخیره…' : editing ? 'ذخیره ویرایش' : 'ثبت توضیحات';
    $('note-cancel').hidden = !editing; $('note-cancel').disabled = saving;
    $('notes-refresh').disabled = !canConnect || refreshing || saving;
  }
  function render() {
    $('note-project').replaceChildren(...plants.map(plant => new Option(`${plant.name} (${number(notes.filter(note => note.plantId === plant.id).length)} مورد)`, plant.id)));
    $('note-project').value = project;
    $('notes-project-name').textContent = `توضیحات ${plants.find(plant => plant.id === project)?.name || 'پروژه'}`;
    const records = notes.filter(note => note.plantId === project).sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
    $('notes-count').textContent = `${number(records.length)} مورد`;
    if (!ready) {
      $('notes-list').replaceChildren(node('p', 'notes-empty', canConnect ? 'برای مشاهده یادداشت‌ها، ارتباط با راه‌انداز داشبورد باید برقرار باشد.' : 'برای مشاهده و ثبت یادداشت‌های ذخیره‌شده، فایل start-dashboard.cmd را اجرا کنید.'));
    } else if (!records.length) {
      $('notes-list').replaceChildren(node('p', 'notes-empty', 'برای این نیروگاه هنوز توضیحی ثبت نشده است. اولین یادداشت را از فرم کنار این بخش اضافه کنید.'));
    } else {
      $('notes-list').replaceChildren(...records.map(note => {
        const card = node('article', 'project-note');
        const header = node('div', 'project-note-head');
        header.append(node('span', `note-kind note-kind-${note.kind}`, kinds[note.kind] || 'یادداشت'));
        const time = node('time', 'note-date', `${note.revision > 1 ? 'ویرایش: ' : 'ثبت: '}${dateFormat.format(new Date(note.updatedAt))}`);
        time.dateTime = note.updatedAt; header.append(time);
        if (note.title) { const title = node('h4', '', note.title); title.dir = 'auto'; card.append(title); }
        const body = node('p', 'project-note-body', note.body); body.dir = 'auto';
        const edit = node('button', 'text-button note-edit', 'ویرایش'); edit.type = 'button'; edit.dataset.noteId = note.id; edit.disabled = saving;
        card.prepend(header); card.append(body, edit);
        return card;
      }));
    }
    controls();
  }
  function changeProject(id) {
    if (id === project || !plants.some(plant => plant.id === id) || editing || saving) return;
    if (project) drafts.set(project, formValue());
    project = id; fillForm(drafts.get(project));
  }
  window.PROJECT_NOTES = { setContext(nextPlants, selectedPlant) {
    plants = nextPlants;
    if (!project) project = plants[0]?.id || '';
    if (selectedPlant !== 'all') changeProject(selectedPlant);
    render();
  } };
  async function responseJson(response) {
    const body = await response.json();
    if (!response.ok) throw new Error(body.error || 'ذخیره یا دریافت یادداشت‌ها انجام نشد.');
    return body;
  }
  async function refresh(showMessage = false) {
    if (!canConnect || refreshing || saving) return;
    refreshing = true; controls(); const startedGeneration = generation;
    try {
      const result = await responseJson(await fetch('/api/notes', { cache: 'no-store' }));
      if (!Array.isArray(result.notes)) throw new Error('پاسخ بخش یادداشت‌ها معتبر نیست.');
      if (startedGeneration !== generation) return;
      const signature = JSON.stringify(result.notes);
      const changed = signature !== fingerprint || !ready;
      notes = result.notes; fingerprint = signature; ready = true;
      if (changed) render();
      if (showMessage) feedback('یادداشت‌ها تازه شدند.');
      else if (changed && !editing) feedback('توضیحات پس از ثبت، در اجراهای بعدی هم در دسترس خواهند بود.');
    } catch (error) {
      if (startedGeneration !== generation) return;
      ready = false; render(); feedback(`ارتباط با یادداشت‌ها برقرار نشد. راه‌انداز را دوباره اجرا کنید. ${error.message}`, true);
    } finally { refreshing = false; controls(); }
  }
  $('note-project').addEventListener('change', event => { changeProject(event.target.value); render(); });
  $('note-form').addEventListener('submit', async event => {
    event.preventDefault(); if (!ready || saving) return;
    const value = formValue();
    if (!value.body.trim()) { feedback('متن توضیحات را وارد کنید.', true); $('note-body').focus(); return; }
    saving = true; generation++; controls();
    const payload = { ...value, plantId: project };
    if (editing) Object.assign(payload, { id: editing.id, expectedRevision: editing.revision });
    try {
      const result = await responseJson(await fetch('/api/notes', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) }));
      notes = notes.filter(note => note.id !== result.note.id).concat(result.note);
      fingerprint = JSON.stringify(notes);
      const wasEditing = Boolean(editing); editing = null;
      if (wasEditing) fillForm(drafts.get(project));
      else { drafts.delete(project); fillForm(); }
      feedback('توضیحات این پروژه ذخیره شد.');
    } catch (error) { feedback(error.message, true); }
    finally { saving = false; render(); }
  });
  $('notes-list').addEventListener('click', event => {
    const button = event.target.closest('[data-note-id]'); if (!button || saving) return;
    const record = notes.find(note => note.id === button.dataset.noteId); if (!record) return;
    if (!editing) drafts.set(project, formValue());
    editing = { ...record }; fillForm(record); render();
    feedback('در حال ویرایش توضیحات ثبت‌شده.'); $('note-body').focus();
  });
  $('note-cancel').addEventListener('click', () => { editing = null; fillForm(drafts.get(project)); render(); feedback('ویرایش لغو شد.'); });
  $('notes-refresh').addEventListener('click', () => refresh(true));
  window.addEventListener('focus', () => refresh());
  setInterval(() => refresh(), 15000);
  controls();
  if (canConnect) refresh();
  else feedback('برای ذخیرهٔ توضیحات، داشبورد را با start-dashboard.cmd اجرا کنید.');
})();
