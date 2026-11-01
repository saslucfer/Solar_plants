(() => {
  'use strict';
  const actions = document.querySelector('.data-actions');
  if (!actions) return;
  const button = document.createElement('button');
  button.id = 'export-pdf'; button.type = 'button'; button.className = 'outline-button';
  button.textContent = 'گزارش کامل PDF';
  button.title = 'خلاصه، تمام مدارک و توضیحات ذخیره‌شدهٔ همهٔ نیروگاه‌ها، مستقل از فیلترهای صفحه';
  const feedback = document.createElement('span');
  feedback.className = 'subtle'; feedback.setAttribute('role', 'status'); feedback.setAttribute('aria-live', 'polite');
  actions.prepend(button); actions.append(feedback);
  const local = location.protocol === 'http:' && ['localhost', '127.0.0.1'].includes(location.hostname);
  button.disabled = !local;
  if (!local) {
    button.title = 'برای گزارش کامل PDF، فایل start-dashboard.cmd را اجرا کنید.';
    feedback.textContent = 'برای گزارش کامل PDF، داشبورد را با راه‌انداز اجرا کنید.';
  }
  button.addEventListener('click', async () => {
    if (document.querySelector('#document-rows .date-input:disabled, #document-rows .status-select:disabled') || document.getElementById('note-save')?.textContent === 'در حال ذخیره…') {
      feedback.textContent = 'ابتدا بارگذاری یا ذخیرهٔ تغییرات را کامل کنید و دوباره گزارش بگیرید.';
      return;
    }
    button.disabled = true; button.setAttribute('aria-busy', 'true'); button.textContent = 'در حال تهیه PDF…';
    feedback.textContent = 'در حال تهیهٔ گزارش کامل همهٔ نیروگاه‌ها…';
    try {
      const response = await fetch('/api/report.pdf', { cache: 'no-store' });
      if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error || 'تهیهٔ گزارش انجام نشد.');
      }
      const blob = await response.blob();
      if (blob.type !== 'application/pdf') throw new Error('پاسخ دریافتی فایل PDF نیست. راه‌انداز را دوباره اجرا کنید.');
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      const date = Object.fromEntries(new Intl.DateTimeFormat('en-US-u-ca-persian', { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'Asia/Tehran' }).formatToParts(new Date()).map(part => [part.type, part.value]));
      link.href = url; link.download = `solar-plants-report-${date.year}-${date.month}-${date.day}.pdf`;
      document.body.append(link); link.click(); link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      feedback.textContent = 'گزارش کامل PDF دریافت شد.';
    } catch (error) { feedback.textContent = error.message || 'ارتباط برای دریافت گزارش برقرار نشد.'; }
    finally { button.disabled = false; button.removeAttribute('aria-busy'); button.textContent = 'گزارش کامل PDF'; }
  });
})();
