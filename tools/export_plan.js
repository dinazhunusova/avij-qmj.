// Экспорт данных урока из приложения (без изменения приложений) в JSON для make_pptx.py
// Использование: node tools/export_plan.js <app_dir> <grade_key> <quarter> <topic_idx> <out.json> [date]
//   app_dir:   sharapatova | avij | shaikhieva | samalova
//   grade_key: '7', '8', '9', '10Ж', '11Қ', 'T5', 'L5' ...
// Нужен playwright (в среде Claude Code уже установлен).
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

(async () => {
  const [app, gk, q, idx, out, date] = process.argv.slice(2);
  if (!app || !gk || !q || !idx || !out) {
    console.error('usage: node export_plan.js <app_dir> <grade_key> <quarter> <topic_idx> <out.json> [date]');
    process.exit(1);
  }
  const file = 'file://' + path.resolve(__dirname, '..', app, 'index.html');
  const browser = await chromium.launch();
  const page = await (await browser.newContext()).newPage();
  await page.goto(file);
  const plan = await page.evaluate(async ([gk, q, idx, date]) => {
    const png = svg => new Promise((res, rej) => {
      const img = new Image();
      img.onload = () => {
        const cv = document.createElement('canvas');
        cv.width = (img.naturalWidth || 640) * 2; cv.height = (img.naturalHeight || 360) * 2;
        const c = cv.getContext('2d'); c.fillStyle = '#fff'; c.fillRect(0, 0, cv.width, cv.height);
        c.drawImage(img, 0, 0, cv.width, cv.height);
        res(cv.toDataURL('image/png').split(',')[1]);
      };
      img.onerror = () => rej(new Error('svg->png'));
      img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
    });
    const p = LESSON.buildPlan(gk, +q, +idx, { date: date || new Date().toISOString().slice(0, 10) });
    const imgs = {};
    for (const k of ['picSvg', 'kindPicSvg', 'chartSvg', 'mapSvg', 'cardsSvg', 'framesSvg']) {
      if (p[k]) { imgs[k.replace('Svg', '')] = await png(p[k]); }
      delete p[k];
    }
    p.images = imgs;
    p.shape = p.t1 ? 'language' : 'science';
    return p;
  }, [gk, q, idx, date || '']);
  await browser.close();
  fs.writeFileSync(out, JSON.stringify(plan));
  console.log('ok', out, plan.shape, plan.title);
})();
