// Vendoor color grabber — شغّله من Console وانت عامل login على aff.ven-door.com
// بيقرا صفحات المنتجات بس (مش بيعدّل أي حاجة) وفي الآخر بينزّل ملف vendoor_products.json
(async () => {
  const ids = __IDS__;
  const out = {};
  const clean = s => (s || '').replace(/\s+/g, ' ').trim();
  for (let i = 0; i < ids.length; i++) {
    const id = ids[i];
    try {
      const r = await fetch('/product/' + id, { credentials: 'include' });
      if (r.url.includes('/login')) { console.error('انت مش عامل login — اعمل login وشغّله تاني'); return; }
      const doc = new DOMParser().parseFromString(await r.text(), 'text/html');
      doc.querySelectorAll('script,style,noscript,svg,link,meta').forEach(e => e.remove());
      const bits = [...doc.querySelectorAll('option,label,button,select,[class*=color],[class*=Color],[class*=variant],[class*=attribute],[data-color],[title],td,th,li')]
        .map(e => clean([e.getAttribute('data-color'), e.getAttribute('title'), e.getAttribute('value'), e.textContent].filter(Boolean).join(' | ')))
        .filter(t => t && t.length < 200);
      out[id] = {
        title: clean(doc.title),
        h: [...doc.querySelectorAll('h1,h2,h3,h4')].map(e => clean(e.textContent)).slice(0, 20),
        bits: [...new Set(bits)].slice(0, 400),
        text: clean((doc.querySelector('main') || doc.body).textContent).slice(0, 20000)
      };
    } catch (e) { out[id] = { error: String(e) }; }
    console.log(`${i + 1} / ${ids.length}`);
    await new Promise(res => setTimeout(res, 250));
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([JSON.stringify(out)], { type: 'application/json' }));
  a.download = 'vendoor_products.json';
  document.body.appendChild(a); a.click();
  console.log('خلصت ✅ الملف نزل: vendoor_products.json');
})();
