const { chromium } = require('playwright');
const docs = [
 { f: 'ex_nrb.png', head: 'СанПиН 2.6.1.2523-09 «Нормы радиационной безопасности (НРБ-99/2009)»', sub: 'IV. Требования к защите от природного облучения в производственных условиях',
   body: ['<b>4.1.</b> Эффективная доза облучения природными источниками излучения всех работников, включая персонал, не должна превышать 5 мЗв в год в производственных условиях (любые профессии и производства).', '<span class="dots">…</span>', '<b>4.3.</b> <mark>Воздействие космических излучений на экипажи самолетов нормируется как природное облучение в производственных условиях по п. 4.1.</mark>'] },
 { f: 'ex_3fz.png', head: 'Федеральный закон от 09.01.1996 № 3-ФЗ «О радиационной безопасности населения»', sub: 'Статья 18. Контроль и учет индивидуальных доз облучения',
   body: ['Контроль и учет индивидуальных доз облучения, полученных гражданами при использовании источников ионизирующего излучения, проведении медицинских рентгенорадиологических процедур, а также <mark>обусловленных естественным радиационным и техногенно измененным радиационным фоном</mark>, осуществляются в рамках единой государственной системы контроля и учета индивидуальных доз облучения…'] },
 { f: 'ex_vk.png', head: 'Воздушный кодекс Российской Федерации от 19.03.1997 № 60-ФЗ', sub: 'Статья 56. Экипаж воздушного судна',
   body: ['1. <mark>Экипаж пилотируемого воздушного судна состоит из летного экипажа</mark> (командира, других лиц летного состава) <mark>и кабинного экипажа</mark> (бортоператоров и бортпроводников).'] },
 { f: 'ex_euratom.png', head: 'Council Directive 2013/59/Euratom of 5 December 2013', sub: 'Article 35. Arrangements in workplaces — paragraph 3', lang: 'en',
   body: ['In the case of undertakings operating aircraft where the effective dose to the crew from cosmic radiation is liable to exceed 6 mSv per year, the relevant requirements set out in this Chapter shall apply, allowing for the specific features of this exposure situation. <mark>Member States shall ensure that where the effective dose to the crew is liable to be above 1 mSv per year, the competent authority requires the undertaking to take appropriate measures, in particular:</mark>',
          '(a) to assess the exposure of the crew concerned;', '(b) to take into account the assessed exposure when organising working schedules with a view to reducing the doses of highly exposed aircrew;', '(c) to inform the workers concerned of the health risks their work involves and their individual dose;', '(d) to apply Article 10(1) to pregnant aircrew.'] },
];
(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' in {} ? undefined : undefined });
  const p = await b.newPage({ deviceScaleFactor: 2 });
  for (const d of docs) {
    const html = `<html><head><meta charset="utf-8"><style>
      body{margin:0;background:#fff;font-family:'Times New Roman','Liberation Serif',serif;}
      .page{width:820px;padding:34px 44px 30px;border:1px solid #bfbfbf;box-shadow:3px 3px 0 #d9d9d9;background:#fffefa}
      .head{font-weight:bold;font-size:20px;text-align:center;margin-bottom:6px}
      .sub{font-size:18px;text-align:center;font-style:italic;margin-bottom:16px;color:#333}
      p{font-size:19px;line-height:1.45;text-align:justify;margin:0 0 8px;text-indent:${d.lang==='en'?0:28}px}
      mark{background:#fff2a8} .dots{display:block;text-align:center;text-indent:0}
    </style></head><body><div class="page" id="pg"><div class="head">${d.head}</div><div class="sub">${d.sub}</div>${d.body.map(x=>'<p>'+x+'</p>').join('')}</div></body></html>`;
    await p.setContent(html);
    await (await p.$('#pg')).screenshot({ path: d.f });
    console.log(d.f);
  }
  await b.close();
})();
