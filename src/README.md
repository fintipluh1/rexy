# Исходники доклада

- `build3.js`: генератор презентаций (`node build3.js OUT.pptx gost|nice`, затем `python3 postprocess.py OUT.pptx`)
- `content3.json`: тексты слайдов, таблицы и список источников
- `notes3_en.js`: английский текст для двух спикеров (Никита и Илья, уровень B1)
- `script_docx.js` и `script3_extra.json`: печатный вариант выступления (`node script_docx.js OUT.docx`)
- `img/`: рисунки, `excerpts/make.js`: рендер фрагментов норм через Playwright
- `article_unpacked/`: распакованная статья (docx); собирать через `cd article_unpacked && zip -qrX ../article.docx .`

Зависимости: `npm install pptxgenjs@3.12.0 jszip docx` (в этой папке).
