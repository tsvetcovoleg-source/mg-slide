## mg-slide

Скрипт `generate_presentation.py` берёт первый блок вопроса из `Word.docx` и подставляет данные в 6-й слайд шаблона `Presentation1.pptx`:
- `тематика` → значение после `Тематика:`
- `вопрос` → значение после `Вопрос:`

### Установка

```bash
python3 -m pip install python-docx python-pptx
```

### Запуск

```bash
python3 generate_presentation.py \
  --word Word.docx \
  --template Presentation1.pptx \
  --output Presentation1_filled.pptx
```

Если аргументы не передавать, используются значения по умолчанию:
- `Word.docx`
- `Presentation1.pptx`
- `Presentation1_filled.pptx`
