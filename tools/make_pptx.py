#!/usr/bin/env python3
"""Презентация к уроку (.pptx) из JSON, экспортированного export_plan.js.

Использование:
  python3 tools/make_pptx.py plan.json out.pptx [--lang kk|ru] [--teacher "Имя"] [--school "Школа"] [--subject "Физика"]

Ответы к заданиям кладутся в заметки слайдов (их видит только учитель).
"""
import argparse, base64, io, json, re, sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

NAVY = RGBColor(0x2F, 0x4F, 0x86); INK = RGBColor(0x1C, 0x27, 0x40); MUTED = RGBColor(0x5A, 0x68, 0x82)
AMBER = RGBColor(0xE9, 0xA2, 0x3B); BG = RGBColor(0xF4, 0xF7, 0xFB); CARD = RGBColor(0xFF, 0xFF, 0xFF)
SOFT = RGBColor(0xE6, 0xED, 0xF8); LINE = RGBColor(0xD8, 0xE0, 0xEC); GREEN = RGBColor(0x2C, 0x7A, 0x54)
FONT = 'Arial'
W, H = 13.333, 7.5

L = {
 'kk': dict(goal='Сабақтың мақсаты', obj='Оқу мақсаттары', think='«Ойлан!»', t_ata='Ата', t_an='Анықта', t_bol='Болжа',
   q_ata='Суретте не бейнеленген? Не көріп тұрсың?', q_an='Суретке қатысты қандай шамалар / ұғымдар бар?', q_bol='Бұл не себепті болады деп ойлайсың? Болжамыңды негізде.',
   newk='Жаңа білім', law='Негізгі заңдылық / формулалар', terms='Терминдер', task='Тапсырма', t2='«Шама – белгі – бірлік»', sample='Үлгі',
   t3='«Зертте!», «Талда!»', t4='«Үлгі бойынша орында»', steps='Берілгені → ХБЖ → Формула → Шешуі → Жауабы', fl='Функционалдық сауаттылық',
   sit='Жағдаят', quest='Сұрақтар', cls='сыныпта', home='үйге', refl='Рефлексия', r1='Бүгін мен … түсіндім.', r2='Маған … қиын болды.', r3='Бұл білімді … қолданамын.',
   hw='Үй тапсырмасы', hwt='№ 5 тапсырманың қалған сұрақтарын орындау.', thanks='Назарларыңызға рахмет!', ans='Жауабы', grade='сынып', pts='балл',
   vocab='Тақырыптық сөздік', word='Сөз', mean='Мағынасы', text='Мәтінмен жұмыс', gaps='Бос орындарды толтыр', match='Сәйкестендір', frames='Суреттер бойынша жұмыс',
   write='Жазылым (үйге)', rules='Ереже', listen='Мәтінді тыңда және оқы', default_subject='Физика'),
 'ru': dict(goal='Цель урока', obj='Цели обучения', think='«Подумай!»', t_ata='Назови', t_an='Определи', t_bol='Предположи',
   q_ata='Что изображено на рисунке? Что ты видишь?', q_an='Какие понятия связаны с рисунком?', q_bol='Как это работает? Обоснуй своё предположение.',
   newk='Новое знание', law='Основные правила / формулы', terms='Термины', task='Задание', t2='«Понятие – пример – назначение»', sample='Образец',
   t3='«Исследуй!», «Проанализируй!»', t4='«Выполни по образцу»', steps='Условие → Алгоритм → Решение → Проверка → Ответ', fl='Функциональная грамотность',
   sit='Ситуация', quest='Вопросы', cls='в классе', home='дома', refl='Рефлексия', r1='Сегодня я понял(а) …', r2='Мне было трудно …', r3='Эти знания я применю …',
   hw='Домашнее задание', hwt='Выполнить оставшиеся вопросы задания № 5.', thanks='Спасибо за внимание!', ans='Ответ', grade='класс', pts='б.',
   vocab='Словарь', word='Слово', mean='Значение', text='Работа с текстом', gaps='Заполни пропуски', match='Сопоставь', frames='Работа по рисункам',
   write='Письмо (дома)', rules='Правило', listen='Послушай и прочитай текст', default_subject='Информатика и ИИ'),
}



_SUP = dict(zip('0123456789+-=()n', '⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿ'))
_SUB = dict(zip('0123456789+-=()', '₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎'))


def U(text):
    """X_{2} -> X₂, 10^{-7} -> 10⁻⁷ (если символ нельзя — оставляем как есть без скобок)."""
    if not isinstance(text, str):
        return text
    def rep(m):
        kind, body = m.group(1), m.group(2)
        table = _SUP if kind == '^' else _SUB
        if all(c in table for c in body):
            return ''.join(table[c] for c in body)
        return kind + body
    return re.sub(r'([\^_])\{([^}]*)\}', rep, text)


def rgb_fill(shape, color):
    shape.fill.solid(); shape.fill.fore_color.rgb = color


def bg(slide, color=BG):
    f = slide.background.fill; f.solid(); f.fore_color.rgb = color


def fit(text, big, small, limit):
    """Размер шрифта по длине текста."""
    n = len(text)
    if n <= limit * 0.5: return big
    if n >= limit * 1.6: return small
    return int(big - (big - small) * (n - limit * 0.5) / (limit * 1.1))


def tb(slide, x, y, w, h, text, size=24, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False, font=FONT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05); tf.margin_top = tf.margin_bottom = Inches(0.03)
    lines = text if isinstance(text, list) else [text]
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = U(ln)
        r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic; r.font.color.rgb = color; r.font.name = font
        if i: p.space_before = Pt(size * 0.35)
    return box


def card(slide, x, y, w, h, fill=CARD, line=LINE, radius=0.08):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    s.adjustments[0] = radius
    rgb_fill(s, fill); s.line.color.rgb = line; s.line.width = Pt(1); s.shadow.inherit = False
    return s


def header(slide, title, tag=None):
    bg(slide)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(W), Inches(1.05)); rgb_fill(bar, NAVY); bar.line.fill.background()
    tb(slide, 0.55, 0.14, 10.8, 0.8, title, 30, True, RGBColor(255, 255, 255), anchor=MSO_ANCHOR.MIDDLE)
    if tag:
        c = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.2), Inches(0.26), Inches(1.6), Inches(0.52)); c.adjustments[0] = 0.5
        rgb_fill(c, AMBER); c.line.fill.background()
        tf = c.text_frame; tf.margin_left = tf.margin_right = 0
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = tag; r.font.size = Pt(18); r.font.bold = True; r.font.color.rgb = INK; r.font.name = FONT
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE


def img(slide, b64, x, y, w, h):
    from PIL import Image
    data = base64.b64decode(b64)
    iw, ih = Image.open(io.BytesIO(data)).size
    k = min(w / iw, h / ih)
    ww, hh = iw * k, ih * k
    slide.shapes.add_picture(io.BytesIO(data), Inches(x + (w - ww) / 2), Inches(y + (h - hh) / 2), Inches(ww), Inches(hh))


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def table(slide, x, y, w, head, rows, col_w=None, size=18, row_h=0.5):
    nr, nc = len(rows) + 1, len(head)
    gs = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(row_h * nr))
    t = gs.table
    tot = sum(col_w) if col_w else nc
    for i in range(nc):
        t.columns[i].width = Inches(w * ((col_w[i] if col_w else 1) / tot))
    def setc(cell, text, bold=False, fill=None, color=INK):
        cell.text = ''
        p = cell.text_frame.paragraphs[0]; r = p.add_run(); r.text = U(str(text))
        r.font.size = Pt(size); r.font.bold = bold; r.font.name = FONT; r.font.color.rgb = color
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        if fill is not None:
            cell.fill.solid(); cell.fill.fore_color.rgb = fill
    for j, h in enumerate(head): setc(t.cell(0, j), h, True, NAVY, RGBColor(255, 255, 255))
    for i, row in enumerate(rows, 1):
        for j, v in enumerate(row): setc(t.cell(i, j), v, False, CARD if i % 2 else SOFT)
    return gs


def bullets(slide, items, x, y, w, h, size=26):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = box.text_frame; tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(size * 0.5)
        r = p.add_run(); r.text = '•  ' + U(it); r.font.size = Pt(size); r.font.name = FONT; r.font.color.rgb = INK


def clean(s):
    return re.sub(r'\*\*', '', s or '')


def title_slide(prs, plan, lab, args):
    s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s, NAVY)
    tb(s, 0.8, 0.5, 11.7, 0.7, args.school or plan.get('school', ''), 20, False, RGBColor(0xC9, 0xD6, 0xEE))
    subj = args.subject or lab['default_subject']
    chip_text = plan['grade'] if subj.lower() in plan['grade'].lower() else f"{subj} · {plan['grade']}"
    chip = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(min(1.0 + 0.2 * len(chip_text), 11.5)), Inches(0.6)); chip.adjustments[0] = 0.5
    rgb_fill(chip, AMBER); chip.line.fill.background()
    tf = chip.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    r = p.add_run(); r.text = chip_text; r.font.size = Pt(20); r.font.bold = True; r.font.color.rgb = INK; r.font.name = FONT
    title = plan['title']
    tb(s, 0.8, 2.4, 11.7, 2.6, title, fit(title, 54, 34, 50), True, RGBColor(255, 255, 255), anchor=MSO_ANCHOR.MIDDLE)
    codes = ', '.join(o[0] for o in plan.get('objectives', []) if o[0])
    if codes: tb(s, 0.8, 5.2, 11.7, 0.5, codes, 20, False, RGBColor(0xC9, 0xD6, 0xEE))
    tb(s, 0.8, 6.3, 11.7, 0.6, args.teacher or plan.get('teacher', ''), 22, True, RGBColor(255, 255, 255))


def goal_slide(prs, plan, lab):
    s = prs.slides.add_slide(prs.slide_layouts[6]); header(s, lab['goal'])
    card(s, 0.6, 1.4, 12.1, 5.5)
    objs = [f"{o[0]}  {o[1]}".strip() for o in plan.get('objectives', [])]
    tb(s, 0.9, 1.55, 11.5, 0.5, lab['obj'], 22, True, NAVY)
    bullets(s, objs, 0.9, 2.05, 11.5, 4.6, fit(' '.join(objs), 28, 20, 220))


def task_text_slide(prs, lab, title, tag, body, code=None, answer=None, extra=None, steps=None):
    s = prs.slides.add_slide(prs.slide_layouts[6]); header(s, title, tag)
    card(s, 0.6, 1.4, 12.1, 5.5)
    size = fit(body, 30, 20, 260)
    tb(s, 0.9, 1.6, 11.5, 3.2 if code else 4.4, body, size, False, INK)
    y = 5.0 if code else 5.9
    if code:
        tb(s, 0.9, 4.5, 11.5, 2.2, code, 16, False, INK, font='Courier New')
    if steps:
        c = card(s, 0.9, 6.0, 11.5, 0.7, SOFT, SOFT); tb(s, 1.0, 6.05, 11.3, 0.6, steps, 20, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)
    if answer: notes(s, f"{lab['ans']}: {answer}")
    return s


def science(prs, plan, lab, args):
    im = plan.get('images', {})
    title_slide(prs, plan, lab, args)
    goal_slide(prs, plan, lab)
    # №1
    s = prs.slides.add_slide(prs.slide_layouts[6]); header(s, f"{lab['task']} 1. {lab['think']}", '№ 1')
    card(s, 0.6, 1.4, 6.4, 5.5)
    if im.get('kindPic'): img(s, im['kindPic'], 0.75, 1.55, 6.1, 5.2)
    y = 1.5
    for tag, q in ((lab['t_ata'], lab['q_ata']), (lab['t_an'], lab['q_an']), (lab['t_bol'], lab['q_bol'])):
        card(s, 7.3, y, 5.4, 1.65, CARD)
        tb(s, 7.5, y + 0.08, 5.0, 0.45, tag, 20, True, NAVY)
        tb(s, 7.5, y + 0.55, 5.0, 1.05, q, 20, False, INK)
        y += 1.8
    notes(s, f"{lab['ans']}: " + (plan.get('terms') or ''))
    # Новое знание
    s = prs.slides.add_slide(prs.slide_layouts[6]); header(s, lab['newk'] + ': ' + plan['title'][:70])
    card(s, 0.6, 1.4, 12.1, 5.5)
    th = plan.get('theory', [])
    bullets(s, th, 0.9, 1.55, 11.5, 3.6, fit(' '.join(th), 26, 19, 380))
    c = card(s, 0.9, 5.2, 11.5, 1.5, SOFT, SOFT)
    tb(s, 1.05, 5.25, 11.2, 0.45, lab['law'], 18, True, NAVY)
    tb(s, 1.05, 5.7, 11.2, 1.0, plan.get('law', ''), fit(plan.get('law', ''), 22, 16, 180), True, INK)
    s = prs.slides.add_slide(prs.slide_layouts[6]); header(s, lab['terms'])
    terms = [t.strip() for t in (plan.get('terms') or '').split(',') if t.strip()]
    x, y = 0.7, 1.7
    for t in terms:
        w = min(0.9 + 0.2 * len(t), 7.5)
        if x + w > 12.7: x, y = 0.7, y + 1.0
        c = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.75)); c.adjustments[0] = 0.5
        rgb_fill(c, SOFT); c.line.color.rgb = NAVY
        tf = c.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = U(t); r.font.size = Pt(22); r.font.bold = True; r.font.color.rgb = NAVY; r.font.name = FONT
        x += w + 0.25
    # №2
    t2 = plan['t2']
    s = prs.slides.add_slide(prs.slide_layouts[6]); header(s, f"{lab['task']} 2. {lab['t2']}", '№ 2')
    table(s, 0.6, 1.5, 12.1, t2['table']['head'], t2['table']['rows'], [0.6, 3.2, 2.6, 2.6], 20, 0.7)
    first = t2['a'].split('; ')[0]
    tb(s, 0.6, 6.2, 12.1, 0.8, f"{lab['sample']} (1): {first}", 20, True, GREEN)
    notes(s, f"{lab['ans']}: {t2['a']}")
    # №3
    t4 = plan['t4']
    task_text_slide(prs, lab, f"{lab['task']} 3. {lab['t3']}", '№ 3', t4['q'], t4.get('code'), t4.get('a'))
    # №4
    t3 = plan['t3']
    task_text_slide(prs, lab, f"{lab['task']} 4. {lab['t4']}", '№ 4', t3['q'], t3.get('code'), t3.get('a'), steps=lab['steps'])
    # №5
    fl = plan['fl']
    s = prs.slides.add_slide(prs.slide_layouts[6]); header(s, f"{lab['task']} 5. {lab['fl']}", '№ 5')
    card(s, 0.6, 1.4, 12.1, 2.4)
    tb(s, 0.8, 1.5, 11.7, 2.2, f"{lab['sit']}. {fl['text']}", fit(fl['text'], 22, 15, 420), False, INK)
    head, rows = fl['table']['head'], fl['table']['rows']
    table(s, 0.6, 4.0, 5.9, head, rows[:6], None, 16, 0.42)
    if im.get('chart'): img(s, im['chart'], 6.8, 3.95, 5.9, 3.1)
    notes(s, f"{lab['ans']}:\n" + '\n'.join(f"{i+1}. {x['a']}" for i, x in enumerate(fl['qs'])))
    s = prs.slides.add_slide(prs.slide_layouts[6]); header(s, f"{lab['fl']}: {lab['quest']}", '№ 5')
    card(s, 0.6, 1.4, 12.1, 5.5)
    qs = [f"{i+1}. {x['q']}" for i, x in enumerate(fl['qs'])]
    bullets(s, [q for q in qs], 0.9, 1.55, 11.5, 5.0, fit(' '.join(qs), 26, 18, 420))
    tb(s, 0.9, 6.35, 11.5, 0.5, f"1–2: {lab['cls']}; {lab['quest'].lower()} 3+: {lab['home']}", 16, False, MUTED, italic=True)
    notes(s, f"{lab['ans']}:\n" + '\n'.join(f"{i+1}. {x['a']}" for i, x in enumerate(fl['qs'])))
    reflection(prs, lab)


def reflection(prs, lab):
    s = prs.slides.add_slide(prs.slide_layouts[6]); header(s, lab['refl'], '№ 6')
    y = 1.6
    for t in (lab['r1'], lab['r2'], lab['r3']):
        card(s, 1.2, y, 10.9, 1.2, CARD); tb(s, 1.5, y + 0.15, 10.3, 0.9, t, 30, True, NAVY, anchor=MSO_ANCHOR.MIDDLE); y += 1.45
    tb(s, 1.2, 6.1, 10.9, 0.9, f"{lab['hw']}: {lab['hwt']}", 22, False, INK)
    s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s, NAVY)
    tb(s, 0.8, 2.8, 11.7, 1.6, lab['thanks'], 54, True, RGBColor(255, 255, 255), PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)


def language(prs, plan, lab, args):
    im = plan.get('images', {}); b = plan['bank']; pk = plan['pack']
    title_slide(prs, plan, lab, args)
    goal_slide(prs, plan, lab)
    # №1
    s = prs.slides.add_slide(prs.slide_layouts[6]); header(s, f"{lab['task']} 1. {lab['think']}", '№ 1')
    card(s, 0.6, 1.4, 6.4, 5.5)
    if im.get('map'): img(s, im['map'], 0.75, 1.55, 6.1, 5.2)
    y = 1.5
    for tag, q in ((lab['t_ata'], 'Суретте қандай сөздер бар? Нені көріп тұрсың?'), (lab['t_an'], 'Қай сөздердің мағынасын білесің?'), (lab['t_bol'], 'Сабақ не туралы болады деп ойлайсың?')):
        card(s, 7.3, y, 5.4, 1.65, CARD); tb(s, 7.5, y + 0.08, 5.0, 0.45, tag, 20, True, NAVY); tb(s, 7.5, y + 0.55, 5.0, 1.05, q, 20, False, INK); y += 1.8
    # Ереже
    s = prs.slides.add_slide(prs.slide_layouts[6]); header(s, lab['rules'] + ': ' + plan['title'][:60])
    card(s, 0.6, 1.4, 12.1, 5.5)
    rules = pk['rule']; bullets(s, rules, 0.9, 1.55, 11.5, 4.2, fit(' '.join(rules), 24, 17, 520))
    tb(s, 0.9, 6.2, 11.5, 0.6, f"{lab['terms']}: {pk['terms']}", 16, True, NAVY)
    # Сөздік
    s = prs.slides.add_slide(prs.slide_layouts[6]); header(s, lab['vocab'])
    table(s, 0.6, 1.5, 12.1, [lab['word'], lab['mean']], [[v[0], v[2]] for v in b['vocab']], [3, 7], 22, 0.7)
    # №2
    t2 = plan['t2']
    s = prs.slides.add_slide(prs.slide_layouts[6]); header(s, f"{lab['task']} 2. {lab['match']}", '№ 2')
    card(s, 0.6, 1.4, 7.0, 5.5)
    if im.get('cards'): img(s, im['cards'], 0.75, 1.55, 6.7, 5.2)
    L_ = 'АӘБВГҒД'
    words = [f"{L_[i]}) {w}" for i, w in enumerate(t2['words'])]
    tb(s, 7.9, 1.5, 4.8, 0.5, lab['word'], 20, True, NAVY)
    tb(s, 7.9, 2.1, 4.8, 4.5, words, 24, False, INK)
    notes(s, f"{lab['ans']}: {t2['pic']}. {t2['meanAns']}")
    # №3
    txt = b['text']
    s = prs.slides.add_slide(prs.slide_layouts[6]); header(s, f"{lab['task']} 3. {lab['t3']}", '№ 3')
    card(s, 0.6, 1.4, 12.1, 5.5)
    body = [txt[0], txt[1]]
    tb(s, 0.9, 1.55, 11.5, 5.2, body, fit(' '.join(body), 24, 16, 700), False, INK)
    s = prs.slides.add_slide(prs.slide_layouts[6]); header(s, f"{lab['task']} 3. {lab['quest']}", '№ 3')
    card(s, 0.6, 1.4, 12.1, 5.5)
    qs = [f"{i+1}. {x['q']}" for i, x in enumerate(plan['t1'])]
    bullets(s, qs, 0.9, 1.55, 11.5, 5.2, fit(' '.join(qs), 24, 17, 420))
    notes(s, f"{lab['ans']}:\n" + '\n'.join(f"{i+1}) {x['a']}" for i, x in enumerate(plan['t1'])))
    # №4
    s = prs.slides.add_slide(prs.slide_layouts[6]); header(s, f"{lab['task']} 4. {lab['t4']}", '№ 4')
    card(s, 0.6, 1.4, 12.1, 5.5)
    g = plan['t3']
    lines = [f"{lab['sample']}: {g[0]['s']} → {g[0]['a']}"] + [f"{i+1}. {x['s']}   " + '   '.join(f"{'АӘБВ'[k]}) {o}" for k, o in enumerate(x['o'])) for i, x in enumerate(g)]
    tb(s, 0.9, 1.55, 11.5, 5.2, lines, fit(' '.join(lines), 24, 16, 520), False, INK)
    notes(s, f"{lab['ans']}: " + '; '.join(f"{i+1} – {x['a']}" for i, x in enumerate(g)))
    # №5
    s = prs.slides.add_slide(prs.slide_layouts[6]); header(s, f"{lab['task']} 5. {lab['fl']}", '№ 5')
    card(s, 0.6, 1.4, 7.0, 5.5)
    if im.get('frames'): img(s, im['frames'], 0.75, 1.55, 6.7, 5.2)
    tb(s, 7.9, 1.5, 4.8, 0.5, lab['frames'], 20, True, NAVY)
    tb(s, 7.9, 2.1, 4.8, 0.9, '1) Суреттерді оқиға желісі бойынша ретке қой (1–4), әр суретке 1 сөйлем айт.', 18, False, MUTED)
    tb(s, 7.9, 3.3, 4.8, 3.5, '2) ' + plan['t5']['genre'][1], fit(plan['t5']['genre'][1], 20, 15, 200), False, INK)
    notes(s, f"{lab['ans']}: " + ', '.join(f"{i+1} – {x}" for i, x in enumerate(plan['t5']['seq'])))
    reflection(prs, lab)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('plan'); ap.add_argument('out')
    ap.add_argument('--lang', default='kk', choices=['kk', 'ru'])
    ap.add_argument('--teacher'); ap.add_argument('--school'); ap.add_argument('--subject')
    args = ap.parse_args()
    plan = json.load(open(args.plan, encoding='utf8'))
    lab = L[args.lang]
    prs = Presentation(); prs.slide_width = Inches(W); prs.slide_height = Inches(H)
    (language if plan.get('shape') == 'language' else science)(prs, plan, lab, args)
    prs.save(args.out)
    print('ok', args.out, len(prs.slides._sldIdLst), 'slides')


if __name__ == '__main__':
    main()
