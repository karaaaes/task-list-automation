import os
import io
import json
import re
from datetime import datetime, timedelta, date
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func, case, desc, text as sqltext
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

app = Flask(__name__)
CORS(app)

app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get(
    'DATABASE_URL',
    'postgresql://postgres:postgres@localhost:5432/kara_taskdb'
)
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

GEMINI_API_KEY = os.environ.get('GEMINI_API_KEY', '')
GEMINI_MODEL = os.environ.get('GEMINI_MODEL', 'gemini-2.5-flash')
if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)


# ---------------- MODEL ----------------
class Task(db.Model):
    __tablename__ = 'tasks'
    id = db.Column(db.Integer, primary_key=True)
    detail = db.Column(db.String(500), nullable=False)
    notes = db.Column(db.Text, nullable=True, default='')
    task_date = db.Column(db.Date, nullable=False)
    status = db.Column(db.String(20), nullable=False, default='Planned')
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    modified_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

    def to_dict(self):
        return {
            'id': self.id,
            'detail': self.detail,
            'notes': self.notes or '',
            'task_date': self.task_date.isoformat() if self.task_date else None,
            'status': self.status,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'modified_at': self.modified_at.isoformat() if self.modified_at else None,
        }


STATUS_ORDER = ['Planned', 'Completed', 'Cancelled']


def status_sort_expr():
    return case(
        (Task.status == 'Planned', 0),
        (Task.status == 'Completed', 1),
        (Task.status == 'Cancelled', 2),
        else_=3,
    )


# ---------------- ROUTES: TASKS ----------------
@app.route('/api/health')
def health():
    return jsonify({'status': 'ok'})


@app.route('/api/tasks', methods=['GET'])
def list_tasks():
    page = int(request.args.get('page', 1))
    per_page = int(request.args.get('per_page', 10))
    search = request.args.get('search', '').strip()
    search_date = request.args.get('search_date', '').strip()
    sort_by = request.args.get('sort_by', 'default')
    status_filter = request.args.get('status', '').strip()

    query = Task.query
    if search:
        query = query.filter(Task.detail.ilike(f'%{search}%'))
    if search_date:
        try:
            d = datetime.strptime(search_date, '%Y-%m-%d').date()
            query = query.filter(func.date(Task.created_at) == d)
        except ValueError:
            pass
    if status_filter and status_filter in STATUS_ORDER:
        query = query.filter(Task.status == status_filter)

    if sort_by == 'created':
        query = query.order_by(desc(Task.created_at))
    elif sort_by == 'modified':
        query = query.order_by(desc(Task.modified_at))
    elif sort_by == 'status':
        query = query.order_by(status_sort_expr(), desc(Task.created_at))
    else:
        query = query.order_by(status_sort_expr(), desc(Task.created_at))

    total = query.count()
    items = query.offset((page - 1) * per_page).limit(per_page).all()
    return jsonify({
        'items': [t.to_dict() for t in items],
        'total': total, 'page': page, 'per_page': per_page,
        'pages': (total + per_page - 1) // per_page if per_page else 1,
    })


@app.route('/api/tasks/<int:task_id>', methods=['GET'])
def get_task(task_id):
    return jsonify(Task.query.get_or_404(task_id).to_dict())


@app.route('/api/tasks', methods=['POST'])
def create_task():
    data = request.get_json() or {}
    detail = (data.get('detail') or '').strip()
    notes = (data.get('notes') or '').strip()
    task_date_str = data.get('task_date')
    status = data.get('status', 'Planned')

    if not detail or not task_date_str:
        return jsonify({'error': 'detail and task_date are required'}), 400
    if status not in STATUS_ORDER:
        status = 'Planned'
    try:
        task_date = datetime.strptime(task_date_str, '%Y-%m-%d').date()
    except ValueError:
        return jsonify({'error': 'invalid task_date format'}), 400

    t = Task(detail=detail, notes=notes, task_date=task_date, status=status)
    db.session.add(t)
    db.session.commit()
    return jsonify(t.to_dict()), 201


@app.route('/api/tasks/<int:task_id>', methods=['PUT'])
def update_task(task_id):
    t = Task.query.get_or_404(task_id)
    data = request.get_json() or {}

    if 'detail' in data:
        t.detail = (data['detail'] or '').strip() or t.detail
    if 'notes' in data:
        t.notes = (data['notes'] or '').strip()
    if 'task_date' in data and data['task_date']:
        try:
            t.task_date = datetime.strptime(data['task_date'], '%Y-%m-%d').date()
        except ValueError:
            return jsonify({'error': 'invalid task_date format'}), 400
    if 'status' in data and data['status'] in STATUS_ORDER:
        t.status = data['status']

    db.session.commit()
    return jsonify(t.to_dict())


@app.route('/api/tasks/<int:task_id>/status', methods=['PATCH'])
def update_status(task_id):
    t = Task.query.get_or_404(task_id)
    data = request.get_json() or {}
    new_status = data.get('status')
    if new_status not in STATUS_ORDER:
        return jsonify({'error': 'invalid status'}), 400
    t.status = new_status
    db.session.commit()
    return jsonify(t.to_dict())


@app.route('/api/tasks/<int:task_id>', methods=['DELETE'])
def delete_task(task_id):
    t = Task.query.get_or_404(task_id)
    db.session.delete(t)
    db.session.commit()
    return jsonify({'success': True})


# ---------------- DASHBOARD ----------------
@app.route('/api/dashboard/summary', methods=['GET'])
def dashboard_summary():
    total = db.session.query(func.count(Task.id)).scalar() or 0
    counts = dict(db.session.query(Task.status, func.count(Task.id)).group_by(Task.status).all())
    by_status = {s: int(counts.get(s, 0)) for s in STATUS_ORDER}

    today = date.today()
    start = today - timedelta(days=29)
    rows = db.session.query(
        func.date(Task.created_at).label('d'),
        func.count(Task.id)
    ).filter(func.date(Task.created_at) >= start).group_by('d').all()
    rows_map = {r[0].isoformat() if hasattr(r[0], 'isoformat') else str(r[0]): int(r[1]) for r in rows}

    series = []
    for i in range(30):
        day = start + timedelta(days=i)
        series.append({'date': day.isoformat(), 'count': rows_map.get(day.isoformat(), 0)})

    return jsonify({'total': int(total), 'by_status': by_status, 'chart': series})


# ============================================================
#                    PPT GENERATION (DARK BROWN)
# ============================================================
BROWN = {
    'bg':          RGBColor(0xF5, 0xEE, 0xE2),  # warm ivory/cream background
    'card':        RGBColor(0xFF, 0xFB, 0xF3),  # near-white warm
    'card_alt':    RGBColor(0xEC, 0xE2, 0xD0),  # tan
    'espresso':    RGBColor(0x2B, 0x18, 0x10),  # deepest brown, primary text
    'coffee':      RGBColor(0x4A, 0x2C, 0x20),  # rich brown headers
    'mocha':       RGBColor(0x6B, 0x44, 0x23),  # medium brown
    'umber':       RGBColor(0x8B, 0x6F, 0x47),  # muted secondary
    'gold':        RGBColor(0xC9, 0xA9, 0x61),  # accent gold
    'gold_soft':   RGBColor(0xE3, 0xCC, 0x95),  # soft gold tint
    'copper':      RGBColor(0xB8, 0x73, 0x33),  # copper accent
    'parchment':   RGBColor(0xEF, 0xE5, 0xD3),  # subtle separator
    'muted':       RGBColor(0x8C, 0x7B, 0x68),  # muted text
    'white':       RGBColor(0xFF, 0xFF, 0xFF),
}

STATUS_BADGE = {
    'Planned':   {'bg': RGBColor(0xE3, 0xCC, 0x95), 'fg': RGBColor(0x4A, 0x2C, 0x20)},
    'Completed': {'bg': RGBColor(0x6B, 0x44, 0x23), 'fg': RGBColor(0xF5, 0xEE, 0xE2)},
    'Cancelled': {'bg': RGBColor(0xC5, 0xB0, 0x9A), 'fg': RGBColor(0x4A, 0x2C, 0x20)},
}


def set_slide_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_rect(slide, left, top, width, height, fill_color, line_color=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_color
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
    shp.shadow.inherit = False
    return shp


def add_rounded(slide, left, top, width, height, fill_color, line_color=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_color
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
    shp.shadow.inherit = False
    return shp


def add_text(slide, left, top, width, height, text, size=14, bold=False,
             color=None, align=PP_ALIGN.LEFT, font='Calibri', italic=False):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0); tf.margin_right = Emu(0)
    tf.margin_top = Emu(0);  tf.margin_bottom = Emu(0)
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font
    if color is not None:
        run.font.color.rgb = color
    return tb


def add_section_header(slide, title_text, subtitle_text=''):
    """Top header used on every content slide: thin gold accent bar + title."""
    # gold accent line
    add_rect(slide, Inches(0.6), Inches(0.55), Inches(0.5), Inches(0.06), BROWN['gold'])
    add_text(slide, Inches(0.6), Inches(0.7), Inches(12), Inches(0.6),
             title_text, size=26, bold=True, color=BROWN['espresso'])
    if subtitle_text:
        add_text(slide, Inches(0.6), Inches(1.2), Inches(12), Inches(0.4),
                 subtitle_text, size=12, color=BROWN['muted'], italic=True)
    # subtle bottom divider
    add_rect(slide, Inches(0.6), Inches(1.7), Inches(12.1), Inches(0.015), BROWN['parchment'])
    # page footer brand
    add_text(slide, Inches(0.6), Inches(7.05), Inches(6), Inches(0.3),
             'KARA TASK LIST', size=9, bold=True, color=BROWN['umber'])
    add_text(slide, Inches(6.7), Inches(7.05), Inches(6), Inches(0.3),
             'Confidential', size=9, color=BROWN['muted'], align=PP_ALIGN.RIGHT)


def build_ppt(plan, tasks, date_from, date_to, status_filter):
    """Build a professional dark-brown themed PPT."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    title = plan.get('title', 'Task Report')
    subtitle = plan.get('subtitle', '')
    summary = plan.get('summary', '')
    highlights = plan.get('highlights', []) or []
    recommendations = plan.get('recommendations', []) or []

    # =========================================================
    # SLIDE 1: COVER
    # =========================================================
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, BROWN['bg'])

    # left dark band
    add_rect(s, Inches(0), Inches(0), Inches(4.5), Inches(7.5), BROWN['espresso'])
    # gold vertical accent
    add_rect(s, Inches(4.5), Inches(0), Inches(0.08), Inches(7.5), BROWN['gold'])

    # brand block on dark side
    add_rect(s, Inches(0.6), Inches(0.7), Inches(0.4), Inches(0.05), BROWN['gold'])
    add_text(s, Inches(0.6), Inches(0.85), Inches(3.5), Inches(0.4),
             'KARA TASK LIST', size=12, bold=True, color=BROWN['gold'])
    add_text(s, Inches(0.6), Inches(1.2), Inches(3.5), Inches(0.4),
             'Executive Report', size=10, color=BROWN['gold_soft'], italic=True)

    # bottom info on dark
    add_text(s, Inches(0.6), Inches(6.4), Inches(3.5), Inches(0.3),
             'PERIOD', size=9, bold=True, color=BROWN['gold'])
    add_text(s, Inches(0.6), Inches(6.7), Inches(3.5), Inches(0.4),
             f'{date_from}  โ€”  {date_to}', size=12, color=BROWN['bg'])

    # title on right (light) side
    add_text(s, Inches(5.2), Inches(2.6), Inches(7.8), Inches(0.4),
             'TASK PERFORMANCE REPORT', size=11, bold=True, color=BROWN['copper'])
    # thin gold divider above title
    add_rect(s, Inches(5.2), Inches(3.05), Inches(0.8), Inches(0.04), BROWN['gold'])
    add_text(s, Inches(5.2), Inches(3.2), Inches(7.8), Inches(1.6),
             title, size=40, bold=True, color=BROWN['espresso'])
    if subtitle:
        add_text(s, Inches(5.2), Inches(4.7), Inches(7.8), Inches(0.8),
                 subtitle, size=15, color=BROWN['mocha'], italic=True)

    if status_filter:
        add_text(s, Inches(5.2), Inches(5.6), Inches(7.8), Inches(0.4),
                 f'Status filter: {status_filter}', size=11, color=BROWN['umber'])
    add_text(s, Inches(5.2), Inches(6.7), Inches(7.8), Inches(0.4),
             f'Prepared {date.today().isoformat()}', size=10, color=BROWN['muted'], italic=True)

    # =========================================================
    # SLIDE 2: EXECUTIVE SUMMARY
    # =========================================================
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, BROWN['bg'])
    add_section_header(s, 'Executive Summary', 'Period-over-period performance at a glance')

    # KPI strip - 4 cards in single row
    total = len(tasks)
    by_status = {st: 0 for st in STATUS_ORDER}
    for t in tasks:
        by_status[t['status']] = by_status.get(t['status'], 0) + 1
    completion_rate = round((by_status['Completed'] / total) * 100) if total else 0

    kpis = [
        ('TOTAL TASKS', str(total)),
        ('PLANNED', str(by_status.get('Planned', 0))),
        ('COMPLETED', str(by_status.get('Completed', 0))),
        ('COMPLETION', f'{completion_rate}%'),
    ]
    card_w = Inches(2.95)
    card_h = Inches(1.4)
    gap = Inches(0.15)
    start_x = Inches(0.6)
    for i, (label, val) in enumerate(kpis):
        left = start_x + (card_w + gap) * i
        # dark card with gold top stripe
        add_rect(s, left, Inches(2.0), card_w, card_h, BROWN['card'])
        add_rect(s, left, Inches(2.0), card_w, Inches(0.06), BROWN['gold'])
        add_text(s, left + Inches(0.3), Inches(2.2), card_w - Inches(0.6), Inches(0.35),
                 label, size=9, bold=True, color=BROWN['umber'])
        add_text(s, left + Inches(0.3), Inches(2.55), card_w - Inches(0.6), Inches(0.85),
                 val, size=36, bold=True, color=BROWN['espresso'])

    # Summary panel
    add_rect(s, Inches(0.6), Inches(3.7), Inches(12.1), Inches(2.95), BROWN['card'])
    add_rect(s, Inches(0.6), Inches(3.7), Inches(0.08), Inches(2.95), BROWN['copper'])
    add_text(s, Inches(0.9), Inches(3.9), Inches(11.5), Inches(0.4),
             'SUMMARY', size=10, bold=True, color=BROWN['copper'])
    add_text(s, Inches(0.9), Inches(4.3), Inches(11.5), Inches(2.2),
             summary or 'No summary available.', size=13, color=BROWN['espresso'])

    # =========================================================
    # SLIDE 3+: TASK LIST (paginated, 10 per slide)
    # =========================================================
    chunk = 10
    total_pages = max(1, (len(tasks) + chunk - 1) // chunk)
    for page_i in range(0, max(len(tasks), 1), chunk):
        s = prs.slides.add_slide(blank)
        set_slide_bg(s, BROWN['bg'])
        current_page = page_i // chunk + 1
        subtitle_str = f'Page {current_page} of {total_pages}  ยท  Showing {page_i+1}-{min(page_i+chunk, len(tasks))} of {len(tasks)}'
        add_section_header(s, 'Task List', subtitle_str)

        # Table header row
        header_y = Inches(1.95)
        add_rect(s, Inches(0.6), header_y, Inches(12.1), Inches(0.45), BROWN['espresso'])
        add_text(s, Inches(0.75), header_y + Inches(0.1), Inches(0.5), Inches(0.3),
                 '#', size=10, bold=True, color=BROWN['gold'])
        add_text(s, Inches(1.3), header_y + Inches(0.1), Inches(5.0), Inches(0.3),
                 'TASK DETAIL', size=10, bold=True, color=BROWN['gold'])
        add_text(s, Inches(6.4), header_y + Inches(0.1), Inches(3.5), Inches(0.3),
                 'NOTES', size=10, bold=True, color=BROWN['gold'])
        add_text(s, Inches(10.0), header_y + Inches(0.1), Inches(1.3), Inches(0.3),
                 'DATE', size=10, bold=True, color=BROWN['gold'])
        add_text(s, Inches(11.4), header_y + Inches(0.1), Inches(1.3), Inches(0.3),
                 'STATUS', size=10, bold=True, color=BROWN['gold'], align=PP_ALIGN.CENTER)

        row_y = header_y + Inches(0.45)
        row_h = Inches(0.42)
        page_tasks = tasks[page_i:page_i+chunk]
        for ridx, t in enumerate(page_tasks):
            row_bg = BROWN['card'] if ridx % 2 == 0 else BROWN['card_alt']
            add_rect(s, Inches(0.6), row_y, Inches(12.1), row_h, row_bg)

            add_text(s, Inches(0.75), row_y + Inches(0.1), Inches(0.5), Inches(0.3),
                     str(page_i + ridx + 1), size=10, bold=True, color=BROWN['mocha'])

            detail = t['detail']
            if len(detail) > 50:
                detail = detail[:47] + '...'
            add_text(s, Inches(1.3), row_y + Inches(0.1), Inches(5.0), Inches(0.3),
                     detail, size=10, color=BROWN['espresso'])

            note_preview = (t.get('notes') or '').replace('\n', ' ').strip()
            if not note_preview:
                note_preview = 'โ€”'
                note_color = BROWN['muted']
            else:
                if len(note_preview) > 38:
                    note_preview = note_preview[:35] + '...'
                note_color = BROWN['mocha']
            add_text(s, Inches(6.4), row_y + Inches(0.1), Inches(3.5), Inches(0.3),
                     note_preview, size=9, color=note_color, italic=True)

            add_text(s, Inches(10.0), row_y + Inches(0.1), Inches(1.3), Inches(0.3),
                     t.get('task_date') or '-', size=10, color=BROWN['mocha'])

            st = t['status']
            badge_colors = STATUS_BADGE.get(st, STATUS_BADGE['Planned'])
            add_rounded(s, Inches(11.45), row_y + Inches(0.08), Inches(1.2), Inches(0.26), badge_colors['bg'])
            add_text(s, Inches(11.45), row_y + Inches(0.1), Inches(1.2), Inches(0.24),
                     st.upper(), size=8, bold=True, color=badge_colors['fg'], align=PP_ALIGN.CENTER)

            row_y += row_h

    # =========================================================
    # SLIDE: HIGHLIGHTS (after Task List)
    # =========================================================
    if highlights:
        s = prs.slides.add_slide(blank)
        set_slide_bg(s, BROWN['bg'])
        add_section_header(s, 'Notable Highlights', 'Tasks of particular significance during this period')

        y = Inches(2.0)
        for i, h in enumerate(highlights[:6]):
            add_rect(s, Inches(0.6), y, Inches(12.1), Inches(0.7), BROWN['card'])
            # gold number badge
            add_rect(s, Inches(0.6), y, Inches(0.08), Inches(0.7), BROWN['gold'])
            add_text(s, Inches(0.85), y + Inches(0.18), Inches(0.5), Inches(0.4),
                     f'{i+1:02d}', size=14, bold=True, color=BROWN['copper'])
            add_text(s, Inches(1.5), y + Inches(0.2), Inches(11.0), Inches(0.4),
                     h, size=12, color=BROWN['espresso'])
            y += Inches(0.82)

    # =========================================================
    # SLIDE: RECOMMENDATIONS
    # =========================================================
    if recommendations:
        s = prs.slides.add_slide(blank)
        set_slide_bg(s, BROWN['bg'])
        add_section_header(s, 'Recommendations', 'Actionable next steps for the upcoming period')

        # 2-column layout
        cols = 2
        card_w2 = Inches(5.95)
        card_h2 = Inches(1.3)
        gap_x = Inches(0.2)
        gap_y = Inches(0.2)
        for idx, r in enumerate(recommendations[:6]):
            col = idx % cols
            row = idx // cols
            left = Inches(0.6) + (card_w2 + gap_x) * col
            top = Inches(2.0) + (card_h2 + gap_y) * row
            add_rect(s, left, top, card_w2, card_h2, BROWN['card'])
            add_rect(s, left, top, Inches(0.06), card_h2, BROWN['copper'])
            add_text(s, left + Inches(0.3), top + Inches(0.15), Inches(0.6), Inches(0.4),
                     f'{idx+1:02d}', size=18, bold=True, color=BROWN['copper'])
            add_text(s, left + Inches(0.95), top + Inches(0.2), card_w2 - Inches(1.2), card_h2 - Inches(0.3),
                     r, size=11, color=BROWN['espresso'])

    # =========================================================
    # CLOSING SLIDE
    # =========================================================
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, BROWN['espresso'])
    # gold horizontal accent
    add_rect(s, Inches(6.0), Inches(3.0), Inches(1.3), Inches(0.06), BROWN['gold'])
    add_text(s, Inches(1), Inches(3.2), Inches(11.3), Inches(1.0),
             'Thank You', size=54, bold=True, color=BROWN['bg'], align=PP_ALIGN.CENTER)
    add_text(s, Inches(1), Inches(4.2), Inches(11.3), Inches(0.5),
             'Questions & Discussion', size=16, color=BROWN['gold_soft'], italic=True, align=PP_ALIGN.CENTER)
    add_text(s, Inches(1), Inches(6.8), Inches(11.3), Inches(0.4),
             'KARA TASK LIST  ยท  Generated for executive review', size=9, color=BROWN['umber'], align=PP_ALIGN.CENTER)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ============================================================
#                    GEMINI PLAN CALL
# ============================================================
def call_gemini_for_plan(tasks, date_from, date_to, status_filter):
    """Ask Gemini to produce a professional, executive-grade content plan."""
    if not GEMINI_API_KEY:
        return {
            'title': 'Task Performance Report',
            'subtitle': f'Period {date_from} to {date_to}',
            'summary': f'This report covers {len(tasks)} task(s) for the selected period.',
            'highlights': [t['detail'] for t in tasks[:5]],
            'recommendations': [
                'Review pending Planned tasks.',
                'Capture lessons learned from Completed items.',
            ],
        }

    task_lines = []
    for t in tasks:
        line = f"- [{t['status']}] {t['task_date']}: {t['detail']}"
        n = (t.get('notes') or '').strip()
        if n:
            n_short = n.replace('\n', ' ')
            if len(n_short) > 200:
                n_short = n_short[:200] + '...'
            line += f"  | notes: {n_short}"
        task_lines.append(line)
    tasks_text = "\n".join(task_lines) if task_lines else "(no tasks)"
    status_text = status_filter if status_filter else 'All'

    system_prompt = (
        "You are a senior management consultant preparing a board-level executive briefing. "
        "Your tone is formal, precise, and outcome-oriented. You synthesize task data into "
        "insights a busy executive can scan in seconds. You never use emojis, exclamation "
        "marks, or casual language. Reply with valid JSON only."
    )

    user_prompt = f"""Produce an executive briefing plan for the following task portfolio.

PERIOD: {date_from} to {date_to}
STATUS FILTER: {status_text}
TASK COUNT: {len(tasks)}

TASKS (each line may include contextual notes after the pipe):
{tasks_text}

Return a JSON object with EXACTLY these keys:
{{
  "title": "string",
  "subtitle": "string",
  "summary": "string",
  "highlights": ["string", "string", ...],
  "recommendations": ["string", "string", ...]
}}

Guidelines:
- title: concise, professional, max 7 words. Examples: "Q2 Operations Review", "Strategic Task Portfolio Update". Avoid hype.
- subtitle: one-line context describing the scope and posture, max 14 words.
- summary: 4-6 sentences. Open with the headline finding. Cover completion posture, themes from task notes, any risks or blockers visible in the data, and forward outlook. Reference real counts. Formal third-person tone.
- highlights: 3-5 specific tasks worth surfacing to leadership. Rewrite each as a crisp one-line insight (not a verbatim copy). Lead with the outcome or impact. Use the notes to add substance where available. Max 20 words each.
- recommendations: 3-5 actionable next steps grounded in the data. Each should be specific enough to assign to an owner. Max 18 words each.

Critical rules:
- No emojis, no markdown, no code fences.
- No filler words like "delve", "leverage", "synergy".
- Output raw JSON only.
"""

    try:
        gen_config = {
            "temperature": 0.5,
            "max_output_tokens": 8192,
            "response_mime_type": "application/json",
        }
        try:
            from google.generativeai.types import GenerationConfig
            gen_config_obj = GenerationConfig(
                temperature=0.5,
                max_output_tokens=8192,
                response_mime_type="application/json",
                thinking_config={"thinking_budget": 0},
            )
            model = genai.GenerativeModel(
                model_name=GEMINI_MODEL,
                system_instruction=system_prompt,
                generation_config=gen_config_obj,
            )
        except (TypeError, ImportError):
            model = genai.GenerativeModel(
                model_name=GEMINI_MODEL,
                system_instruction=system_prompt,
                generation_config=gen_config,
            )

        resp = model.generate_content(user_prompt)
        text = (resp.text or '').strip()
        text = re.sub(r'^```(?:json)?\s*|\s*```$', '', text, flags=re.MULTILINE)
        return json.loads(text)
    except Exception as e:
        print(f"[Gemini error] {e}")
        return {
            'title': 'Task Performance Report',
            'subtitle': f'Period {date_from} to {date_to}',
            'summary': f'This report covers {len(tasks)} task(s). (AI generation unavailable: {e})',
            'highlights': [t['detail'] for t in tasks[:5]],
            'recommendations': ['Review the task list manually.'],
        }


@app.route('/api/generate-ppt', methods=['POST'])
def generate_ppt():
    data = request.get_json() or {}
    date_from = data.get('date_from')
    date_to = data.get('date_to')
    status_filter = data.get('status', '').strip()

    if not date_from or not date_to:
        return jsonify({'error': 'date_from and date_to are required'}), 400
    try:
        df = datetime.strptime(date_from, '%Y-%m-%d').date()
        dt = datetime.strptime(date_to, '%Y-%m-%d').date()
    except ValueError:
        return jsonify({'error': 'invalid date format'}), 400

    q = Task.query.filter(Task.task_date >= df, Task.task_date <= dt)
    if status_filter and status_filter in STATUS_ORDER:
        q = q.filter(Task.status == status_filter)
    q = q.order_by(status_sort_expr(), Task.task_date.asc())
    tasks = [t.to_dict() for t in q.all()]

    if not tasks:
        return jsonify({'error': 'No tasks found for the selected range/status.'}), 404

    plan = call_gemini_for_plan(tasks, date_from, date_to, status_filter)
    buf = build_ppt(plan, tasks, date_from, date_to, status_filter)
    filename = f"kara_task_report_{date_from}_to_{date_to}.pptx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
    )


# ---------------- INIT ----------------
def init_db():
    with app.app_context():
        db.create_all()
        # Auto-migrate: add notes column if it doesn't exist (for existing DBs)
        try:
            db.session.execute(sqltext("ALTER TABLE tasks ADD COLUMN IF NOT EXISTS notes TEXT DEFAULT ''"))
            db.session.commit()
        except Exception as e:
            db.session.rollback()
            print(f"[init_db] migration skipped: {e}")


if __name__ == '__main__':
    init_db()
    app.run(host='0.0.0.0', port=8888, debug=False)
