#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
disclosure_scan.py — الطبقة الحتمية (A) لمِحَكّ الأخلاقيات والنزاهة.

يفحص ملف بحثٍ (.docx أو نصّ) ويجيب — دون إنترنت ودون أي تبعيّات (Python
القياسي فقط) — عن ثلاثة أسئلة حتميّة قابلة للتحقق آلياً:

  A1) جرد الإفصاح: أيّ بيانات النزاهة حاضرة نصّاً؟ (موافقة أخلاقية/لجنة،
      موافقة مستنيرة، تضارب مصالح، تمويل، توافر بيانات، مساهمة المؤلفين،
      الإفصاح عن استخدام الذكاء الاصطناعي، تسجيل مسبق/تسجيل تجربة).
  A2) اتساق الإفصاح داخلياً: تناقضات قابلة للكشف رياضياً/منطقياً —
      تاريخ الموافقة الأخلاقية مقابل نافذة جمع البيانات، تعارض «لا تضارب»
      مع وجود تمويل، «البيانات متاحة عند الطلب» بلا جهة اتصال، وصلاحية
      بنية DOI/رابط المستودع (صلاحية شكلية فقط — لا تحقّق من الحلّ).
  A3) بصمات النزاهة (رايات لا أحكام): عبارات مُعذَّبة (tortured phrases)
      في النصّ الإنجليزي، وقرائن مطابع الورق (paper mills)، وقرائن التقطيع
      والنشر المكرّر (salami). كلها رايات تستدعي نظراً بشرياً، لا اتهامات.

الفرز أولاً: أيّ نقصٍ في الإفصاح يُقاس بمصفوفة العائلة المعرفية. غيابُ
لجنة أخلاقيات في رسالة فقهٍ ليس عيباً بل هو المتوقَّع؛ لذا يقبل السكربت
وسمَ العائلة ويكتم الرايات غير المنطبقة. مرّر --family لضبط ذلك.

ما لا يفعله (يُعلَن صراحةً):
  - لا يُثبت سوء سلوك ولا ينفيه. غيابُ الدليل ليس دليل الغياب.
  - لا يتّهم مصدراً أو باحثاً؛ يرفع قرائن قابلة للتتبّع فقط.
  - لا يحلّ روابط ولا DOI عبر الشبكة (تلك الطبقة B، بأدوات النموذج).
  - لا يعيد التدقيق الإحصائي (تلك مهمة analysis-results-auditor).

الاستعمال:
    python disclosure_scan.py paper.docx
    python disclosure_scan.py paper.docx --family legal --json out.json
    python disclosure_scan.py --self-test

العائلات المقبولة لـ --family (متطابقة مع أخوات المِحَكّ):
    quantitative | qualitative | hermeneutic | historical
    comparative  | legal       | mixed       | unknown(افتراضي)
"""
import sys, re, json, argparse, zipfile
from html import unescape

# ════════════════════════════════════════════════════════════════════════
# القسم 0: قراءة النصّ (docx/نصّ) + سحب حقول وورد الرسمية
# ════════════════════════════════════════════════════════════════════════

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def read_docx_text(path):
    """يجمع نصّ المتن والترويسة والتذييل. يفصل الفقرات بأسطر والخلايا بجدولة."""
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist()
                 if n == 'word/document.xml'
                 or re.match(r'word/(header|footer)\d*\.xml', n)]
        chunks = []
        for n in names:
            xml = z.read(n).decode('utf-8', 'ignore')
            xml = re.sub(r'</w:p>', '\n', xml)
            xml = re.sub(r'</w:tc>', '\t', xml)
            texts = re.findall(r'<w:t\b[^>]*>(.*?)</w:t>', xml, re.S)
            chunks.append(unescape(''.join(texts)))
    return '\n'.join(chunks)


def read_docx_hyperlinks(path):
    """يسحب أهداف الروابط التشعّبية من علاقات الملف (للطبقة A2: صلاحية الرابط)."""
    urls = []
    try:
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                if n.endswith('.rels'):
                    rel = z.read(n).decode('utf-8', 'ignore')
                    urls += re.findall(r'Target="(https?://[^"]+)"', rel)
    except Exception:
        pass
    return urls


def load(path):
    if path.lower().endswith('.docx'):
        return read_docx_text(path), read_docx_hyperlinks(path)
    with open(path, encoding='utf-8') as f:
        text = f.read()
    # اسحب أي روابط ظاهرة في النصّ الخام أيضاً
    urls = re.findall(r'https?://[^\s)>\]]+', text)
    return text, urls


# ════════════════════════════════════════════════════════════════════════
# القسم 1: مصفوفة انطباق الإفصاح على العائلة المعرفية (الفرز أولاً)
# ════════════════════════════════════════════════════════════════════════
# لكل بند إفصاح: على أيّ العائلات يُتوقَّع؟ الغياب في عائلة لا تتوقّعه ليس عيباً.
# القيم: "required" يُتوقّع بقوة | "conditional" يعتمد على وجود شرطٍ (بيانات
# بشرية/تمويل...) | "na" غير منطبق أصلاً على هذه العائلة.

FAMILIES = ['quantitative', 'qualitative', 'hermeneutic', 'historical',
            'comparative', 'legal', 'mixed', 'unknown']

# البنود الثمانية
DISCLOSURES = ['ethics_approval', 'informed_consent', 'coi', 'funding',
               'data_availability', 'author_contributions', 'ai_use',
               'preregistration']

DISCLOSURE_AR = {
    'ethics_approval':      'الموافقة الأخلاقية / لجنة أخلاقيات البحث',
    'informed_consent':     'الموافقة المستنيرة للمشاركين',
    'coi':                  'الإفصاح عن تضارب المصالح',
    'funding':              'الإفصاح عن مصدر التمويل',
    'data_availability':    'بيان توافر البيانات',
    'author_contributions': 'بيان مساهمة المؤلفين',
    'ai_use':               'الإفصاح عن استخدام الذكاء الاصطناعي',
    'preregistration':      'التسجيل المسبق / تسجيل التجربة',
}

# مصفوفة الانطباق. "conditional" معناه: لا يُتوقّع إلا إن تحقّق شرطه
# (كوجود مشاركين بشر، أو وجود تمويل مُعلَن) — والسكربت يكشف الشرط بنفسه حين يمكن.
APPLICABILITY = {
    #                     ethics  consent   coi     fund    data    contrib ai      prereg
    'quantitative': dict(ethics_approval='conditional', informed_consent='conditional',
                         coi='required', funding='required', data_availability='required',
                         author_contributions='conditional', ai_use='conditional',
                         preregistration='conditional'),
    'qualitative':  dict(ethics_approval='conditional', informed_consent='required',
                         coi='required', funding='conditional', data_availability='conditional',
                         author_contributions='conditional', ai_use='conditional',
                         preregistration='na'),
    'hermeneutic':  dict(ethics_approval='na', informed_consent='na',
                         coi='conditional', funding='conditional', data_availability='na',
                         author_contributions='conditional', ai_use='conditional',
                         preregistration='na'),
    'historical':   dict(ethics_approval='na', informed_consent='na',
                         coi='conditional', funding='conditional', data_availability='conditional',
                         author_contributions='conditional', ai_use='conditional',
                         preregistration='na'),
    'comparative':  dict(ethics_approval='conditional', informed_consent='conditional',
                         coi='conditional', funding='conditional', data_availability='conditional',
                         author_contributions='conditional', ai_use='conditional',
                         preregistration='na'),
    'legal':        dict(ethics_approval='na', informed_consent='na',
                         coi='conditional', funding='conditional', data_availability='na',
                         author_contributions='conditional', ai_use='conditional',
                         preregistration='na'),
    'mixed':        dict(ethics_approval='conditional', informed_consent='required',
                         coi='required', funding='required', data_availability='required',
                         author_contributions='conditional', ai_use='conditional',
                         preregistration='conditional'),
    'unknown':      dict(ethics_approval='conditional', informed_consent='conditional',
                         coi='conditional', funding='conditional', data_availability='conditional',
                         author_contributions='conditional', ai_use='conditional',
                         preregistration='conditional'),
}


# ════════════════════════════════════════════════════════════════════════
# القسم 2: A1 — جرد حضور بنود الإفصاح (عربي + إنجليزي)
# ════════════════════════════════════════════════════════════════════════
# لكل بند: أنماط تدلّ على وجود عبارة إفصاح صريحة. الكشف عن الحضور لا المضمون.
# ملاحظة: نبحث عن العبارة الاصطلاحية، لا مجرد الكلمة، لتقليل الإيجابيات الكاذبة.

PATTERNS = {
    'ethics_approval': [
        r'لجنة\s+(?:أخلاقيات|الأخلاقيات|أخلاقية)',
        r'الموافقة\s+الأخلاقية', r'موافقة\s+أخلاقية',
        r'المصادقة\s+الأخلاقية', r'إجازة\s+أخلاقية',
        r'\beth(?:ics|ical)\s+(?:approval|committee|clearance|board)\b',
        r'\bIRB\b', r'\binstitutional\s+review\s+board\b',
        r'\bethics\s+statement\b',
    ],
    'informed_consent': [
        r'الموافقة\s+المستنيرة', r'موافقة\s+مستنيرة',
        r'الموافقة\s+المسبقة\s+للمشارك', r'موافقة\s+المشاركين',
        r'إقرار\s+المشارك', r'التراضي\s+المستنير',
        r'\binformed\s+consent\b', r'\bconsent\s+(?:was\s+)?obtained\b',
        r'\bwritten\s+consent\b',
    ],
    'coi': [
        r'تضارب\s+(?:المصالح|في\s+المصالح)', r'تعارض\s+المصالح',
        r'تضارب\s+مصالح', r'لا\s+يوجد\s+تضارب',
        r'\bconflict[s]?\s+of\s+interest\b', r'\bcompeting\s+interest[s]?\b',
        r'\bno\s+conflict\b', r'\bdeclare[sd]?\s+no\b',
    ],
    'funding': [
        r'(?:مصدر|جهة)\s+التمويل', r'تمويل\s+(?:البحث|الدراسة)',
        r'المنحة\s+البحثية', r'دعم\s+مالي', r'مُموّل\s+من',
        r'لم\s+(?:يتلقَّ|يتلق|يحصل).{0,30}تمويل',
        r'\bfunding\b', r'\bgrant\s+(?:number|no|#)\b',
        r'\bfinancial\s+support\b', r'\bfunded\s+by\b',
        r'\bno\s+funding\b', r'\breceived\s+no\s+(?:specific\s+)?funding\b',
    ],
    'data_availability': [
        r'توافر\s+البيانات', r'إتاحة\s+البيانات', r'البيانات\s+متاحة',
        r'بيانات\s+الدراسة\s+متاحة', r'مستودع\s+البيانات',
        r'\bdata\s+availability\b', r'\bdata\s+(?:are|is)\s+available\b',
        r'\bavailable\s+(?:up)?on\s+(?:reasonable\s+)?request\b',
        r'\bsupplementary\s+data\b', r'\brepository\b',
    ],
    'author_contributions': [
        r'مساهمة\s+(?:المؤلفين|الباحثين)', r'إسهام\s+المؤلفين',
        r'دور\s+كل\s+(?:مؤلف|باحث)',
        r'\bauthor\s+contributions?\b', r'\bCRediT\b',
        r'\bcontributor[s]?\s+roles?\b',
    ],
    'ai_use': [
        r'استخدام\s+الذكاء\s+الاصطناعي', r'أدوات\s+الذكاء\s+الاصطناعي',
        r'نماذج\s+لغوية', r'الذكاء\s+الاصطناعي\s+التوليدي',
        r'\bgenerative\s+AI\b', r'\bAI[- ]assisted\b',
        r'\blarge\s+language\s+model[s]?\b', r'\bChatGPT\b', r'\bGPT-\d',
        r'\bAI\s+(?:tool|use|disclosure)\b',
    ],
    'preregistration': [
        r'التسجيل\s+المسبق', r'تسجيل\s+(?:مسبق|التجربة)',
        r'بروتوكول\s+مُسجَّل', r'مُسجَّل\s+مسبقاً',
        r'\bpre-?registration\b', r'\bpre-?registered\b',
        r'\bClinicalTrials\.gov\b', r'\bNCT\d{8}\b',
        r'\bPROSPERO\b', r'\bOSF\.io\b', r'\bregistration\s+number\b',
    ],
}


def scan_presence(text):
    """يعيد {بند: {'present': bool, 'hits': [شواهد قصيرة]}}."""
    out = {}
    low = text
    for key, pats in PATTERNS.items():
        hits = []
        for p in pats:
            for m in re.finditer(p, low, re.I):
                s = max(0, m.start() - 30)
                e = min(len(low), m.end() + 30)
                snippet = re.sub(r'\s+', ' ', low[s:e]).strip()
                hits.append(snippet)
                if len(hits) >= 3:
                    break
            if len(hits) >= 3:
                break
        out[key] = {'present': bool(hits), 'hits': hits[:3]}
    return out


# ════════════════════════════════════════════════════════════════════════
# القسم 3: A2 — فحوص الاتساق الداخلي (حتميّة، قابلة للتحقق منطقياً)
# ════════════════════════════════════════════════════════════════════════

AR_DIGITS = str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789')


def _years(text):
    t = text.translate(AR_DIGITS)
    return [int(y) for y in re.findall(r'\b(19|20)\d{2}\b', t) if True] \
        and [int(m.group(0)) for m in re.finditer(r'\b(?:19|20)\d{2}\b', t)]


def check_coi_vs_funding(text, presence):
    """تعارض إفصاحيّ: تصريح «لا تضارب/لا تمويل» مع وجود جهة تمويل مسمّاة."""
    flags = []
    t = text
    declares_no = bool(re.search(
        r'(?:لا\s+يوجد|لم\s+يتلقَّ|لم\s+يتلق).{0,40}(?:تمويل|دعم)'
        r'|no\s+(?:specific\s+)?funding|received\s+no\s+funding', t, re.I))
    names_funder = bool(re.search(
        r'(?:م[ُ\u0640]?وّل\s+من|بتمويل\s+من|بدعم\s+من|مُموّل\s+من'
        r'|المنحة\s+رقم|منحة\s+بحثية\s+رقم)'
        r'|funded\s+by\s+[A-Z\u0600-\u06FF]|grant\s+(?:number|no|#)\s*[:\-]?\s*\w', t, re.I))
    if declares_no and names_funder:
        flags.append({
            'type': 'coi_funding_contradiction',
            'severity': 'orange',
            'note': ('الملف يصرّح بعدم وجود تمويل/دعم، ويذكر في موضعٍ آخر جهة '
                     'تمويل أو رقم منحة. تعارضٌ إفصاحيّ يستدعي مراجعة الباحث — '
                     'قد يكون سهواً في الصياغة لا أكثر.'),
        })
    # تصريح «لا تضارب» مع تمويل من جهة قد تكون ذات مصلحة (راية تنبيه لا حكم)
    declares_no_coi = bool(re.search(
        r'لا\s+يوجد\s+تضارب|no\s+conflict|declare[sd]?\s+no', t, re.I))
    if declares_no_coi and names_funder:
        flags.append({
            'type': 'coi_declared_none_with_funding',
            'severity': 'yellow',
            'note': ('تصريحٌ بعدم تضارب المصالح مع وجود تمويل مُعلَن. هذا ليس '
                     'خللاً بذاته — التمويل لا يعني تضارباً — لكن يُستحسن أن '
                     'يذكر الباحث علاقة المموّل بموضوع البحث صراحةً.'),
        })
    return flags


def check_consent_without_ethics(text, presence, family):
    """في العائلات ذات المشاركين: وجود موافقة مستنيرة بلا ذكرٍ للجنة الأخلاقيات."""
    flags = []
    app = APPLICABILITY.get(family, APPLICABILITY['unknown'])
    if presence['informed_consent']['present'] and \
       not presence['ethics_approval']['present'] and \
       app.get('ethics_approval') != 'na':
        flags.append({
            'type': 'consent_without_ethics_body',
            'severity': 'yellow',
            'note': ('يُذكر أخذُ الموافقة المستنيرة من المشاركين، دون ذكرٍ '
                     'صريح للجنة/جهة الموافقة الأخلاقية. إن كان البحث على بشر، '
                     'يُستحسن ذكر الجهة المُجيزة. (قد تكون الجامعة لا تشترط لجنة '
                     'رسمية — تحقّق من سياق المؤسسة قبل عدّه نقصاً).'),
        })
    return flags


def check_data_on_request(text):
    """«متاحة عند الطلب» دون جهة اتصال/بريد — بيانٌ منقوص لا اتهام."""
    flags = []
    on_request = re.search(
        r'(?:متاحة\s+عند\s+الطلب|available\s+(?:up)?on\s+(?:reasonable\s+)?request)',
        text, re.I)
    if on_request:
        window = text[max(0, on_request.start() - 200):on_request.end() + 200]
        has_contact = bool(re.search(
            r'[\w.\-]+@[\w.\-]+\.\w+|المؤلف\s+المراسل|corresponding\s+author', window, re.I))
        if not has_contact:
            flags.append({
                'type': 'data_on_request_no_contact',
                'severity': 'yellow',
                'note': ('بيان «البيانات متاحة عند الطلب» دون جهة اتصال أو بريد '
                         'المؤلف المراسل قربه. يُستحسن إضافة وسيلة تواصل صريحة كي '
                         'يكون البيان قابلاً للتفعيل.'),
            })
    return flags


def check_repo_link_validity(urls):
    """صلاحية بنيوية (شكلية) لروابط المستودعات و DOI — لا تحقّق من الحلّ."""
    flags = []
    for u in urls:
        # DOI بصيغة غير صالحة بنيوياً
        m = re.search(r'doi\.org/(\S+)', u, re.I)
        if m:
            doi = m.group(1).rstrip('.,);]')
            if not re.match(r'10\.\d{4,9}/\S+', doi):
                flags.append({
                    'type': 'malformed_doi',
                    'severity': 'orange',
                    'note': f'رابط DOI ذو بنية غير صالحة شكلياً: {u} — راجع الصيغة. '
                            '(هذا خطأ نسخ محتمل، لا دليل تلفيق).',
                })
        # مستودعات بيانات معروفة: نذكرها للطبقة B (تحقّق الحلّ)
    known_repos = [u for u in urls if re.search(
        r'(?:zenodo\.org|figshare\.com|osf\.io|dryad|dataverse|github\.com|'
        r'mendeley\.com/datasets|data\.mendeley)', u, re.I)]
    return flags, known_repos


def check_ethics_date_hint(text):
    """قرينة تاريخية: رقم/تاريخ موافقة أخلاقية يُذكر — يُحال للطبقة B لمقارنته
    بنافذة جمع البيانات (لا يمكن الجزم آلياً بلا سياق، فنرفع تنبيهاً لا حكماً)."""
    flags = []
    m = re.search(r'(?:رقم\s+الموافقة|approval\s+(?:number|no|code|ref))'
                  r'\s*[:\-]?\s*([\w/\-\.]+)', text, re.I)
    if m:
        flags.append({
            'type': 'ethics_approval_ref_found',
            'severity': 'info',
            'note': (f'ذُكر مرجع موافقة أخلاقية: «{m.group(1)}». يُنصح بالتحقق '
                     'يدوياً أن تاريخ الموافقة يسبق بدء جمع البيانات (فحص الطبقة B).'),
        })
    return flags


# ════════════════════════════════════════════════════════════════════════
# القسم 4: A3 — بصمات النزاهة (رايات لا أحكام)
# ════════════════════════════════════════════════════════════════════════
# مهم: العبارات المُعذَّبة ظاهرة نصّ إنجليزيّ (نتاج أدوات إعادة صياغة تتحايل
# على كواشف الاستلال). لا تُطبَّق على العربية. ونحذّر: قد تكون بعضها ترجمةً
# حرفية بريئة من باحثٍ عربيّ لا ناطقٍ بالإنجليزية — فهي راية تحقّق لا تهمة.

TORTURED_PHRASES = {
    # المصطلح المُعذَّب : الأصل المُرجَّح (من أدبيات Cabanac وآخرين)
    'colossal information': 'big data',
    'huge information': 'big data',
    'immense information': 'big data',
    'counterfeit consciousness': 'artificial intelligence',
    'counterfeit neural network': 'artificial neural network',
    'man-made brainpower': 'artificial intelligence',
    'profound learning': 'deep learning',
    'AI (simulated intelligence)': 'AI (artificial intelligence)',
    'machine learning (ML)': None,  # سليم — للتباين فقط، لا يُرفع
    'irregular esteem': 'random value',
    'arbitrary esteem': 'random value',
    'bosom malignancy': 'breast cancer',
    'bosom disease': 'breast cancer',
    'lung disease': None,
    'kidney disappointment': 'kidney failure',
    'heart disappointment': 'heart failure',
    'liver disappointment': 'liver failure',
    'motion picture': None,
    'flag to commotion': 'signal to noise',
    'flag commotion proportion': 'signal to noise ratio',
    'mean square blunder': 'mean squared error',
    'underlying foundations': 'roots',
    'guide vector machine': 'support vector machine',
    'gullible bayes': 'naive bayes',
    'choice tree': 'decision tree',
    'irregular timberland': 'random forest',
    'irregular backwoods': 'random forest',
    'convolutional neural organization': 'convolutional neural network',
    'neural organization': 'neural network',
    'facial acknowledgment': 'facial recognition',
    'discourse acknowledgment': 'speech recognition',
    'design acknowledgment': 'pattern recognition',
    'expansive learning': 'broad learning',
    'leftover organization': 'residual network',
    'exchange learning': 'transfer learning',
    'guide component': 'feature map',
    'component extraction': None,
    'ground reality': 'ground truth',
    'ground substance': 'ground truth',
    'goal work': 'objective function',
    'misfortune work': 'loss function',
    'actuation work': 'activation function',
    'slope plummet': 'gradient descent',
    'inclination plummet': 'gradient descent',
    'overfitting issue': None,
}
# رشّح ما قيمته None (سليم/للتباين فقط)
TORTURED_ACTIVE = {k: v for k, v in TORTURED_PHRASES.items() if v}

# قرائن مطابع الورق (paper mills) — رايات ضعيفة تُجمَع ولا تُفرَد
PAPER_MILL_HINTS = [
    (r'\bas\s+per\s+our\s+knowledge\b', 'صياغة نمطية شائعة في نصوص مُولَّدة/مُصنَّعة'),
    (r'\bit\s+is\s+worth\s+to\s+mention\s+that\b', 'صياغة نمطية (خطأ نحويّ متكرر في مصانع الورق)'),
    (r'\bmentioned\s+elsewhere\b.{0,40}\bmentioned\s+elsewhere\b', 'تكرار حشوٍ نمطيّ'),
]

# قرائن التقطيع/النشر المكرّر (salami) — تُرفع للنظر البشريّ فقط
SALAMI_HINTS = [
    (r'(?:الجزء\s+(?:الأول|الثاني|الثالث)|part\s+(?:one|two|1|2)\s+of)',
     'إشارة إلى كون البحث «جزءاً» من عمل أكبر — تحقّق ألا يكون تقطيعاً لوحدة بحثية واحدة'),
    (r'(?:نُشر\s+جزء|previously\s+published|بحث\s+سابق\s+للباحث\s+نفسه)',
     'إشارة إلى نشرٍ سابق متصل — تحقّق من عدم ازدواج النشر'),
]


def scan_tortured(text):
    """يبحث العبارات المُعذَّبة في النصّ (إنجليزيّ أساساً). رايات لا أحكام."""
    hits = []
    low = text.lower()
    for bad, good in TORTURED_ACTIVE.items():
        if bad.lower() in low:
            idx = low.find(bad.lower())
            s = max(0, idx - 25)
            e = min(len(text), idx + len(bad) + 25)
            hits.append({
                'phrase': bad,
                'likely_original': good,
                'context': re.sub(r'\s+', ' ', text[s:e]).strip(),
            })
    return hits


def scan_hints(text, patterns):
    out = []
    for pat, why in patterns:
        m = re.search(pat, text, re.I)
        if m:
            out.append({'match': m.group(0), 'why': why})
    return out


# ════════════════════════════════════════════════════════════════════════
# القسم 5: التجميع والإخراج
# ════════════════════════════════════════════════════════════════════════

SEV_AR = {'red': '🔴', 'orange': '🟠', 'yellow': '🟡', 'info': 'ℹ️'}


def audit(text, urls, family):
    family = family if family in FAMILIES else 'unknown'
    app = APPLICABILITY[family]
    presence = scan_presence(text)

    # جرد الإفصاح مع حكم الانطباق
    inventory = {}
    for key in DISCLOSURES:
        applic = app.get(key, 'conditional')
        pres = presence[key]['present']
        if applic == 'na':
            status = 'not_applicable'
        elif pres:
            status = 'present'
        elif applic == 'required':
            status = 'missing_expected'
        else:  # conditional وغائب
            status = 'missing_conditional'
        inventory[key] = {
            'label': DISCLOSURE_AR[key],
            'applicability': applic,
            'present': pres,
            'status': status,
            'evidence': presence[key]['hits'],
        }

    # فحوص الاتساق
    consistency = []
    consistency += check_coi_vs_funding(text, presence)
    consistency += check_consent_without_ethics(text, presence, family)
    consistency += check_data_on_request(text)
    link_flags, known_repos = check_repo_link_validity(urls)
    consistency += link_flags
    consistency += check_ethics_date_hint(text)

    # بصمات النزاهة
    tortured = scan_tortured(text)
    mill = scan_hints(text, PAPER_MILL_HINTS)
    salami = scan_hints(text, SALAMI_HINTS)

    return {
        'family': family,
        'disclosure_inventory': inventory,
        'consistency_flags': consistency,
        'integrity_fingerprints': {
            'tortured_phrases': tortured,
            'paper_mill_hints': mill,
            'salami_hints': salami,
        },
        'repos_for_layer_b': known_repos,
        'limits': [
            'الغياب لا يثبت سوء سلوك؛ ولا يثبت غياب البند إن كان في موضع لم يُفحص.',
            'الكشف عن حضور البند لا يقيس كفايته أو صدقه — النظر البشريّ يحسم.',
            'العبارات المُعذَّبة راية تحقّق لا تهمة؛ قد تكون ترجمة حرفية بريئة.',
            'لا يحلّ هذا السكربت روابط ولا DOI؛ ذلك من عمل الطبقة B بأدوات النموذج.',
        ],
    }


def human_report(rep):
    L = []
    L.append('═' * 60)
    L.append('🛡️  تقرير جرد الأخلاقيات والنزاهة الحتميّ (disclosure_scan)')
    L.append('═' * 60)
    L.append(f'العائلة المعرفية المعتمدة للفرز: {rep["family"]}')
    L.append('')
    L.append('— جرد الإفصاح (الحضور مقيسٌ بمصفوفة العائلة) —')
    marks = {'present': '🟢 حاضر',
             'missing_expected': '🔵 غائب ويُتوقَّع',
             'missing_conditional': '⚪ غائب (مشروط/قد لا يلزم)',
             'not_applicable': '➖ غير منطبق على العائلة'}
    for key, v in rep['disclosure_inventory'].items():
        L.append(f'  {marks[v["status"]]:<28} · {v["label"]}')
    L.append('')

    cf = rep['consistency_flags']
    L.append(f'— فحوص الاتساق الداخلي: {len(cf)} تنبيه —')
    if not cf:
        L.append('  لا تعارضات إفصاحية داخلية مكتشَفة آلياً.')
    for f in cf:
        L.append(f'  {SEV_AR.get(f["severity"], "•")} [{f["type"]}]')
        L.append(f'     {f["note"]}')
    L.append('')

    fp = rep['integrity_fingerprints']
    tp, mh, sh = fp['tortured_phrases'], fp['paper_mill_hints'], fp['salami_hints']
    L.append('— بصمات النزاهة (رايات تحقّق لا أحكام) —')
    if not (tp or mh or sh):
        L.append('  لم تُرفع أي بصمة آلية. (لا يعني ذلك نفي وجود مشكلة).')
    if tp:
        L.append(f'  ⚠️ عبارات مُعذَّبة محتملة ({len(tp)}) — تحقّق: قد تكون ترجمة بريئة:')
        for h in tp[:8]:
            L.append(f'     • «{h["phrase"]}» ← الأرجح: «{h["likely_original"]}»')
    if mh:
        L.append(f'  ⚠️ قرائن صياغة نمطية ({len(mh)}):')
        for h in mh:
            L.append(f'     • «{h["match"]}» — {h["why"]}')
    if sh:
        L.append(f'  ⚠️ قرائن تقطيع/نشر مكرّر ({len(sh)}):')
        for h in sh:
            L.append(f'     • «{h["match"]}» — {h["why"]}')
    L.append('')

    if rep['repos_for_layer_b']:
        L.append('— روابط مستودعات للتحقق في الطبقة B (حلّ الرابط) —')
        for u in rep['repos_for_layer_b']:
            L.append(f'     ↗ {u}')
        L.append('')

    L.append('— حدود يجب إعلانها —')
    for lim in rep['limits']:
        L.append(f'  · {lim}')
    L.append('═' * 60)
    L.append('تذكير: هذه أدلّة يبني عليها المُحكِّم حكمه، لا أحكام نهائية.')
    return '\n'.join(L)


# ════════════════════════════════════════════════════════════════════════
# القسم 6: الفحص الذاتي
# ════════════════════════════════════════════════════════════════════════

def self_test():
    ok = True

    def check(name, cond):
        nonlocal ok
        print(f'  {"✓" if cond else "✗"} {name}')
        ok = ok and cond

    # 1) عائلة فقهية: لجنة الأخلاقيات غير منطبقة
    t1 = 'هذا بحث في فقه الضرورة والمصلحة، يعتمد المنهج الأصوليّ في الترجيح.'
    r1 = audit(t1, [], 'legal')
    check('الفقه: لجنة الأخلاقيات «غير منطبقة»',
          r1['disclosure_inventory']['ethics_approval']['status'] == 'not_applicable')

    # 2) كمّية: توافر البيانات غائب ويُتوقَّع
    t2 = 'دراسة كمّية بمقياس ليكرت، حُلّلت بـ SPSS، وأُجري اختبار t.'
    r2 = audit(t2, [], 'quantitative')
    check('الكمّية: توافر البيانات «غائب ويُتوقَّع»',
          r2['disclosure_inventory']['data_availability']['status'] == 'missing_expected')

    # 3) كشف حضور تضارب المصالح (عربي)
    t3 = 'يقرّ الباحثون بعدم وجود تضارب في المصالح. مُوّل البحث من جامعة الموصل.'
    r3 = audit(t3, [], 'quantitative')
    check('كشف حضور بند تضارب المصالح',
          r3['disclosure_inventory']['coi']['present'])

    # 4) تعارض «لا تمويل» مع ذكر جهة تمويل
    t4 = 'لم يتلقَّ البحث أيّ تمويل. وقد مُوّل من قبل وزارة التعليم العالي.'
    r4 = audit(t4, [], 'quantitative')
    check('كشف تعارض التمويل الإفصاحيّ',
          any(f['type'] == 'coi_funding_contradiction' for f in r4['consistency_flags']))

    # 5) «متاحة عند الطلب» بلا جهة اتصال
    t5 = 'The data are available upon reasonable request from the authors.'
    r5 = audit(t5, [], 'quantitative')
    check('كشف «عند الطلب» بلا جهة اتصال',
          any(f['type'] == 'data_on_request_no_contact' for f in r5['consistency_flags']))

    # 6) عبارة مُعذَّبة إنجليزية
    t6 = 'We used counterfeit consciousness and profound learning on colossal information.'
    r6 = audit(t6, [], 'quantitative')
    check('كشف ثلاث عبارات مُعذَّبة',
          len(r6['integrity_fingerprints']['tortured_phrases']) >= 3)

    # 7) DOI مشوّه بنيوياً
    r7 = audit('انظر الرابط.', ['https://doi.org/not-a-valid-doi'], 'quantitative')
    check('كشف DOI مشوّه بنيوياً',
          any(f['type'] == 'malformed_doi' for f in r7['consistency_flags']))

    # 8) مستودع بيانات يُلتقَط للطبقة B
    r8 = audit('البيانات على المستودع.', ['https://zenodo.org/record/12345'], 'quantitative')
    check('التقاط رابط مستودع للطبقة B',
          len(r8['repos_for_layer_b']) == 1)

    # 9) الفقه لا يُعاقَب على غياب توافر البيانات
    check('الفقه: توافر البيانات «غير منطبق»',
          r1['disclosure_inventory']['data_availability']['status'] == 'not_applicable')

    print('\n' + ('✅ نجحت كل الفحوص' if ok else '❌ فشل فحص واحد أو أكثر'))
    return 0 if ok else 1


# ════════════════════════════════════════════════════════════════════════

def main():
    ap = argparse.ArgumentParser(
        description='جرد الأخلاقيات والنزاهة الحتميّ (الطبقة A). دون إنترنت ولا تبعيّات.')
    ap.add_argument('file', nargs='?', help='مسار .docx أو ملف نصّي')
    ap.add_argument('--family', default='unknown',
                    help='العائلة المعرفية للفرز (quantitative/legal/...)')
    ap.add_argument('--json', metavar='OUT', help='اكتب تقريراً مهيكلاً إلى ملف')
    ap.add_argument('--self-test', action='store_true', help='فحص ذاتي على أمثلة معروفة')
    args = ap.parse_args()

    if args.self_test:
        sys.exit(self_test())
    if not args.file:
        ap.error('مرّر ملفاً، أو استعمل --self-test')

    text, urls = load(args.file)
    rep = audit(text, urls, args.family)

    if args.json:
        with open(args.json, 'w', encoding='utf-8') as f:
            json.dump(rep, f, ensure_ascii=False, indent=2)
        print(f'✓ كُتب التقرير المهيكل: {args.json}')
    print(human_report(rep))


if __name__ == '__main__':
    main()
