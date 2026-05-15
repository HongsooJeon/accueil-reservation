"""
データクレンジング・正規化モジュール
"""
import re
from .pdf_parser import Reservation

# ── 完全一致の短縮 ─────────────────────────────────────────
_PLAN_EXACT = {
    "お席のみのご予約": "席のみ",
    "席のみのご予約":   "席のみ",
    "席のみ":          "席のみ",
}

# 【】内がこれらを含む場合は「修飾子ブラケット」として読み飛ばす
_BRACKET_QUALIFIERS = {"ＯＺ限定", "平日価格", "週末価格", "土日価格", "期間限定", "ＯＺ"}
# 上記に加え、「価格」または「限定」を含むブラケットも修飾子とみなす（例: 土日祝日価格, 平日ディナー限定価格）
_BRACKET_QUALIFIER_RE = re.compile(r'価格|限定')

# メインプラン名キーワード → 代表ラベル
_MAIN_KW_MAP = [
    ("チェリーアフタヌーンティー",     "AT"),
    ("いちごみるくアフタヌーンティー", "AT"),
    ("アフタヌーンティー",             "AT"),
    ("チェリーハイティー",             "チェリーHT"),
    ("いちごみるくハイティー",         "ハイティー"),
    ("ハイティー",                     "ハイティー"),
    ("アニバーサリーランチ",           "アニバーサリーランチ"),
    ("アニバランチ",                   "アニバーサリーランチ"),
    ("夜カフェプラン",                 "夜カフェプラン"),
    ("スペシャルプラン",               "女子会プラン"),
    ("アクイーユセットBプラン",        "アクイーユセットBプラン"),
    ("アクイーユセットAプラン",        "アクイーユセットAプラン"),
    ("お誕生日・記念日をお祝い",       "バースデーコース"),
    ("最大３ｈ飲み放題",              "女子会プラン＋３H"),
    ("最大3h飲み放題",               "女子会プラン＋３H"),
    ("最大3H飲み放題",               "女子会プラン＋３H"),
]

# 媒体名の正規化テーブル
SOURCE_ALIASES = {
    "OZmall": "OZmall",
    "Ozmall": "OZmall",
    "食べログ": "食べログ",
    "ぐるなび": "ぐるなび",
    "ヒトサラ": "ヒトサラ",
    "Google": "Google",
    "ホットペッパー": "ホットペッパー",
    "HP": "HP",
    "直接": "直接",
    "電話": "直接",
}

# ぐるなびのダミー名パターン
DUMMY_NAME_RE = re.compile(r'(グルナビ|グーグル|ヒトサラ|ゴルフ)\s*(グーグル|ヨヤク|グルナビ)?')


# ── オプションテキスト整形 ──────────────────────────────────

def _clean_opt(opt: str) -> str:
    """個別オプション文字列を短縮・正規化する"""
    opt = opt.strip()
    opt = re.sub(r'^選べる', '', opt)
    opt = re.sub(r'(\d+)時間', r'\1h', opt)
    # ○hカフェフリー → ○H飲み放題（全角数字・大文字H）
    _ZEN = str.maketrans('0123456789', '０１２３４５６７８９')
    opt = re.sub(
        r'([0-9０-９]+)[hｈH]カフェフリー',
        lambda m: m.group(1).translate(_ZEN) + 'H飲み放題',
        opt,
    )
    # 大文字H → h に統一（半角数字のみ、全角は変換しない）
    opt = re.sub(r'([0-9]+)H(?=飲み放題|フリー)', lambda m: f'{m.group(1)}h', opt)
    # 2ｈ飲み放題 → 飲み放題2h（順序入れ替え）
    opt = re.sub(r'^([0-9０-９]+)[hｈ]飲み放題', lambda m: f'飲み放題{m.group(1)}h', opt)
    # ソフトドリンク2H飲み放題 → ソフトDR飲み放題2h
    opt = re.sub(r'ソフトドリンク([0-9]+)h飲み放題', r'ソフトDR飲み放題\1h', opt)
    opt = re.sub(r'滞在時間無制限', '滞在無制限', opt)
    opt = re.sub(r'滞在カフェフリー', 'カフェフリー', opt)
    opt = re.sub(r'滞在フリータイム', '滞在フリー', opt)
    opt = re.sub(r'ノンアル(?:コール)?(?:\d*)?ドリンク', 'ノンアル', opt)
    opt = re.sub(r'ワンドリンク', 'カフェ1杯', opt)
    opt = re.sub(r'食後の', '', opt)
    opt = re.sub(r'メッセージプレート', 'プレート', opt)
    opt = re.sub(r'プレートでお祝い?', 'プレート', opt)
    opt = re.sub(r'付き$', '', opt).strip()
    # (平日、3/15～4/30) → (平日)  括弧内の日付範囲を除去
    opt = re.sub(r'\(([^）)]+?)[、,]\s*\d{1,2}/\d{1,2}[～~][^)）]{0,15}\)', r'(\1)', opt)
    # (平日限定) / (女性限定) / (土日祝...) などの末尾修飾子を除去
    opt = re.sub(r'\((?:平日|週末|土日|期間|女性|男性|限定)[^)]*\)$', '', opt).strip()
    # PDF截断による不完全な文字列を補完
    opt = re.sub(r'^メッ$', 'プレート', opt)           # メッセージプレート截断
    opt = re.sub(r'^カフェフ$', 'カフェフリー', opt)   # カフェフリー截断
    return opt.strip()


def _extract_at_opts(plan: str) -> str:
    """AT系プランの +オプション部分を返す（例: '＋カフェ1杯(平日)'）"""
    # OZmall スタイル: 【いちごみるくAT / チェリーAT】+opts
    m = re.search(r'【[^】]*アフタヌーンティー[^】]*】(.+)', plan)
    if m:
        raw = m.group(1).lstrip('+＋＆&')
        parts = re.split(r'[+＋＆&]', raw)
        opts = []
        for p in parts:
            c = _clean_opt(p)
            if c and len(c) <= 18 and '円' not in c and '名' not in c and 'OK' not in c:
                opts.append(c)
            if len(opts) >= 3:
                break
        return '＋' + '+'.join(opts) if opts else ''

    # 直接・HP / 食べログ スタイル: (いちごみるく|チェリー)AT＊option
    m = re.search(r'(?:いちごみるく|チェリー)?アフタヌーンティー[^＊*\n]*[＊*]([^ 　（(\n]+)', plan)
    if m:
        c = _clean_opt(m.group(1))
        if c and len(c) <= 15:
            return '＋' + c
    return ''


def _extract_ht_opts(plan: str) -> str:
    """ハイティー系プランのオプション部分を返す"""
    # 食べログ/直接生テキスト: ハイティー＊opts（＊区切り）
    # pdf_parser組み立て済み: ハイティー＋opts（＋区切り）のどちらも処理
    m = re.search(r'ハイティー[^＊*＋+\n]*[＊*＋+](.+)', plan)
    if not m:
        return ''
    raw = m.group(1)
    opts = []
    if 'SpecialTea' in raw or 'Special Tea' in raw:
        opts.append('SpecialTea')
    if 'スパークリング' in raw:
        opts.append('乾杯スパークリング')
    elif '乾杯' in raw:
        opts.append('乾杯')
    if re.search(r'[２2][hｈ]|2時間', raw):
        if 'フリーフロー' in raw:
            opts.append('フリーフロー2h')
        elif 'フリー' in raw:
            opts.append('フリー2h')
    return '＋' + '+'.join(opts) if opts else ''


# ── メインプラン名フォーマット ──────────────────────────────

def _format_plan(plan: str) -> str:
    if not plan:
        return ""

    # ゴミ検出：第1行の時刻+名前が混入しているパターン
    if re.match(r'^\d{1,2}:\d{2}\s', plan) or '様' in plan[:15]:
        return ""

    # 完全一致の短縮
    if plan in _PLAN_EXACT:
        return _PLAN_EXACT[plan]

    # 席のみ検出（ぐるなびなどのプラン説明文に含まれる場合）
    if 'お席のみ' in plan or re.search(r'席のみ(?:のご予約)?', plan):
        return '席のみ'

    # 先頭ノイズ除去：日付範囲 / NEW〇…〇 / NEW
    plan = re.sub(r'^\s*\d{1,2}/\d{1,2}[～~]\s*', '', plan)
    plan = re.sub(r'^(NEW〇[^〇]*〇|NEW)\s*', '', plan)

    # ─ ハイティー系プランを優先処理 ─
    if 'ハイティー' in plan:
        label = 'チェリーHT' if 'チェリー' in plan else 'ハイティー'
        return label + _extract_ht_opts(plan)

    # ─ AT系プランを優先処理 ─
    if 'チェリーアフタヌーンティー' in plan:
        return 'AT' + _extract_at_opts(plan)
    if 'いちごみるくアフタヌーンティー' in plan or 'アフタヌーンティー' in plan:
        return 'AT' + _extract_at_opts(plan)

    # ─ 食べログ 『』 形式のプラン名 ─
    m_kagi = re.search(r'[『｢]([^』｣]+)[』｣]', plan)
    if m_kagi:
        inner = m_kagi.group(1)
        for kw, label in _MAIN_KW_MAP:
            if kw in inner:
                return label  # ラベル自体にオプション込みのため追加抽出不要
        return inner[:22].strip()

    # ─ スペシャルプラン → 女子会プラン＋○H ─
    if 'スペシャルプラン' in plan:
        m_h = re.search(r'([0-9０-９]+)\s*(?:時間|[hｈH])', plan)
        if m_h:
            h = m_h.group(1).translate(str.maketrans('0123456789', '０１２３４５６７８９'))
            return f'女子会プラン＋{h}H'
        return '女子会プラン'

    # ─ 【】ブラケットを解析 ─
    main_label = None
    after_bracket = ""
    last_end = 0

    for m in re.finditer(r'【([^】]+)】', plan):
        content = m.group(1)
        last_end = m.end()
        if any(q in content for q in _BRACKET_QUALIFIERS) or _BRACKET_QUALIFIER_RE.search(content):
            continue
        # ★以降を除去してメイン名を抽出
        content_clean = re.sub(r'★.*$', '', content).strip()
        for kw, label in _MAIN_KW_MAP:
            if kw in content_clean:
                main_label = label
                after_bracket = plan[m.end():].strip()
                break
        if main_label is None:
            main_label = content_clean[:20]
            after_bracket = plan[m.end():].strip()
        break

    # ─ 修飾子ブラケットのみ or ブラケットなし → テキスト検索 ─
    if main_label is None:
        remainder = plan[last_end:].strip() if last_end else plan
        for kw, label in _MAIN_KW_MAP:
            if kw in remainder:
                main_label = label
                idx = remainder.find(kw)
                after_bracket = remainder[idx + len(kw):]
                break
        if main_label is None:
            return plan[:22].strip()

    # ─ バースデーコース（食べログ）→ 松セット ─
    if main_label == "バースデーコース":
        return "松セット"

    # ─ 夜カフェプラン → OZ松 / OZ竹 ─
    if main_label == "夜カフェプラン":
        if '4皿' in plan:
            return 'OZ松'
        elif '3皿' in plan:
            return 'OZ竹'
        return 'OZ松'  # 皿数不明の場合はデフォルト

    # ─ オプション抽出 ─
    if not after_bracket:
        return main_label

    first_plus = re.search(r'[+＋]', after_bracket)
    if not first_plus:
        return main_label

    opts_raw = after_bracket[first_plus.start():].lstrip('+＋')
    # まず + で分割し、さらに各パートを ＆ で分割
    primary_parts = re.split(r'[+＋]', opts_raw)
    parts = []
    for pp in primary_parts:
        parts.extend(re.split(r'[＆&]', pp))
    opts = []
    for p in parts:
        c = _clean_opt(p.strip())
        if c and len(c) <= 15 and '円' not in c and '名' not in c and 'OK' not in c:
            opts.append(c)
        if len(opts) >= 3:
            break

    if opts:
        return main_label + '+' + '+'.join(opts)
    return main_label


# ── その他の正規化関数 ─────────────────────────────────────

def _normalize_source(src: str) -> str:
    for key, val in SOURCE_ALIASES.items():
        if key in src:
            return val
    return src


def _normalize_name(name: str) -> str:
    return name.strip().replace("　", " ")


def _infer_purpose(r: Reservation) -> str:
    if r.purpose:
        return r.purpose
    text = r.plan + r.notes + r.anniversary_msg
    mapping = {
        "誕生日": "誕生日", "バースデー": "誕生日", "記念日": "記念日",
        "女子会": "女子会", "デート": "デート", "同窓会": "同窓会",
        "ファミリー": "家族", "家族": "家族",
        "歓迎会": "歓迎会", "送別会": "送別会",
    }
    for kw, val in mapping.items():
        if kw in text:
            return val
    return ""


def normalize(reservations: list[Reservation]) -> list[Reservation]:
    for r in reservations:
        r.name    = _normalize_name(r.name)
        r.source  = _normalize_source(r.source)
        r.plan    = _format_plan(r.plan)
        r.purpose = _infer_purpose(r)
        if r.table.startswith("T") and r.table[1:].isdigit():
            r.table = f"T{int(r.table[1:]):02d}"
        # Google予約: ダミー名（グーグル ヨヤク）を備考の実名で置き換える
        if r.source == "Google" and DUMMY_NAME_RE.search(r.name) and r.notes:
            r.name  = r.notes.strip()
            r.notes = ""
    reservations.sort(key=lambda x: x.time_start)
    return reservations
