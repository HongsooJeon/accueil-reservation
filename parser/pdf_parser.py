"""
ebica PDF → 予約データ構造化モジュール
"""
import re
import logging
from dataclasses import dataclass, field
from typing import Optional
import pdfplumber

logger = logging.getLogger(__name__)


@dataclass
class Reservation:
    time_start: str = ""          # HH:MM
    time_end: str = ""            # HH:MM
    name: str = ""                # カナ
    party_size: int = 0
    table: str = ""               # T○○
    source: str = ""              # 媒体
    plan: str = ""
    price_regular: Optional[int] = None   # 正規金額（税込）
    price_pay: Optional[int] = None       # 支払金額（税込）
    discount: Optional[int] = None        # 割引額（正の整数で保持）
    discount_label: str = ""              # 例: "OZポイント　－1,000円"
    has_allergy: bool = False
    allergy_detail: str = ""             # アレルギー内容テキスト
    anniversary_msg: str = ""
    notes: str = ""
    purpose: str = ""
    raw_lines: list = field(default_factory=list)  # デバッグ用


# ── 正規表現パターン ────────────────────────────────────────────────────
RE_TIME_START  = re.compile(r'^(\d{1,2}:\d{2})\s')
RE_TIME_END    = re.compile(r'[〜～~](\d{1,2}:\d{2})')
RE_NAME        = re.compile(r'([ァ-ヶー\s　]+)\s*様')
RE_PARTY       = re.compile(r'(\d+)\s*人')
RE_TABLE       = re.compile(r'テーブル\s*(\w+)')
RE_SOURCE      = re.compile(r'(.+?)\s*[（(]取込[）)]')
RE_PRICE_TOT   = re.compile(r'合計金額[（(]税[・.]?サ?込[）)]\s*[：:]\s*([\d,]+)\s*円')
RE_PRICE_PAY   = re.compile(r'支払[いi]?\s*金額[（(]税[・.]?サ?込[）)]\s*[：:]\s*([\d,]+)\s*円')
RE_OZ_POINT    = re.compile(r'OZポイント\s*[：:]\s*([\d,]+)')
RE_OZ_DAY      = re.compile(r'[ＯO][ＺZ]の日[^：:]*[：:]\s*(.+)')
RE_ALLERGY_FLAG = re.compile(r'●\s*アレルギー')
RE_ANNIV_MSG   = re.compile(r'回答[：:]\s*(.+)')
RE_NOTES       = re.compile(r'お客様[よりのご]*[要望リクエスト]*[：:]\s*(.+)')
RE_PURPOSE_KW  = re.compile(r'(女子会|誕生日|バースデー|記念日|デート|同窓会|歓迎会|送別会|結婚記念|家族|ファミリー)')
RE_OZ_PLAN     = re.compile(r'OZmallプラン名[：:]\s*(.+)')
RE_PLAN_LINE   = re.compile(r'プラン名[：:]\s*(.+)')

# AT集計対象キーワード
AT_KEYWORDS = ["いちごみるく", "アフタヌーンティー", "AT＋", "AT+"]
ANIVA_KEYWORDS = ["アニバーサリーランチ", "アニバランチ"]


def _parse_int(text: str) -> Optional[int]:
    try:
        return int(text.replace(",", "").strip())
    except (ValueError, AttributeError):
        return None


def _normalize_time(t: str) -> str:
    """例：'9:00' → '09:00'"""
    parts = t.split(":")
    return f"{int(parts[0]):02d}:{parts[1]}"


def _split_blocks(lines: list[str]) -> list[list[str]]:
    """時刻行を起点に1予約ブロックへ分割する"""
    blocks: list[list[str]] = []
    current: list[str] = []
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if RE_TIME_START.match(line):
            if current:
                blocks.append(current)
            current = [line]
        else:
            if current:
                current.append(line)
    if current:
        blocks.append(current)
    return blocks


def _parse_block(block: list[str]) -> Reservation:
    r = Reservation(raw_lines=block)

    full_text = "\n".join(block)

    # ── 来店時間 ──────────────────────────────
    m = RE_TIME_START.match(block[0])
    if m:
        r.time_start = _normalize_time(m.group(1))

    # ── 退席時間 ──────────────────────────────
    m = RE_TIME_END.search(full_text)
    if m:
        r.time_end = _normalize_time(m.group(1))

    # ── 予約者名（カナ） ───────────────────────
    m = RE_NAME.search(block[0])
    if m:
        r.name = m.group(1).strip()
    else:
        # 後続行も探す
        for line in block[1:4]:
            m = RE_NAME.search(line)
            if m:
                r.name = m.group(1).strip()
                break

    # ── 媒体 ──────────────────────────────────
    m = RE_SOURCE.search(block[0])
    if m:
        r.source = m.group(1).strip()
    else:
        # OZmall等は別行から
        if "OZmall" in full_text:
            r.source = "OZmall"
        elif "食べログ" in full_text:
            r.source = "食べログ"
        elif "ぐるなび" in full_text:
            r.source = "ぐるなび"
        elif "ヒトサラ" in full_text:
            r.source = "ヒトサラ"
        elif "Google" in full_text or "グーグル" in full_text:
            r.source = "Google"
        elif "ホームページ" in full_text or "HP" in block[0]:
            r.source = "HP"
        elif "直接" in full_text or "電話" in full_text:
            r.source = "直接"
        else:
            r.source = "直接"

    # ── 人数 ──────────────────────────────────
    for line in block:
        m = RE_PARTY.search(line)
        if m:
            r.party_size = int(m.group(1))
            break

    # ── テーブル ──────────────────────────────
    for line in block:
        m = RE_TABLE.search(line)
        if m:
            r.table = f"T{m.group(1).strip()}"
            break

    # ── プラン名 ──────────────────────────────
    oz_plan = RE_OZ_PLAN.search(full_text)
    if oz_plan:
        r.plan = oz_plan.group(1).strip()
    else:
        plan_line = RE_PLAN_LINE.search(full_text)
        if plan_line:
            r.plan = plan_line.group(1).strip()
        else:
            # ぐるなび/ヒトサラの [プラン名] タグ形式
            for i, line in enumerate(block):
                if '[プラン名]' in line:
                    after_tag = line[line.index('[プラン名]') + len('[プラン名]'):].strip()
                    if after_tag:
                        r.plan = after_tag
                    elif i + 1 < len(block):
                        next_line = block[i + 1].strip()
                        next_line = re.sub(r'^●[^\s]+\s+', '', next_line)
                        if next_line:
                            r.plan = next_line
                    break

    # 直接/HP予約のATプラン検出（複数列分割フォーマット）
    # PDF多段組で「ちごみるくアフタヌーン」「ティー＊option」が別行に分割される場合
    if not r.plan:
        for i, line in enumerate(block):
            if 'ちごみるくアフタヌーン' in line:
                option = None
                # 後続行でティー＊optionを検索
                for j in range(i, min(i + 4, len(block))):
                    m = re.search(r'ティー[＊*]([^\s●]{2,})', block[j])
                    if m:
                        option = m.group(1)
                        # 末尾の不要テキストを除去
                        option = re.sub(r'Web予約.*|お一人様.*|OK\d.*|\([^\)]*\)$', '', option)
                        option = re.sub(r'付[き]?$', '', option).strip()
                        break
                if option:
                    r.plan = f'いちごみるくアフタヌーンティー＊{option}'
                else:
                    r.plan = 'いちごみるくアフタヌーンティー'
                break

    # ── 金額 ──────────────────────────────────
    m = RE_PRICE_TOT.search(full_text)
    if m:
        r.price_regular = _parse_int(m.group(1))

    m = RE_PRICE_PAY.search(full_text)
    if m:
        r.price_pay = _parse_int(m.group(1))

    # 割引額の計算
    if r.price_regular and r.price_pay:
        diff = r.price_regular - r.price_pay
        if diff > 0:
            r.discount = diff
    # OZポイントの明示記載
    m = RE_OZ_POINT.search(full_text)
    if m:
        pt = _parse_int(m.group(1))
        if pt and pt > 0:
            if not r.discount:
                r.discount = pt
            r.discount_label = f"OZポイント　－{pt:,}円"

    # OZの日フラグ
    m = RE_OZ_DAY.search(full_text)
    if m and m.group(1).strip() not in ("", "なし", "0"):
        # 差分から計算済みでなければ固定500円
        if not r.discount and r.price_regular:
            r.discount = 500
        if r.discount and not r.discount_label:
            r.discount_label = f"OZの日エントリー　－{r.discount:,}円"

    # ── アレルギー ─────────────────────────────
    # ●アレルギー はebicaのフィールドラベルとして全行に入るため使用しない。
    # OZmall系は明示的なフィールド行の値で判定する。
    allergy_keys = ["アレルギー（本人）", "アレルギー（同伴者）", "苦手な食材（本人）", "苦手な食材（同伴者）"]
    details = []
    for key in allergy_keys:
        idx = full_text.find(key)
        if idx != -1:
            same_line = full_text[idx + len(key):].split("\n")[0]
            after = re.sub(r'^[：:\s\]]+', '', same_line).strip()
            if after and after not in ("なし", "ナシ", "無", ""):
                r.has_allergy = True
                details.append(after)
    if details:
        r.allergy_detail = "・".join(details)

    # ── 記念日メッセージ ───────────────────────
    m = RE_ANNIV_MSG.search(full_text)
    if m:
        msg = m.group(1).strip()
        # OZmallの「希望しない」「希望する」「なし」はメッセージなしとみなす
        if msg and not re.match(r'^希望しない|^希望する|^なし$|^無$', msg):
            r.anniversary_msg = msg

    # ── 備考・要望 ─────────────────────────────
    m = RE_NOTES.search(full_text)
    if m:
        r.notes = m.group(1).strip()

    # ── 直接予約のコース列・備考列テキスト補完 ─────────────
    # ebica PDF の多段組テキストを補完する。
    # 媒体連携（OZmall/食べログ/ぐるなびなど）は専用フィールドで処理済みのためスキップ。

    # テーブル番号後テキストのスキップ対象プレフィックス
    _SKIP_TAIL = (
        'OZmall', '食べログ', 'ぐるなび', 'ヒトサラ', 'Google',
        'Web予約', '●', '【媒体連携】', 'その他', 'HP', '直接', 'ホームページ',
    )
    # ●アレルギー後テキストの除外キーワード（フィールドラベル・予約番号など）
    _SKIP_AFTER_ALLERGY = (
        'OZmall', '食べログ', 'ぐるなび', 'ヒトサラ',
        '予約No', '席名', '[',
    )
    # プランテキストと判断するキーワード（既にplanが取れている場合に除外）
    _PLAN_KW = re.compile(r'【[^】]+】|アフタヌーンティー|平日価格|週末価格|円\s*[（(]税|コース|ランチ|ディナー')

    if not r.notes:
        # ── Pattern A: テーブル番号の後ろのテキスト（コース/備考列） ──
        m = re.search(r'テーブル\s*\w+\s+((?:(?!●).)+)$', block[0])
        if m:
            tail = m.group(1).strip()
            # SNS(Twitter/Instagram)の後ろに続くメモは取得する
            _SNS = ('Twitter', 'Instagram')
            for sns in _SNS:
                if tail.startswith(sns):
                    tail = tail[len(sns):].strip()
                    break
            else:
                # 媒体ラベル・数字のみは無効
                if any(tail.startswith(s) for s in _SKIP_TAIL) or re.match(r'^\d{5,}', tail):
                    tail = ''

            if tail:
                # 既にプランが取れており内容がプランテキストなら備考に入れない
                if r.plan and _PLAN_KW.search(tail):
                    tail = ''

            if tail:
                # 複数行にまたがる続きを収集
                course_parts = [tail]
                for line in block[1:5]:
                    if '●' in line:
                        break
                    lc = re.sub(r'^[～〜~]\d{1,2}:\d{2}\s*', '', line)
                    if '様' in lc:
                        lc = re.sub(r'^.*?様\s*', '', lc)
                    else:
                        # 様なし行: 日本語・全角が開始する位置からコース部分を抽出
                        jp_m = re.search(
                            r'[ぁ-んァ-ヶー一-龯ａ-ｚＡ-Ｚ０-９×・＊＋！－、。].+',
                            lc
                        )
                        lc = jp_m.group(0) if jp_m else re.sub(r'^\d+\s*', '', lc)
                    lc = lc.strip()
                    if lc and not any(lc.startswith(s) for s in _SKIP_TAIL) \
                            and not lc.startswith('['):
                        course_parts.append(lc)
                    else:
                        break

                # カタカナの行またぎはスペースなし結合、それ以外は全角スペース区切り
                result = course_parts[0] if course_parts else ''
                for part in course_parts[1:]:
                    if result and part:
                        if re.search(r'[ァ-ヶー]$', result) \
                                and re.search(r'^[ァ-ヶー]', part):
                            result += part
                        else:
                            result += '　' + part
                    elif part:
                        result = part
                if result and len(result) > 3:
                    r.notes = result

    if not r.notes:
        # ── Pattern B: ●アレルギー の直後に続く備考テキスト（キョウゴク様パターン） ──
        for line in block[:3]:
            m = re.search(r'●アレルギー\s+(.{5,})', line)
            if m:
                candidate = m.group(1).strip()
                # フィールドラベル・予約番号・媒体情報は除外
                if any(candidate.startswith(s) for s in _SKIP_AFTER_ALLERGY):
                    break
                if '●' in candidate or candidate.endswith(':') or candidate.endswith('：'):
                    break
                r.notes = candidate
                break

    if not r.notes:
        # ── Pattern C: ぐるなびの [要望・相談] フィールド ──
        m = re.search(r'\[要望・相談\]\s*\n(.+?)(?=\n\[|\Z)', full_text, re.DOTALL)
        if m:
            gnavi_note = m.group(1).strip()
            # 空・ダッシュ・フィールドタグ（"[備考]"など）は除外
            if gnavi_note and gnavi_note not in ('--', '') \
                    and not gnavi_note.startswith('['):
                r.notes = gnavi_note

    # ── (大人X)(子供Y) 記載を備考に追加 ────────────────────
    child_m = re.search(r'[（(]子供\s*(\d+)[）)]', full_text)
    if child_m:
        adult_m = re.search(r'[（(]大人\s*(\d+)[）)]', full_text)
        parts_note = f"子供{child_m.group(1)}名含む"
        if adult_m:
            parts_note = f"大人{adult_m.group(1)}名・子供{child_m.group(1)}名"
        r.notes = (parts_note + "　" + r.notes).strip() if r.notes else parts_note

    # ── 目的 ──────────────────────────────────
    # full_textには「記念日メッセージ」のラベルが含まれるため、限定フィールドで検索
    purpose_text = " ".join(filter(None, [r.plan, r.notes, r.anniversary_msg]))
    m = RE_PURPOSE_KW.search(purpose_text)
    if m:
        r.purpose = m.group(1)
    elif r.anniversary_msg:
        r.purpose = "記念日"

    return r


def parse_pdf(pdf_path: str) -> list[Reservation]:
    """PDFファイルを解析して予約リストを返す"""
    all_lines: list[str] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=3, y_tolerance=3)
            if text:
                all_lines.extend(text.split("\n"))

    if not all_lines:
        logger.warning("PDFからテキストを抽出できませんでした: %s", pdf_path)
        return []

    blocks = _split_blocks(all_lines)
    logger.info("抽出ブロック数: %d", len(blocks))

    reservations: list[Reservation] = []
    for i, block in enumerate(blocks):
        try:
            r = _parse_block(block)
            if r.time_start:
                reservations.append(r)
            else:
                logger.debug("時間未取得のブロックをスキップ [block %d]: %s", i, block[:2])
        except Exception as e:
            logger.error("ブロック解析エラー [block %d]: %s\n%s", i, e, "\n".join(block))

    return reservations


def is_at_plan(plan: str) -> bool:
    return any(kw in plan for kw in AT_KEYWORDS)


def is_aniva_plan(plan: str) -> bool:
    return any(kw in plan for kw in ANIVA_KEYWORDS)
