"""
カフェ アクイーユ 恵比寿 ── 予約管理表 自動生成

使い方:
    python main.py                            # input/ の最新 PDF を自動検出
    python main.py --date 20260417            # 日付指定で PDF を検索
    python main.py --pdf ./input/xxxxx.pdf    # PDF パスを直接指定
"""
import argparse
import glob
import logging
import os
import re
import sys
from datetime import datetime

from openpyxl import Workbook

from parser import parse_pdf, normalize
from builder import build_sheet_main, build_sheet_alert, build_sheet_at

# ── ログ設定 ─────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

INPUT_DIR  = os.path.join(os.path.dirname(__file__), "input")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")


# ── PDF 検索 ──────────────────────────────────────────────────────────────────

def _find_pdf(date_str: str | None = None) -> str | None:
    """
    input/ フォルダから PDF を探す。
    date_str が指定された場合はファイル名にその文字列を含むものを優先。
    複数ある場合は最終更新日時が新しいものを返す。
    """
    pattern = os.path.join(INPUT_DIR, "*.pdf")
    all_files = glob.glob(pattern)

    if not all_files:
        return None

    if date_str:
        matched = [f for f in all_files if date_str in os.path.basename(f)]
        if matched:
            return max(matched, key=os.path.getmtime)
        # 日付一致なし → None を返してエラーにする
        return None

    # 引数なし → 最新ファイル
    return max(all_files, key=os.path.getmtime)


def _extract_date_from_filename(path: str) -> str:
    """ファイル名から YYYYMMDD を抽出。取得できなければ今日の日付を使う"""
    m = re.search(r'(\d{8})', os.path.basename(path))
    if m:
        return m.group(1)
    return datetime.now().strftime("%Y%m%d")


# ── エラー出力ヘルパー ────────────────────────────────────────────────────────

def _abort(msg: str):
    """日本語エラーメッセージを表示して終了"""
    print(f"\n[エラー] {msg}\n", file=sys.stderr)
    sys.exit(1)


# ── メイン ───────────────────────────────────────────────────────────────────

def main():
    p = argparse.ArgumentParser(
        description="ebica PDF → カフェ アクイーユ 予約管理 Excel 変換",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    p.add_argument("--date", metavar="YYYYMMDD",
                   help="処理対象の日付（例: 20260417）")
    p.add_argument("--pdf",  metavar="PATH",
                   help="PDF ファイルのパスを直接指定")
    args = p.parse_args()

    os.makedirs(INPUT_DIR,  exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # ── PDF ファイルの特定 ────────────────────────────────────────
    if args.pdf:
        # --pdf で直接指定
        pdf_path = os.path.abspath(args.pdf)
        if not os.path.isfile(pdf_path):
            _abort(
                f"指定された PDF が見つかりません。\n"
                f"   パス: {pdf_path}\n"
                f"   ファイル名を確認してください。"
            )
    else:
        # 自動検索（--date あり / なし）
        pdf_path = _find_pdf(args.date)

        if pdf_path is None:
            if args.date:
                _abort(
                    f"日付 {args.date} に一致する PDF が input/ フォルダに見つかりません。\n"
                    f"   ・ファイル名に {args.date} が含まれているか確認してください。\n"
                    f"   ・ファイルが input/ フォルダに置かれているか確認してください。"
                )
            else:
                _abort(
                    "input/ フォルダに PDF ファイルが見つかりません。\n"
                    "   ebica からダウンロードした PDF を input/ フォルダに置いてください。\n"
                    "   または --pdf オプションで直接パスを指定することもできます。"
                )

    logger.info("処理対象 PDF: %s", pdf_path)

    # 日付の確定（--date > ファイル名から抽出 > 今日）
    date_str = args.date or _extract_date_from_filename(pdf_path)
    logger.info("処理日付: %s", date_str)

    # ── PDF 解析 ──────────────────────────────────────────────────
    print(f"\n[PDF解析中] {os.path.basename(pdf_path)}")
    try:
        reservations = parse_pdf(pdf_path)
    except Exception as e:
        _abort(
            f"PDF の読み込みに失敗しました。\n"
            f"   原因: {e}\n"
            f"   PDF が破損していないか、または ebica からの正規の出力か確認してください。"
        )

    if not reservations:
        print("[警告] 予約データが 0 件でした。PDF の内容を確認してください。")

    reservations = normalize(reservations)

    yyyy, mm, dd = date_str[:4], date_str[4:6], date_str[6:]
    lunch  = [r for r in reservations if r.time_start < "16:00"]
    dinner = [r for r in reservations if r.time_start >= "16:00"]
    print(f"[完了] 予約件数: 計 {len(reservations)} 件  "
          f"（ランチ {len(lunch)} 件 / ディナー {len(dinner)} 件）")
    print(f"       合計人数: {sum(r.party_size for r in reservations)} 名\n")

    # ── Excel 生成 ────────────────────────────────────────────────
    wb = Workbook()

    print("[Sheet 1] 予約管理 生成中...")
    try:
        build_sheet_main(wb, reservations, date_str)
    except Exception as e:
        _abort(f"Sheet 1 の生成に失敗しました: {e}")

    print("[Sheet 2] 会計注意リスト 生成中...")
    try:
        build_sheet_alert(wb, reservations, date_str)
    except Exception as e:
        _abort(f"Sheet 2 の生成に失敗しました: {e}")

    print("[Sheet 3] AT集計（製造用） 生成中...")
    try:
        build_sheet_at(wb, reservations, date_str)
    except Exception as e:
        _abort(f"Sheet 3 の生成に失敗しました: {e}")

    # ── 保存 ──────────────────────────────────────────────────────
    out_filename = f"アクイーユ_予約管理_{date_str}.xlsx"
    out_path = os.path.join(OUTPUT_DIR, out_filename)

    try:
        wb.save(out_path)
    except PermissionError:
        _abort(
            f"ファイルを保存できませんでした。\n"
            f"   '{out_filename}' が Excel で開かれている場合は閉じてから再実行してください。\n"
            f"   保存先: {out_path}"
        )
    except Exception as e:
        _abort(f"ファイルの保存に失敗しました: {e}")

    print(f"\n*** 完了 ***")
    print(f"   出力ファイル: {out_path}")
    print(f"   日付: {yyyy}/{mm}/{dd}")
    print()



if __name__ == "__main__":
    main()
