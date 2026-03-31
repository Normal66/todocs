# By Constantine Sidorov, 2026

from __future__ import annotations

import argparse
import re
from pathlib import Path

from generate_acts import build_acts, read_sheet_rows, render_all
from make_template import TemplatePaths, make_template


def _must_exist_file(p: str) -> Path:
    path = Path(p)
    if not path.exists() or not path.is_file():
        raise argparse.ArgumentTypeError(f"Файл не найден: {p}")
    return path


def _default_out_dir(repo_root: Path) -> Path:
    return repo_root / "out"


def _sanitize_name(p: Path) -> str:
    # Для имени шаблона берём только “чистое” базовое имя файла.
    s = re.sub(r"[^\w\-.]+", "_", p.stem, flags=re.UNICODE).strip("_")
    return s or "template"


def main() -> None:
    repo_root = Path(__file__).resolve().parents[1]

    ap = argparse.ArgumentParser(
        prog="todocs",
        description="Generate DOCX from Excel using Word sample (yellow highlight -> placeholders).",
    )
    ap.add_argument("--word", required=True, type=_must_exist_file, help="Исходный Word (docx) с жёлтыми полями")
    ap.add_argument("--excel", required=True, type=_must_exist_file, help="Источник данных Excel (xlsx)")
    ap.add_argument("--dop", default=str(repo_root / "DOP.JSON"), help="DOP.JSON (ACT_DATE/BASIS/MOL)")
    ap.add_argument("--out", default=str(_default_out_dir(repo_root)), help="Папка для результата")
    ap.add_argument(
        "--template-out",
        default=None,
        help="Куда сохранить сгенерированный шаблон (docx). По умолчанию: template_<word>.docx",
    )
    args = ap.parse_args()

    word: Path = args.word
    excel: Path = args.excel
    dop = Path(args.dop) if args.dop else None
    out_dir = Path(args.out)

    template_out = (
        Path(args.template_out)
        if args.template_out
        else (repo_root / f"template_{_sanitize_name(word)}.docx")
    )

    make_template(
        TemplatePaths(
            source_docx=word,
            output_docx=template_out,
            excel_xlsx=excel,
        )
    )

    rows = read_sheet_rows(excel)
    acts = build_acts(rows)
    if not acts:
        raise SystemExit("Не удалось собрать ни одного акта из Excel (проверь данные).")

    render_all(
        template_docx=template_out,
        excel_xlsx=excel,
        acts=acts,
        out_dir=out_dir,
        dop_json=dop,
    )

    print(f"OK: template={template_out}")
    print(f"OK: out_dir={out_dir} files={len(acts)}")


if __name__ == "__main__":
    main()

