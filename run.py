# run.py
"""
Runner do PDF→Excel Ingestor
- Defaults: input=./entrada, mapping=./mapping.yaml, outdir=./saida
- Tenta achar o modelo "MODELO_PLANILHA_INCLUSAO.xlsx" na raiz
- Aceita overrides por argumentos de linha de comando
  Ex.: python run.py -i .\tests\*.pdf -t ".\MODELO_PLANILHA_INCLUSAO.xlsx"

Mudança principal:
- Não chama o ingestor se NENHUM PDF for encontrado nos caminhos informados
  (evita o erro repetido do ingestor). Retorna código 2.
"""

from __future__ import annotations
import os
import sys
from pathlib import Path
import argparse
from glob import glob


def _inputs_have_pdfs(inputs: list[str]) -> bool:
    """Verifica se há pelo menos um PDF resolvendo diretórios, globs e arquivos individuais."""
    for item in inputs:
        p = Path(item)
        if p.is_dir():
            if any(p.glob("*.pdf")):
                return True
        elif any(ch in item for ch in "*?[]"):
            if glob(item):
                return True
        else:
            if p.suffix.lower() == ".pdf" and p.exists():
                return True
    return False


def main() -> int:
    root = Path(__file__).resolve().parent

    # Defaults (coerentes com o pdf_excel_ingestor.py)
    default_input = [str(root / "entrada")]
    default_mapping = str(root / "mapping.yaml")
    default_outdir = str(root / "saida")

    # Template padrão: ENV > arquivo na raiz (se existir) > None
    env_template = os.environ.get("PDF_INGESTOR_TEMPLATE")
    probe_template = root / "MODELO_PLANILHA_INCLUSAO.xlsx"
    default_template = env_template if env_template else (str(probe_template) if probe_template.exists() else None)

    default_xlsx_name = os.environ.get("PDF_INGESTOR_XLSX_NAME", "inclusao_beneficiarios.xlsx")
    default_loglevel = os.environ.get("PDF_INGESTOR_LOG", "INFO")

    ap = argparse.ArgumentParser(description="Runner do pdf_excel_ingestor.py (delegador de argumentos)")
    ap.add_argument("-i", "--input", nargs="+", default=default_input, help="Arquivos/pastas/glob de PDFs")
    ap.add_argument("-m", "--mapping", default=default_mapping, help="YAML de mapeamento (mapping.yaml)")
    ap.add_argument("-t", "--template", default=default_template, help="Modelo XLSX (opcional; tenta achar na raiz)")
    ap.add_argument("-o", "--outdir", default=default_outdir, help="Diretório de saída")
    ap.add_argument("--xlsx-name", default=default_xlsx_name, help="Nome do XLSX final")
    ap.add_argument("--loglevel", default=default_loglevel, help="Nível de log (DEBUG/INFO/WARNING/ERROR)")

    # repassa também os flags do ingestor para conveniência
    ap.add_argument("--check-template", action="store_true", help="Só valida o template e sai")
    ap.add_argument("--audit-columns", action="store_true", help="Gera saida/column_map.csv")
    ap.add_argument("--write-markers", action="store_true", help="Escreve linha de marcadores <<key>>")
    ap.add_argument("--write-even-if-errors", action="store_true", help="Inclui linhas com erro no XLSX")
    ap.add_argument("--dump-text", action="store_true", help="Salva texto extraído por página em out/trace")
    ap.add_argument("--relax-required", action="store_true", help="Não rejeita por campos obrigatórios")
    ap.add_argument("--fresh", action="store_true", help="Recria o XLSX a partir do template (não acumula)")
    ap.add_argument("--no-reports", action="store_true", help="Não gerar relatórios")
    ap.add_argument("--force-ocr", action="store_true", help="Força OCR em todas as páginas")
    ap.add_argument(
        "--ocr-lang",
        default=os.environ.get("PDF_INGESTOR_OCR_LANG", "por+eng"),
        help="Idiomas do OCR (tesseract/ocrmypdf)",
    )
    ap.add_argument("--debug-argv", action="store_true", help="Mostra os argumentos efetivos passados ao ingestor")
    args = ap.parse_args()

    # Garante pastas básicas
    Path(args.outdir).mkdir(parents=True, exist_ok=True)
    for inp in args.input:
        p = Path(inp)
        if p.is_dir():
            p.mkdir(parents=True, exist_ok=True)

    # Se nenhum PDF for encontrado, avisa e sai com código 2 (não chama o ingestor).
    if not _inputs_have_pdfs(args.input):
        # Mensagem amigável quando usando o default
        if args.input == default_input:
            print(
                f"[run] Coloque PDFs em {default_input[0]} ou use --input \"caminho\\*.pdf\".",
                file=sys.stderr,
            )
        else:
            joined = ", ".join(args.input)
            print(f"[run] Nenhum PDF encontrado em: {joined}", file=sys.stderr)
        return 2

    # Monta argv e delega para o argparse do pdf_excel_ingestor.py
    argv = [
        "pdf_excel_ingestor.py",
        "--input", *args.input,
        "--mapping", args.mapping,
        "--outdir", args.outdir,
        "--loglevel", args.loglevel,
        "--xlsx-name", args.xlsx_name,
        "--ocr-lang", args.ocr_lang,
    ]
    if args.template:
        argv += ["--template", args.template]
    # Flags booleanos
    for flag in (
        "check_template",
        "audit_columns",
        "write_markers",
        "write_even_if_errors",
        "dump_text",
        "relax_required",
        "fresh",
        "no_reports",
        "force_ocr",
    ):
        if getattr(args, flag.replace("-", "_")):
            argv.append("--" + flag.replace("_", "-"))

    if args.debug_argv:
        print("[run] argv ->", " ".join([f'"{a}"' if " " in a else a for a in argv]), file=sys.stderr)

    # Redireciona para o main() do ingestor
    sys.argv = argv
    sys.path.insert(0, str(root))  # garante import do mesmo diretório
    import pdf_excel_ingestor  # arquivo pdf_excel_ingestor.py na raiz

    return pdf_excel_ingestor.main()


if __name__ == "__main__":
    raise SystemExit(main())