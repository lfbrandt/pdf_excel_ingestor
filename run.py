# run.py
"""
Runner do PDF→Excel Ingestor

Funções:
- Lê PDFs de entrada (pasta, glob ou arquivos).
- Aplica o mapeamento YAML.
- Usa um modelo XLSX como base (template da planilha).
- Invoca o pdf_excel_ingestor.py com segurança.

Comportamentos inteligentes:
- Busca automaticamente o template MODELO_PLANILHA_INCLUSAO.xlsx se não for informado.
- Permite sobrescrever via variável de ambiente PDF_INGESTOR_TEMPLATE.
- Bloqueia execução caso nenhum PDF seja encontrado.
- Repassa todos os flags opcionais ao ingestor.

Uso:
    python run.py
    python run.py -i entrada/*.pdf
    python run.py -t MODELO_PLANILHA_INCLUSAO.xlsx
    python run.py --debug-argv
"""

from __future__ import annotations
import os
import sys
from pathlib import Path
import argparse
from glob import glob


# ---------------------------------------------------------
# Utilitário: detectar PDFs reais nas entradas fornecidas
# ---------------------------------------------------------
def _inputs_have_pdfs(inputs: list[str]) -> bool:
    """Verifica se há pelo menos um PDF resolvendo diretórios, globs e arquivos individuais."""
    for item in inputs:
        p = Path(item)

        # Caso seja um diretório
        if p.is_dir():
            if any(p.glob("*.pdf")):
                return True

        # Caso seja um padrão glob (ex.: *.pdf)
        elif any(x in item for x in "*?[]"):
            if glob(item):
                return True

        # Caso seja um arquivo individual
        else:
            if p.suffix.lower() == ".pdf" and p.exists():
                return True

    return False


# ---------------------------------------------------------
# Autodetecção do template XLSX
# ---------------------------------------------------------
def _auto_detect_template(root: Path) -> str | None:
    """
    Tenta encontrar automaticamente o template XLSX.
    Prioridade:
    1. PDF_INGESTOR_TEMPLATE (env)
    2. MODELO_PLANILHA_INCLUSAO.xlsx na raiz
    3. Primeiro XLSX na raiz (exceto planilhas de saída)
    """

    # 1. Variável de ambiente
    env_template = os.environ.get("PDF_INGESTOR_TEMPLATE")
    if env_template:
        p = Path(env_template)
        if not p.is_absolute():
            p = root / env_template
        if p.exists():
            return str(p)

    # 2. Template padrão esperado
    probe = root / "MODELO_PLANILHA_INCLUSAO.xlsx"
    if probe.exists():
        return str(probe)

    # 3. Qualquer XLSX na raiz (menos arquivos de saída)
    for xlsx in root.glob("*.xlsx"):
        name = xlsx.name.lower()
        if "inclusao" in name or "saida" in name:
            continue
        return str(xlsx)

    return None


# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
def main() -> int:
    root = Path(__file__).resolve().parent

    # Defaults coerentes com o projeto
    default_input = [str(root / "entrada")]
    default_mapping = str(root / "mapping.yaml")
    default_outdir = str(root / "saida")

    # Template detectado automaticamente
    default_template = _auto_detect_template(root)

    default_xlsx_name = os.environ.get("PDF_INGESTOR_XLSX_NAME", "inclusao_beneficiarios.xlsx")
    default_loglevel = os.environ.get("PDF_INGESTOR_LOG", "INFO")

    # -----------------------------------------------------
    # CLI
    # -----------------------------------------------------
    ap = argparse.ArgumentParser(description="Delegador para pdf_excel_ingestor.py")
    ap.add_argument("-i", "--input", nargs="+", default=default_input, help="Arquivos/pastas/glob de PDFs")
    ap.add_argument("-m", "--mapping", default=default_mapping, help="Arquivo mapping.yaml")
    ap.add_argument("-t", "--template", default=default_template, help="Modelo XLSX (template)")
    ap.add_argument("-o", "--outdir", default=default_outdir, help="Diretório de saída")
    ap.add_argument("--xlsx-name", default=default_xlsx_name, help="Nome do arquivo XLSX final")
    ap.add_argument("--loglevel", default=default_loglevel, help="Nível de log (DEBUG/INFO/WARNING/ERROR)")

    # repassa flags nativos do ingestor
    flags = [
        "check-template", "audit-columns", "write-markers", "write-even-if-errors",
        "dump-text", "relax-required", "fresh", "no-reports", "force-ocr"
    ]
    for f in flags:
        ap.add_argument(f"--{f}", action="store_true")

    ap.add_argument("--ocr-lang", default=os.environ.get("PDF_INGESTOR_OCR_LANG", "por+eng"))
    ap.add_argument("--debug-argv", action="store_true")

    args = ap.parse_args()

    # -----------------------------------------------------
    # Valida PDFs
    # -----------------------------------------------------
    if not _inputs_have_pdfs(args.input):
        if args.input == default_input:
            print(f"[run] Nenhum PDF encontrado em {default_input[0]}", file=sys.stderr)
        else:
            print(f"[run] Nenhum PDF encontrado nos caminhos: {', '.join(args.input)}", file=sys.stderr)
        return 2

    # Criar pastas básicas
    Path(args.outdir).mkdir(parents=True, exist_ok=True)

    # -----------------------------------------------------
    # Monta sys.argv para pdf_excel_ingestor.py
    # -----------------------------------------------------
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

    for f in flags:
        attr = f.replace("-", "_")
        if getattr(args, attr):
            argv.append(f"--{f}")

    if args.debug_argv:
        print("[run] argv ->", " ".join([f'"{a}"' if " " in a else a for a in argv]))

    # -----------------------------------------------------
    # Executa o ingestor real
    # -----------------------------------------------------
    sys.argv = argv
    sys.path.insert(0, str(root))

    import pdf_excel_ingestor
    return pdf_excel_ingestor.main()


if __name__ == "__main__":
    raise SystemExit(main())