"""Main entry point: parse Excel, generate Word docs and ledger."""

import logging
import sys
from pathlib import Path

from api_client import BureauDB
import config
from excel_parser import parse_excel, split_rows
from ledger_generator import generate_ledger
from word_generator import generate_word_docs


def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )


def main():
    setup_logging()
    logger = logging.getLogger(__name__)

    base_dir = Path(__file__).parent
    source_excel = base_dir / config.SOURCE_EXCEL
    output_dir = base_dir / config.OUTPUT_DIR

    if not source_excel.exists():
        logger.error(f"Source Excel not found: {source_excel}")
        sys.exit(1)

    # Resolve template paths
    templates = {}
    for cat, rel_path in config.TEMPLATES.items():
        tpl = base_dir / rel_path
        if not tpl.exists():
            logger.error(f"Template not found: {tpl}")
            sys.exit(1)
        templates[cat] = str(tpl)

    # Create bureau DB
    db = BureauDB()

    # Step 1: Parse Excel
    logger.info("Step 1: Parsing Excel...")
    rows = parse_excel(source_excel)

    # Stats
    from collections import Counter
    stats = Counter(r.classification for r in rows)
    logger.info(f"Classification stats: {dict(stats)}")

    # Step 2: Split rows
    logger.info("Step 2: Splitting rows...")
    items = split_rows(rows)
    logger.info(f"Total items after splitting: {len(items)}")

    # Step 2.5: Pre-check bureau coverage
    logger.info("Step 2.5: Checking bureau coverage...")
    coverage = db.check_coverage(items)
    logger.info(f"Bureau coverage: {coverage['covered']}/{coverage['total']} ({coverage['coverage_pct']:.1f}%)")

    if coverage["missing"]:
        logger.warning(f"Missing bureau for {len(coverage['missing'])} credit codes")
        template_path = output_dir / "待补充登记机关.xlsx"
        output_dir.mkdir(parents=True, exist_ok=True)
        db.export_template(coverage["missing"], str(template_path))
        logger.info(f"Template exported to: {template_path}")
        logger.info("Please fill in the template and re-run, or use 'import' command:")
        logger.info(f"  python3 main.py import <filled_template.xlsx>")
        db.close()
        sys.exit(1)

    # Step 3: Generate Word documents
    logger.info("Step 3: Generating Word documents...")
    word_results = generate_word_docs(items, templates, output_dir, db)
    success_count = sum(1 for r in word_results if r.get("filename"))
    error_count = sum(1 for r in word_results if r.get("error"))
    logger.info(f"Word generation: {success_count} success, {error_count} errors")

    # Step 4: Generate ledger Excel
    logger.info("Step 4: Generating ledger Excel...")
    ledger_path = generate_ledger(items, output_dir, db)
    logger.info(f"Ledger generated: {ledger_path}")

    db.close()

    # Summary
    print("\n" + "=" * 60)
    print("Generation Complete!")
    print(f"  Word documents: {success_count} files")
    if error_count:
        print(f"  Errors: {error_count}")
    print(f"  Ledger: {ledger_path}")
    print(f"  Output directory: {output_dir}")
    print("=" * 60)


def import_template(template_path: str):
    """Import a filled bureau template into the database."""
    setup_logging()
    logger = logging.getLogger(__name__)

    if not Path(template_path).exists():
        logger.error(f"Template file not found: {template_path}")
        sys.exit(1)

    db = BureauDB()
    result = db.import_template(template_path)
    db.close()

    print(f"Import complete: {result['imported']} entries imported, {result['skipped']} skipped")
    if result["errors"]:
        for e in result["errors"]:
            print(f"  Error: {e}")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "import":
        if len(sys.argv) < 3:
            print("Usage: python3 main.py import <filled_template.xlsx>")
            sys.exit(1)
        import_template(sys.argv[2])
    else:
        main()
