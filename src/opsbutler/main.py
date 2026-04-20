import argparse
import logging
import sys
from pathlib import Path

from opsbutler.config import load_config
from opsbutler.excel_parser import load_excel, load_schedule_sheet
from opsbutler.llm_client import create_llm_client
from opsbutler.plan_generator import PlanGenerator
from opsbutler.word_generator import WordGenerator


def setup_logging(level: str = "INFO"):
    logging.basicConfig(
        level=getattr(logging, level.upper()),
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def main():
    parser = argparse.ArgumentParser(
        description="OpsButler-LLM: Generate deployment implementation plans from Excel checklists"
    )
    parser.add_argument(
        "--excel", "-e",
        required=True,
        help="Path to the Excel checklist file"
    )
    parser.add_argument(
        "--output", "-o",
        required=True,
        help="Path for the output Word document"
    )
    parser.add_argument(
        "--config", "-c",
        default="config.yaml",
        help="Path to config file (default: config.yaml)"
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Logging level (default: INFO)"
    )

    args = parser.parse_args()
    setup_logging(args.log_level)
    logger = logging.getLogger(__name__)

    # Validate input file
    excel_path = Path(args.excel)
    if not excel_path.exists():
        logger.error(f"Excel file not found: {excel_path}")
        sys.exit(1)

    # Load config
    config = load_config(args.config)
    logger.info(f"Config loaded from {args.config}")

    # Parse Excel
    logger.info(f"Parsing Excel file: {excel_path}")
    excel_payload = load_excel(str(excel_path), config)
    logger.info(f"Parsed {len(excel_payload.sheets)} sheets, {excel_payload.summary.total_rows} total rows")

    # Parse "变更安排" schedule sheet separately
    schedule_table = load_schedule_sheet(str(excel_path))
    if schedule_table:
        logger.info(f"Schedule table parsed: {len(schedule_table.rows)} rows")

    # Create LLM client
    llm_client = create_llm_client(config.llm)
    logger.info(f"LLM client created: provider={config.llm.provider}, model={config.llm.model}")

    # Generate plan
    logger.info("Generating implementation plan via LLM...")
    generator = PlanGenerator(llm_client, config)
    plan = generator.generate(excel_payload, schedule_table=schedule_table)
    logger.info(f"Plan generated: {plan.task_count} tasks, {plan.module_count} modules")

    # Generate Word document
    output_path = Path(args.output).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    word_gen = WordGenerator()
    result = word_gen.generate(plan, str(output_path))
    logger.info(f"Word document saved to: {result['output_file']}")
    if "zip_files" in result:
        for zf in result["zip_files"]:
            logger.info(f"Zip attachment saved to: {zf}")

    print(f"\nSuccessfully generated: {output_path}")
    if "zip_files" in result:
        print("Zip attachments:")
        for zf in result["zip_files"]:
            print(f"  - {zf}")


if __name__ == "__main__":
    main()
