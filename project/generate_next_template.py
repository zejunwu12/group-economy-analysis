"""Command-line entry point for next-quarter template generation."""

import argparse
import logging
from pathlib import Path
import sys

from engine.config_loader import ConfigLoader
from engine.next_template import generate_next_template
from engine.period import QuarterContext, QuarterError
from logger import setup_logger


logger = logging.getLogger(__name__)


def _quarter_argument(value: str) -> str:
    try:
        return QuarterContext.parse(value).code
    except QuarterError as exc:
        raise argparse.ArgumentTypeError(str(exc)) from exc


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="从当前季度最终汇总表生成下季度填报模板"
    )
    parser.add_argument(
        "--quarter",
        "-q",
        required=True,
        type=_quarter_argument,
        metavar="YYYYQn",
        help="源最终汇总表所属季度，例如 2026Q2",
    )
    parser.add_argument(
        "--source",
        "-s",
        required=True,
        type=Path,
        help="当前季度经审核的最终汇总表路径",
    )
    parser.add_argument(
        "--config",
        "-c",
        type=Path,
        default=None,
        help="配置文件路径；默认使用本脚本同目录的 config.yaml",
    )
    return parser.parse_args(argv)


def _resolve_config_path(config_path: Path | None) -> Path:
    if config_path is None:
        return Path(__file__).resolve().with_name("config.yaml")
    return config_path.expanduser().resolve()


def run(
    quarter_code: str,
    source_path: str | Path,
    config_path: str | Path | None = None,
) -> dict:
    current_period = QuarterContext.parse(quarter_code)
    resolved_config = _resolve_config_path(
        Path(config_path) if config_path is not None else None
    )
    config = ConfigLoader(str(resolved_config)).load()
    runtime = config["runtime"]
    output_dir = (
        Path(runtime["_config_dir"]) / runtime["output_dir"]
    ).resolve()
    log_path = setup_logger(
        runtime.get("log_level", "INFO"),
        str(output_dir),
        file_prefix="next_template",
    )

    logger.info("源季度: %s", current_period.label)
    logger.info("目标季度: %s", current_period.next().label)
    logger.info("源最终汇总表: %s", Path(source_path).expanduser().resolve())
    logger.info("配置文件: %s", resolved_config)
    result = generate_next_template(source_path, current_period, config)
    return {
        "output_path": result.output_path,
        "log_path": log_path,
        "source_quarter": result.source_quarter,
        "target_quarter": result.target_quarter,
        "header_replacements": result.header_replacements,
        "rolled_cells": result.rolled_cells,
        "cleared_cells": result.cleared_cells,
        "comments_cleared": result.comments_cleared,
    }


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    try:
        result = run(args.quarter, args.source, config_path=args.config)
    except Exception:
        logging.getLogger(__name__).exception("下季度模板生成失败")
        return 1

    print(f"下季度模板: {result['output_path']}")
    print(f"处理日志: {result['log_path']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
