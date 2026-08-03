"""Static configuration reuse regression tests."""

from pathlib import Path
import tempfile
import unittest

import yaml

from engine.config_loader import ConfigError, ConfigLoader


def _minimal_config() -> dict:
    return {
        "ownership_files": {
            "测试权属": {"file": "测试权属.xlsx", "units": ["测试单位"]}
        },
        "reports": {
            f"report{report_id}": {"sheet_name": f"报表{report_id}"}
            for report_id in range(1, 9)
        },
        "runtime": {
            "template_path": "template.xlsx",
            "data_dir": "data",
            "output_dir": "output",
            "output_filename": "汇总{quarter_label}.xlsx",
        },
    }


class ConfigLoaderTests(unittest.TestCase):
    def _write_config(self, directory: str, config: dict) -> Path:
        path = Path(directory) / "config.yaml"
        with path.open("w", encoding="utf-8") as config_file:
            yaml.safe_dump(config, config_file, allow_unicode=True)
        return path

    def test_loads_reusable_config_without_quarter_section(self):
        with tempfile.TemporaryDirectory() as directory:
            path = self._write_config(directory, _minimal_config())
            config = ConfigLoader(str(path)).load()

        self.assertNotIn("quarter", config)
        self.assertEqual(config["unit_to_owner"], {"测试单位": "测试权属"})

    def test_rejects_legacy_quarter_section_with_migration_message(self):
        raw = _minimal_config()
        raw["quarter"] = {
            "label": "2026年第二季度",
            "end_date": "2026年6月30日",
        }

        with tempfile.TemporaryDirectory() as directory:
            path = self._write_config(directory, raw)
            with self.assertRaisesRegex(ConfigError, "--quarter YYYYQn"):
                ConfigLoader(str(path)).load()

    def test_preserves_optional_next_template_rules(self):
        raw = _minimal_config()
        raw["next_template"] = {
            "output_filename": "模板{quarter_label}.xlsx",
            "reports": {},
        }

        with tempfile.TemporaryDirectory() as directory:
            path = self._write_config(directory, raw)
            config = ConfigLoader(str(path)).load()

        self.assertEqual(config["next_template"], raw["next_template"])


if __name__ == "__main__":
    unittest.main()
