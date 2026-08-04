"""Config loader for YAML configuration"""

import os
from typing import Any

import yaml


ConfigDict = dict[str, Any]


class ConfigError(Exception):
    """配置错误"""
    pass


class ConfigLoader:
    """加载和验证 config.yaml"""

    def __init__(self, config_path: str):
        """
        Args:
            config_path: config.yaml 文件的路径
        """
        self.config_path = config_path
        self._raw: ConfigDict | None = None

    def load(self) -> ConfigDict:
        """加载并验证配置，返回结构化配置字典"""
        if not os.path.exists(self.config_path):
            raise ConfigError(f"配置文件不存在: {self.config_path}")

        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                self._raw = yaml.safe_load(f)
        except yaml.YAMLError as exc:
            raise ConfigError(f"配置文件 YAML 格式错误: {exc}") from exc

        if self._raw is None:
            raise ConfigError("配置文件为空")

        self._validate_structure()
        config = self._build_config()
        # 存储 config.yaml 所在目录，供 reader 等模块解析相对路径
        config["runtime"]["_config_dir"] = os.path.dirname(
            os.path.abspath(self.config_path)
        )
        return config

    def _require_raw(self) -> ConfigDict:
        """返回已读取的原始配置，并向类型检查器明确其非空。"""
        if not isinstance(self._raw, dict):
            raise ConfigError("配置文件根节点必须是字典")
        return self._raw

    def _validate_structure(self) -> None:
        """验证配置文件的顶层结构"""
        raw = self._require_raw()

        required_sections = ["ownership_files", "reports", "runtime"]
        for section in required_sections:
            if section not in raw:
                raise ConfigError(f"配置文件缺少必填节: '{section}'")

        if "quarter" in raw:
            raise ConfigError(
                "quarter 已改为运行参数，请删除配置文件中的 quarter 节，"
                "并使用 --quarter YYYYQn（例如 --quarter 2026Q2）"
            )

        # 验证 ownership_files
        if not isinstance(raw["ownership_files"], dict):
            raise ConfigError("ownership_files 必须是字典")
        for key, entry in raw["ownership_files"].items():
            if not isinstance(entry, dict):
                raise ConfigError(f"ownership_files.{key} 必须是字典")
            if "units" not in entry:
                raise ConfigError(f"ownership_files.{key} 缺少 'units' 字段")
            if "file" not in entry:
                raise ConfigError(f"ownership_files.{key} 缺少 'file' 字段")
            if entry["file"] is not None and not isinstance(entry["file"], str):
                raise ConfigError(f"ownership_files.{key}.file 必须是字符串或 null")
            if not isinstance(entry["units"], list) or not entry["units"]:
                raise ConfigError(f"ownership_files.{key}.units 必须是非空列表")
            if any(
                not isinstance(unit, str) or not unit.strip()
                for unit in entry["units"]
            ):
                raise ConfigError(
                    f"ownership_files.{key}.units 中的单位名称必须是非空字符串"
                )

        # 验证 reports
        if not isinstance(raw["reports"], dict):
            raise ConfigError("reports 必须是字典")
        for i in range(1, 9):
            key = f"report{i}"
            if key not in raw["reports"]:
                raise ConfigError(f"reports 缺少 '{key}'")
            report = raw["reports"][key]
            if not isinstance(report, dict):
                raise ConfigError(f"reports.{key} 必须是字典")
            if (
                not isinstance(report.get("sheet_name"), str)
                or not report["sheet_name"].strip()
            ):
                raise ConfigError(f"reports.{key}.sheet_name 必须是非空字符串")

        # 验证 runtime
        rt = raw["runtime"]
        if not isinstance(rt, dict):
            raise ConfigError("runtime 必须是字典")
        for field in ["template_path", "data_dir", "output_dir", "output_filename"]:
            if field not in rt:
                raise ConfigError(f"runtime 节缺少必填字段: '{field}'")
            if not isinstance(rt[field], str) or not rt[field].strip():
                raise ConfigError(f"runtime.{field} 必须是非空字符串")

        reports_to_run = rt.get("reports_to_run", list(range(1, 9)))
        if (
            not isinstance(reports_to_run, list)
            or any(
                isinstance(report_id, bool)
                or not isinstance(report_id, int)
                or report_id not in range(1, 9)
                for report_id in reports_to_run
            )
        ):
            raise ConfigError("runtime.reports_to_run 必须是由 1~8 组成的列表")

    def _build_config(self) -> ConfigDict:
        """从原始配置构建结构化的配置字典，补充反向索引等衍生数据"""
        raw = self._require_raw()
        cfg: ConfigDict = {}

        # 所有权属文件映射（保留原始结构）
        cfg["ownership_files"] = raw["ownership_files"]

        # 构建 unit_name -> ownership_key 的反向索引
        # 例如: "中侨集团" -> "中侨集团", "侨乡文体" -> "中侨集团"
        unit_to_owner: dict[str, str] = {}
        for owner_key, entry in raw["ownership_files"].items():
            for unit_name in entry["units"]:
                if unit_name in unit_to_owner:
                    raise ConfigError(
                        f"单位名称 '{unit_name}' 在多个权属中出现: "
                        f"'{unit_to_owner[unit_name]}' 和 '{owner_key}'"
                    )
                unit_to_owner[unit_name] = owner_key
        cfg["unit_to_owner"] = unit_to_owner

        # 报表配置
        cfg["reports"] = raw["reports"]

        # 下季度模板生成规则（可选；由独立生成流程执行进一步校验）
        if "next_template" in raw:
            cfg["next_template"] = raw["next_template"]

        # 统一字体配置（可选）
        if "font" in raw:
            self._validate_font(raw["font"])
            cfg["font"] = raw["font"]

        # 运行时配置
        cfg["runtime"] = raw["runtime"]

        return cfg

    def _validate_font(self, font_config: dict) -> None:
        """校验 font 配置节（可选）。"""
        if not isinstance(font_config, dict):
            raise ConfigError("font 必须是字典")

        if "enabled" in font_config and not isinstance(
            font_config["enabled"], bool
        ):
            raise ConfigError("font.enabled 必须是布尔值")

        if "name" in font_config:
            if not isinstance(font_config["name"], str) or not font_config[
                "name"
            ].strip():
                raise ConfigError("font.name 必须是非空字符串")

        if "size" in font_config:
            size = font_config["size"]
            if not isinstance(size, (int, float)) or isinstance(size, bool):
                raise ConfigError("font.size 必须是数字")

        if "bold" in font_config and not isinstance(
            font_config["bold"], bool
        ):
            raise ConfigError("font.bold 必须是布尔值")

        overrides = font_config.get("report_overrides")
        if overrides is not None:
            if not isinstance(overrides, dict):
                raise ConfigError("font.report_overrides 必须是字典")
            for report_id, override in overrides.items():
                if not isinstance(report_id, int) or report_id not in range(
                    1, 9
                ):
                    raise ConfigError(
                        f"font.report_overrides 的键必须是 1~8 的整数，"
                        f"实际: {report_id}"
                    )
                if not isinstance(override, dict):
                    raise ConfigError(
                        f"font.report_overrides.{report_id} 必须是字典"
                    )
                if "name" in override and (
                    not isinstance(override["name"], str)
                    or not override["name"].strip()
                ):
                    raise ConfigError(
                        f"font.report_overrides.{report_id}.name "
                        f"必须是非空字符串"
                    )
                if "size" in override:
                    override_size = override["size"]
                    if not isinstance(
                        override_size, (int, float)
                    ) or isinstance(override_size, bool):
                        raise ConfigError(
                            f"font.report_overrides.{report_id}.size "
                            f"必须是数字"
                        )
                if "bold" in override and not isinstance(
                    override.get("bold"), bool
                ):
                    raise ConfigError(
                        f"font.report_overrides.{report_id}.bold "
                        f"必须是布尔值"
                    )
