import os
import re
from dataclasses import dataclass, field
from pathlib import Path
import yaml


@dataclass
class LLMConfig:
    provider: str = "openai_compatible"
    base_url: str = "https://api.openai.com/v1"
    api_key: str = ""
    model: str = "gpt-4o"
    temperature: float = 0.3
    max_tokens: int = 4096
    retry_count: int = 2
    think: bool = False


@dataclass
class ExcelConfig:
    action_column_candidates: list[str] = field(default_factory=lambda: ["操作类型", "操作", "变更事项", "变更类型"])
    app_column_candidates: list[str] = field(default_factory=lambda: ["APPID", "应用", "应用名称"])
    skip_sheets: list[str] = field(default_factory=lambda: ["变更安排"])


@dataclass
class MappingConfig:
    rules_file: str = "./mapping_rules.md"


@dataclass
class WordConfig:
    output_dir: str = "./output"


@dataclass
class Config:
    llm: LLMConfig = field(default_factory=LLMConfig)
    excel: ExcelConfig = field(default_factory=ExcelConfig)
    mapping: MappingConfig = field(default_factory=MappingConfig)
    word: WordConfig = field(default_factory=WordConfig)


def _interpolate_env_vars(value: str) -> str:
    """Replace ${VAR_NAME} patterns with environment variable values."""
    def replacer(match):
        var_name = match.group(1)
        return os.environ.get(var_name, match.group(0))
    return re.sub(r'\$\{(\w+)\}', replacer, value)


def _interpolate_dict(d: dict) -> dict:
    """Recursively interpolate environment variables in dict values."""
    result = {}
    for key, value in d.items():
        if isinstance(value, str):
            result[key] = _interpolate_env_vars(value)
        elif isinstance(value, dict):
            result[key] = _interpolate_dict(value)
        elif isinstance(value, list):
            result[key] = [_interpolate_env_vars(v) if isinstance(v, str) else v for v in value]
        else:
            result[key] = value
    return result


def _dict_to_config(data: dict) -> Config:
    """Convert raw dict to Config dataclass."""
    llm_data = data.get("llm", {})
    excel_data = data.get("excel", {})
    mapping_data = data.get("mapping", {})
    word_data = data.get("word", {})

    return Config(
        llm=LLMConfig(
            provider=llm_data.get("provider", "openai_compatible"),
            base_url=llm_data.get("base_url", "https://api.openai.com/v1"),
            api_key=llm_data.get("api_key", ""),
            model=llm_data.get("model", "gpt-4o"),
            temperature=float(llm_data.get("temperature", 0.3)),
            max_tokens=int(llm_data.get("max_tokens", 4096)),
            retry_count=int(llm_data.get("retry_count", 2)),
            think=bool(llm_data.get("think", False)),
        ),
        excel=ExcelConfig(
            action_column_candidates=excel_data.get("action_column_candidates", ["操作类型", "操作", "变更事项", "变更类型"]),
            app_column_candidates=excel_data.get("app_column_candidates", ["APPID", "应用", "应用名称"]),
            skip_sheets=excel_data.get("skip_sheets", ["变更安排"]),
        ),
        mapping=MappingConfig(
            rules_file=mapping_data.get("rules_file", "./mapping_rules.md"),
        ),
        word=WordConfig(
            output_dir=word_data.get("output_dir", "./output"),
        ),
    )


def load_config(path: str = "config.yaml") -> Config:
    """Load and parse YAML config file with environment variable interpolation."""
    config_path = Path(path)
    if not config_path.exists():
        return Config()

    with open(config_path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    interpolated = _interpolate_dict(raw)
    return _dict_to_config(interpolated)
