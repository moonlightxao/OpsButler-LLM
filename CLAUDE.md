# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

OpsButler-LLM generates deployment implementation plans (Word documents) from Excel go-live checklists. It's a Chinese-language tool for IT operations teams. The pipeline reads an Excel file, makes 3 LLM calls to analyze/map/generate content, and produces a structured Word document.

## Commands

```bash
# Install (editable)
pip install -e .

# Run the pipeline
python -m opsbutler --excel sample/上线checklist.xlsx --output output/实施方案.docx

# CLI options
#   --excel / -e     (required) Input Excel path
#   --output / -o    (required) Output Word path
#   --config / -c    (default: config.yaml) Config file path
#   --log-level      (default: INFO) DEBUG/INFO/WARNING/ERROR
```

No test suite, linter, or CI/CD is configured.

## Architecture

Linear pipeline: **Excel → LLM (3 calls) → Word**

```
main.py (CLI entry point)
  → config.py          Loads config.yaml with ${ENV_VAR} interpolation
  → excel_parser.py    Dynamic Excel parsing (auto-detects columns by candidate name lists)
  → llm_client.py      Factory: OpenAI-compatible or Ollama client, with retry + exponential backoff
  → plan_generator.py  Orchestrates 4 pipeline steps (3 LLM calls + 1 grouping step)
  → word_generator.py  Builds Word doc via python-docx
```

Source code is under `src/opsbutler/`. The package is installed as `opsbutler`.

## Pipeline Steps (in plan_generator.py)

1. **Step Mapping** (LLM call 1): Maps Excel rows to platform change steps using `mapping_rules.md` + prompt templates. Returns `StepMappingResult`.
2. **Data Grouping** (no LLM): Groups mapped rows by step → operation type. Returns `list[StepDetail]`.
3. **Summary Generation** (LLM call 2): Produces changed apps, reason/purpose, impact analysis.
4. **Risk Analysis** (LLM call 3): Produces verification plan, rollback plan, risk analysis.

All data models are Pydantic v2 classes in `models.py` (13 classes total).

## Key Design Points

- **LLM client abstraction**: `create_llm_client()` factory returns either `OpenAICompatibleClient` or `OllamaClient`. Both inherit from `LLMClient` base class.
- **LLM response parsing**: `extract_json()` has 5 fallback strategies for extracting JSON from LLM output (handles markdown fences, surrounding text, balanced-brace extraction).
- **Dynamic Excel parsing**: Column detection uses configurable candidate name lists in `config.yaml` with exact match then substring fallback. No hardcoded column names.
- **Prompt templates**: Plain text files in `prompts/` using Python `str.format()` with `{variable}` placeholders.
- **Mapping rules**: `mapping_rules.md` defines 6 platform step types. The "ROMA任务与事件" sheet maps to two separate steps (ROMA任务 and ROMA事件), filtered by which name column is non-empty.
- **Ollama think mode**: Config `llm.think` flag enables chain-of-thought reasoning for Ollama models.

## Configuration (config.yaml)

Four sections: `llm` (provider, model, API key with env var interpolation), `excel` (column name candidates, sheets to skip), `mapping` (path to rules file), `word` (output directory).
