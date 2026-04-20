from __future__ import annotations

import copy
import json
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable, Iterable, Mapping, Sequence


STANDARD_PATTERN = re.compile(
    r"\b(?:CCSS(?:\.MATH(?:\.CONTENT)?)?\.)?(?:[A-Z0-9]{1,6}\.){2,6}[A-Z0-9]{1,6}\b"
)
READY_TRIGGER_PATTERN = re.compile(r"^\s*(ready|lp:?)(?:\s+|:\s*)(.*)$", re.IGNORECASE)
SESSION_REQUEST_PATTERN = re.compile(r"\b(?:session|day)\s*([12])\b", re.IGNORECASE)
ISO_DATE_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")
DOT_SLASH_DATE_PATTERN = re.compile(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$")
DECK_DATE_PATTERN = re.compile(r"(20\d{2})[-_.](\d{1,2})[-_.](\d{1,2})")


class LessonPlanError(RuntimeError):
    """Raised when the local lesson-plan pipeline cannot complete safely."""


def ensure_directory(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def normalize_whitespace(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "")).strip()


def clean_source_line(value: str) -> str:
    cleaned = normalize_whitespace(value.replace("\u200b", " ").replace("\ufeff", " "))
    cleaned = re.sub(r"\s+([,.;:?!])", r"\1", cleaned)
    return normalize_whitespace(cleaned)


def clean_line(value: str) -> str:
    cleaned = clean_source_line(value)
    cleaned = re.sub(r"\s*Reveal:\s*.*$", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+(Mindset|Workspace|Analyzing)\s*$", "", cleaned, flags=re.IGNORECASE)
    return normalize_whitespace(cleaned)


def unique_preserve(
    values: Iterable[str],
    *,
    cleaner: Callable[[str], str] | None = None,
) -> list[str]:
    cleaner = cleaner or clean_line
    seen: set[str] = set()
    results: list[str] = []
    for value in values:
        cleaned = cleaner(value)
        if not cleaned:
            continue
        lowered = cleaned.lower()
        if lowered in seen:
            continue
        seen.add(lowered)
        results.append(cleaned)
    return results


def truncate(value: str, limit: int) -> str:
    cleaned = clean_line(value)
    if len(cleaned) <= limit:
        return cleaned
    return cleaned[: limit - 1].rstrip() + "…"


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:  # pragma: no cover - defensive
        raise LessonPlanError(f"Missing JSON file: {path}") from exc
    except json.JSONDecodeError as exc:
        raise LessonPlanError(f"Malformed JSON file: {path}") from exc


def write_json(path: Path, payload: Any) -> None:
    ensure_directory(path.parent)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=True) + "\n", encoding="utf-8")


def write_text(path: Path, text: str) -> None:
    ensure_directory(path.parent)
    path.write_text(text, encoding="utf-8")


def sanitize_artifact_path(path: Path, base_dir: Path | None = None) -> str:
    resolved = path.resolve()
    if base_dir is not None:
        try:
            return str(resolved.relative_to(base_dir.resolve()))
        except ValueError:
            pass
    return resolved.name


def deep_merge(base: Mapping[str, Any], override: Mapping[str, Any]) -> dict[str, Any]:
    result = copy.deepcopy(dict(base))
    for key, value in override.items():
        if isinstance(value, Mapping) and isinstance(result.get(key), Mapping):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result


def load_effective_config(config_path: Path, school_defaults_path: Path) -> dict[str, Any]:
    school_defaults = read_json(school_defaults_path)
    generator_config = read_json(config_path)
    merged = deep_merge(school_defaults, generator_config)
    merged.setdefault("lesson_duration_minutes", 55)
    merged.setdefault("active_student_supports", [])
    merged.setdefault("enable_esol_supports", True)
    merged.setdefault("esol_support_limit", 3)
    merged.setdefault("timing", {})
    merged.setdefault("phase_timing", {})
    merged.setdefault("input", {})
    merged.setdefault("extracted", {})
    merged.setdefault("output", {})
    merged.setdefault("templates", {})
    return merged


def select_slide_deck(
    slides_dir: Path,
    explicit_deck: str = "",
    base_dir: Path | None = None,
) -> Path:
    base_dir = base_dir or slides_dir.parent
    if explicit_deck:
        deck_path = Path(explicit_deck)
        if not deck_path.is_absolute():
            deck_path = (base_dir / explicit_deck).resolve()
        if not deck_path.exists():
            raise LessonPlanError(f"Requested slide deck was not found: {deck_path}")
        if deck_path.suffix.lower() != ".pptx":
            raise LessonPlanError(f"Requested slide deck must be a .pptx file: {deck_path}")
        return deck_path

    decks = sorted(
        path
        for path in slides_dir.glob("*.pptx")
        if path.is_file() and not path.name.startswith("~$")
    )
    if not decks:
        raise LessonPlanError(
            f"No .pptx files were found in {slides_dir}. Place one teacher slide deck in that folder or use --deck."
        )
    ranked = sorted(
        decks,
        key=lambda path: (
            deck_relevance_score(path),
            extract_date_from_filename(path.name) or "",
            path.stat().st_mtime,
            path.name.lower(),
        ),
        reverse=True,
    )
    return ranked[0]


def read_optional_agenda(agenda_dir: Path) -> list[str]:
    if not agenda_dir.exists():
        return []
    agenda_files = sorted(path for path in agenda_dir.iterdir() if path.is_file())
    if not agenda_files:
        return []

    agenda_lines: list[str] = []
    for path in agenda_files:
        suffix = path.suffix.lower()
        if suffix in {".txt", ".md"}:
            agenda_lines.extend(path.read_text(encoding="utf-8").splitlines())
        elif suffix == ".json":
            payload = read_json(path)
            if isinstance(payload, dict):
                for value in payload.values():
                    if isinstance(value, list):
                        agenda_lines.extend(str(item) for item in value)
                    else:
                        agenda_lines.append(str(value))
            elif isinstance(payload, list):
                agenda_lines.extend(str(item) for item in payload)
    return unique_preserve(agenda_lines)


def parse_standards(text: str) -> list[str]:
    return unique_preserve(STANDARD_PATTERN.findall(text))


def infer_date(date_override: str = "") -> str:
    override = clean_line(date_override)
    return override or date.today().isoformat()


def parse_trigger_text(raw_text: str) -> dict[str, Any]:
    normalized = normalize_whitespace(raw_text)
    if not normalized:
        return {
            "raw": "",
            "triggered": False,
            "command": "",
            "date": "",
            "requested_session_numbers": [],
        }

    match = READY_TRIGGER_PATTERN.match(normalized)
    if match:
        command = match.group(1).lower().rstrip(":")
        remainder = normalize_whitespace(match.group(2))
    else:
        command = ""
        remainder = normalized

    requested_sessions = [int(value) for value in SESSION_REQUEST_PATTERN.findall(remainder)]
    requested_sessions = sorted(set(number for number in requested_sessions if number in (1, 2)))
    remainder = SESSION_REQUEST_PATTERN.sub("", remainder)
    parsed_date = parse_date_token(remainder)

    if normalized and not parsed_date and command:
        raise LessonPlanError(
            f"Could not parse the requested lesson-plan date from trigger text: '{normalized}'."
        )

    return {
        "raw": normalized,
        "triggered": bool(command),
        "command": command,
        "date": parsed_date,
        "requested_session_numbers": requested_sessions,
    }


def parse_date_token(raw_value: str) -> str:
    candidate = clean_source_line(raw_value).replace(" ", "")
    if not candidate:
        return ""
    if ISO_DATE_PATTERN.match(candidate):
        try:
            return datetime.strptime(candidate, "%Y-%m-%d").date().isoformat()
        except ValueError as exc:
            raise LessonPlanError(f"Invalid ISO date in trigger text: {raw_value}") from exc

    match = DOT_SLASH_DATE_PATTERN.match(candidate)
    if not match:
        return ""

    month = int(match.group(1))
    day_value = int(match.group(2))
    year = int(match.group(3))
    try:
        return date(year, month, day_value).isoformat()
    except ValueError as exc:
        raise LessonPlanError(f"Invalid date in trigger text: {raw_value}") from exc


def resolve_run_date(trigger_context: Mapping[str, Any], config: Mapping[str, Any]) -> str:
    trigger_date = clean_line(str(trigger_context.get("date", "")))
    if trigger_date:
        return trigger_date
    return infer_date(str(config.get("date_override", "")))


def extract_date_from_filename(filename: str) -> str:
    match = DECK_DATE_PATTERN.search(filename)
    if not match:
        return ""
    year = int(match.group(1))
    month = int(match.group(2))
    day_value = int(match.group(3))
    try:
        return date(year, month, day_value).isoformat()
    except ValueError:
        return ""


def deck_relevance_score(path: Path) -> int:
    name = path.name.lower()
    score = 0
    if "editable lesson presentation" in name:
        score += 10
    if "lesson" in name:
        score += 6
    if "session" in name:
        score += 4
    if "teacher" in name:
        score += 2
    if "review" in name:
        score += 1
    if "notebook" in name:
        score -= 8
    if "lesson plan" in name:
        score -= 6
    if "template" in name:
        score -= 3
    return score


def sentence_case_label(label: str) -> str:
    return clean_line(label.replace("_", " ")).title()


def join_slide_numbers(numbers: Sequence[int]) -> str:
    return ", ".join(str(number) for number in numbers)


def format_session_docx_name(run_date: str, session_label: str) -> str:
    label = clean_line(session_label).replace(" ", "_")
    return f"{run_date}_Lesson_Plan_{label}.docx"


def sentence_list(values: Iterable[str]) -> list[str]:
    return [clean_line(value) for value in values if clean_line(value)]


def keyword_in_text(text: str, keywords: Sequence[str]) -> bool:
    lowered = text.lower()
    return any(keyword in lowered for keyword in keywords)


def get_slide_lookup(raw_deck: Mapping[str, Any]) -> dict[int, dict[str, Any]]:
    return {int(slide["slide_number"]): dict(slide) for slide in raw_deck.get("slides", [])}


def simple_schema_validate(
    instance: Any,
    schema: Mapping[str, Any],
    path: str = "$",
    *,
    root_schema: Mapping[str, Any] | None = None,
) -> list[str]:
    errors: list[str] = []
    root_schema = root_schema or schema
    schema_type = schema.get("type")
    if schema_type and not _matches_type(instance, schema_type):
        return [f"{path}: expected {schema_type}, found {type(instance).__name__}"]

    if schema_type == "object":
        if not isinstance(instance, dict):
            return [f"{path}: expected object"]
        required = schema.get("required", [])
        for key in required:
            if key not in instance:
                errors.append(f"{path}: missing required key '{key}'")
        properties = schema.get("properties", {})
        additional_allowed = schema.get("additionalProperties", True)
        if additional_allowed is False:
            for key in instance.keys():
                if key not in properties:
                    errors.append(f"{path}: unexpected key '{key}'")
        for key, subschema in properties.items():
            if key in instance:
                errors.extend(
                    simple_schema_validate(
                        instance[key],
                        _resolve_ref(root_schema, subschema),
                        f"{path}.{key}",
                        root_schema=root_schema,
                    )
                )
    elif schema_type == "array":
        if not isinstance(instance, list):
            return [f"{path}: expected array"]
        min_items = schema.get("minItems")
        if min_items is not None and len(instance) < min_items:
            errors.append(f"{path}: expected at least {min_items} items")
        item_schema = schema.get("items")
        if item_schema:
            resolved = _resolve_ref(root_schema, item_schema)
            for index, item in enumerate(instance):
                errors.extend(simple_schema_validate(item, resolved, f"{path}[{index}]", root_schema=root_schema))
    elif schema_type == "string":
        min_length = schema.get("minLength")
        if min_length is not None and len(str(instance)) < min_length:
            errors.append(f"{path}: expected string length >= {min_length}")
        enum = schema.get("enum")
        if enum is not None and instance not in enum:
            errors.append(f"{path}: expected one of {enum}")
    elif schema_type == "integer":
        minimum = schema.get("minimum")
        if minimum is not None and int(instance) < minimum:
            errors.append(f"{path}: expected integer >= {minimum}")
    elif schema_type == "boolean":
        pass

    enum = schema.get("enum")
    if enum is not None and instance not in enum:
        errors.append(f"{path}: expected one of {enum}")
    return errors


def validate_against_schema(instance: Any, schema_path: Path) -> list[str]:
    schema = read_json(schema_path)
    try:
        from jsonschema import Draft202012Validator  # type: ignore
    except ImportError:
        return simple_schema_validate(instance, schema, root_schema=schema)

    validator = Draft202012Validator(schema)
    return [f"{'.'.join(str(part) for part in error.absolute_path) or '$'}: {error.message}" for error in validator.iter_errors(instance)]


def _matches_type(instance: Any, schema_type: str) -> bool:
    if schema_type == "object":
        return isinstance(instance, dict)
    if schema_type == "array":
        return isinstance(instance, list)
    if schema_type == "string":
        return isinstance(instance, str)
    if schema_type == "integer":
        return isinstance(instance, int) and not isinstance(instance, bool)
    if schema_type == "boolean":
        return isinstance(instance, bool)
    return True


def _resolve_ref(root_schema: Mapping[str, Any], schema: Mapping[str, Any]) -> Mapping[str, Any]:
    ref = schema.get("$ref")
    if not ref:
        return schema
    if not ref.startswith("#/$defs/"):
        raise LessonPlanError(f"Unsupported schema ref: {ref}")
    ref_name = ref.split("/")[-1]
    defs = root_schema.get("$defs", {})
    if ref_name not in defs:
        raise LessonPlanError(f"Missing schema definition for ref: {ref_name}")
    return defs[ref_name]
