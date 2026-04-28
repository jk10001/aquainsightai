from __future__ import annotations
from typing import Any, MutableMapping
import os
import json
from pathlib import Path
from termcolor import cprint
from dotenv import load_dotenv

load_dotenv()


# Alias type used throughout the app
Alias = str

# Map each alias to its configuration options. Each entry must supply at least
# ``model``, ``key_env`` and ``api_type``. Additional keys will be propagated
# to :func:`cfg` callers.
models_file = Path(__file__).with_name("models.json")
with open(models_file, "r") as f:
    _LOOKUP: dict[Alias, MutableMapping[str, Any]] = json.load(f)


# Map of alias -> True if the corresponding environment variable exists and is non-empty
def _has_api_key(entry: MutableMapping[str, Any]) -> bool:
    """Return ``True`` if the configured API key environment variable is set."""

    raw_value = os.getenv(entry["key_env"])
    if raw_value is None:
        return False

    if isinstance(raw_value, str):
        return raw_value.strip() != ""

    return bool(raw_value)


API_KEY_STATUS: dict[Alias, bool] = {
    alias: _has_api_key(entry) for alias, entry in _LOOKUP.items()
}

_MISSING_KEYS = [alias for alias, has_key in API_KEY_STATUS.items() if not has_key]
if _MISSING_KEYS:
    cprint(
        "Warning: missing API keys for LLM aliases: "
        + ", ".join(sorted(_MISSING_KEYS)),
        "red",
    )


# list of all available model aliases
ALIASES = list(_LOOKUP.keys())

# list of all vision-capable model aliases
ALIASES_VISION = [
    alias
    for alias, entry in _LOOKUP.items()
    if entry.get("metadata", {}).get("vision_capable")
]

# Metadata for frontend role availability and other model capabilities.
MODEL_METADATA: dict[Alias, MutableMapping[str, Any]] = {
    alias: entry.get("metadata", {}) for alias, entry in _LOOKUP.items()
}

# Additional information for frontend (e.g., parameter schemas)
MODEL_INFO: dict[Alias, MutableMapping[str, Any]] = {
    alias: entry.get("additional_info", {}) for alias, entry in _LOOKUP.items()
}


def cfg(alias: Alias, **extra: Any) -> list[dict]:
    """Return ag2-style ``config_list`` for ``alias``.

    Any ``extra`` keyword arguments supplied are merged into the returned
    configuration. In addition, entries in ``_LOOKUP`` may themselves define
    extra key-value pairs which are also included in the result, except for
    'additional_info', which is not propagated.
    """
    entry = _LOOKUP[alias]
    base = {
        "model": entry["model"],
        "api_key": os.getenv(entry["key_env"]),
        "api_type": entry["api_type"],
    }

    if not base["api_key"] or (
        isinstance(base["api_key"], str) and base["api_key"].strip() == ""
    ):
        raise RuntimeError(
            f"API key environment variable '{entry['key_env']}' is not set for alias '{alias}'."
        )

    for k, v in entry.items():
        if k not in {"model", "key_env", "api_type", "metadata", "additional_info"}:
            base[k] = v

    if "reasoning_effort" in extra and base["api_type"] == "responses":
        extra["reasoning"] = {"effort": extra.pop("reasoning_effort")}

    base.update(extra)
    return [base]
