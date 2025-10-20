from pathlib import Path
import json
from typing import Any, Dict, Union

"""
getJSONdata.py

Load JSON data from variables.json into module-level globals so other modules can import them.

Behavior:
- Loads variables.json from the same directory as this file on import.
- Exposes VARIABLES (dict) and creates module-level variables for each top-level key
    (only if the key is a valid Python identifier).
- Provides reload(path=None) to reload from disk and get(key, default) to access values.
"""


_JSON_PATH = Path(__file__).resolve().parent / "variables.json"

# internal storage
_data: Dict[str, Any] = {}

def _load(path: Union[str, Path] = _JSON_PATH) -> None:
        global _data
        p = Path(path)
        try:
                with p.open("r", encoding="utf-8") as fh:
                        _data = json.load(fh) or {}
        except:
                _data = {}
        # Inject top-level keys as module globals (if valid identifiers)
        module_globals = globals()
        for key, val in list(_data.items()):
                if isinstance(key, str) and key.isidentifier():
                        module_globals[key] = val
        module_globals["VARIABLES"] = _data

def reload(path: Union[str, Path, None] = None) -> None:
        """
        Reload variables from disk.
        If path is None, reloads from the default variables.json next to this module.
        """
        _load(_JSON_PATH if path is None else path)

def get(key: str, default: Any = None) -> Any:
        """Return the value for key from the loaded variables, or default if missing."""
        return _data.get(key, default)

# Initial load at import time
_load()