import sys
from pathlib import Path

_plugin_root = str(Path(__file__).resolve().parent)
if _plugin_root not in sys.path:
    sys.path.insert(0, _plugin_root)

from hermes_outlook.plugin import register

__all__ = ["register"]
