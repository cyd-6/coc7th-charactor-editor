"""Shared default template identity for the application and maintenance tools."""
import json
from pathlib import Path

TEMPLATE_DIRECTORY = Path(__file__).resolve().parents[1] / 'assets' / 'templates'
_metadata = json.loads((TEMPLATE_DIRECTORY / 'manifest.json').read_text(encoding='utf-8'))
TEMPLATE_VERSION = _metadata['version']
TEMPLATE_FILENAME = _metadata['filename']
TEMPLATE_PATH = TEMPLATE_DIRECTORY / TEMPLATE_FILENAME
