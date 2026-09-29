"""Keep application, maintenance tools and downloadable template in sync."""
import json
import os
import shutil
import subprocess
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app import TEMPLATE_PATH as APP_TEMPLATE_PATH, app, get_catalog
from coc7_card.template_config import TEMPLATE_FILENAME, TEMPLATE_PATH, TEMPLATE_VERSION

ROOT = Path(__file__).resolve().parents[1]


def test_default_template_version_and_download():
    assert TEMPLATE_VERSION == '26.3'
    assert TEMPLATE_FILENAME == f'COC7空白卡CY{TEMPLATE_VERSION}.xlsx'
    assert APP_TEMPLATE_PATH == TEMPLATE_PATH == get_catalog().template_path
    assert f'版本号CY{TEMPLATE_VERSION}' in get_catalog().template_revision_note
    with TestClient(app) as client:
        bootstrap = client.get('/api/bootstrap')
        assert bootstrap.status_code == 200
        meta = bootstrap.json()['meta']
        assert meta['template_filename'] == TEMPLATE_FILENAME
        assert meta['template_revision_note'] == get_catalog().template_revision_note
        download = client.get(f'/assets/templates/{TEMPLATE_FILENAME}')
        assert download.status_code == 200
        assert download.content == TEMPLATE_PATH.read_bytes()


def test_node_maintenance_tools_use_same_template(tmp_path):
    node = os.environ.get('COC7_NODE') or shutil.which('node')
    if not node:
        pytest.skip('Node.js is required to verify maintenance tool configuration')
    module = (ROOT / 'scripts/template_config.mjs').as_uri()
    result = subprocess.run(
        [node, '--input-type=module', '-e',
         f'import {{ templatePath, templateVersion }} from {json.dumps(module)}; '
         'console.log(JSON.stringify({templatePath, templateVersion}));'],
        cwd=tmp_path, check=True, capture_output=True, encoding='utf-8', timeout=15,
    )
    actual = json.loads(result.stdout)
    assert Path(actual['templatePath']) == TEMPLATE_PATH
    assert actual['templateVersion'] == TEMPLATE_VERSION
