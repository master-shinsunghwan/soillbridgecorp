"""Versioned, persistent catalog document uploads (original bytes retained)."""
import io
import sqlite3
import time
import uuid
import zipfile
from pathlib import Path

LIMIT = 50 * 1024 * 1024
KINDS = {'excel': {'.xls', '.xlsx'}, 'ppt': {'.ppt', '.pptx'}}


def connect(config):
    config.mkdir(parents=True, exist_ok=True)
    c = sqlite3.connect(config / 'catalog_status.db', timeout=20)
    c.execute('CREATE TABLE IF NOT EXISTS catalog_documents (kind TEXT PRIMARY KEY, name TEXT, file TEXT, revision INTEGER, updated TEXT, actor TEXT)')
    c.execute('CREATE TABLE IF NOT EXISTS catalog_document_history (kind TEXT, name TEXT, file TEXT, revision INTEGER, updated TEXT, actor TEXT)')
    return c


def documents(config):
    c = connect(config)
    try:
        rows = c.execute('SELECT kind,name,file,revision,updated,actor FROM catalog_documents').fetchall()
        return {r[0]: dict(zip(('name', 'file', 'revision', 'updated', 'actor'), r[1:])) for r in rows}
    finally:
        c.close()


def save(config, kind, name, raw, revision, actor):
    if kind not in KINDS or not isinstance(name, str) or not name or len(name) > 180 or any(ord(x) < 32 or x in '/\\' for x in name):
        raise ValueError('파일 이름을 확인해 주세요.')
    ext = Path(name).suffix.lower()
    if ext not in KINDS[kind] or not 0 < len(raw) <= LIMIT:
        raise ValueError('엑셀은 XLS/XLSX, PPT는 PPT/PPTX 파일을 50MB 이내로 선택해 주세요.')
    if ext in {'.xlsx', '.pptx'}:
        try:
            with zipfile.ZipFile(io.BytesIO(raw)) as z:
                required = 'xl/workbook.xml' if ext == '.xlsx' else 'ppt/presentation.xml'
                if '[Content_Types].xml' not in z.namelist() or required not in z.namelist():
                    raise ValueError()
        except (zipfile.BadZipFile, ValueError):
            raise ValueError('파일 내용이 선택한 문서 형식과 일치하지 않습니다.') from None
    elif not raw.startswith(bytes.fromhex('D0CF11E0A1B11AE1')):
        raise ValueError('정상적인 XLS 또는 PPT 문서가 아닙니다.')
    c = connect(config)
    folder = config / 'catalog_documents'
    folder.mkdir(exist_ok=True)
    target = folder / (uuid.uuid4().hex + ext)
    try:
        c.execute('BEGIN IMMEDIATE')
        row = c.execute('SELECT revision FROM catalog_documents WHERE kind=?', (kind,)).fetchone()
        if (row[0] if row else 0) != revision:
            raise ValueError('다른 직원이 파일을 변경했습니다. 파일 목록을 새로고침한 후 다시 등록해 주세요.')
        target.write_bytes(raw)
        values = (kind, name, target.name, revision + 1, time.strftime('%Y-%m-%dT%H:%M:%SZ', time.gmtime()), actor)
        c.execute('INSERT OR REPLACE INTO catalog_documents VALUES (?,?,?,?,?,?)', values)
        c.execute('INSERT INTO catalog_document_history VALUES (?,?,?,?,?,?)', values)
        c.commit()
    except Exception:
        target.unlink(missing_ok=True)
        raise
    finally:
        c.close()
    return documents(config)
