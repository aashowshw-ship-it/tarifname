"""Crash-safe, owner-scoped job snapshots for Patent Atolyesi.

Storage is private to the application process, not a browser session. Render free
filesystems are ephemeral: durable means across page reload/disconnect and
worker process restarts on the same filesystem, NOT across redeploys.
"""
from __future__ import annotations

import hashlib
import json
import os
import pickle
import secrets
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Any

ROOT = Path(os.getenv('PATENT_JOB_STORE', '/tmp/patent_atolyesi_jobs')).resolve()
MAX_FILES = 100
MAX_FILE_BYTES = 100 * 1024 * 1024

def _root() -> Path:
    ROOT.mkdir(mode=0o700, parents=True, exist_ok=True)
    return ROOT

def _id(value: str) -> str:
    if len(value) != 32 or any(c not in '0123456789abcdef' for c in value):
        raise ValueError('Geçersiz iş kimliği.')
    return value

def _folder(job_id: str) -> Path:
    return _root() / _id(job_id)

def _atomic_write(path: Path, data: bytes) -> None:
    fd, tmp = tempfile.mkstemp(prefix='.part-', dir=str(path.parent))
    try:
        with os.fdopen(fd, 'wb') as f:
            os.fchmod(f.fileno(), 0o600)
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)

def _json_write(path: Path, data: dict) -> None:
    _atomic_write(path, json.dumps(data, ensure_ascii=False, sort_keys=True).encode('utf-8'))

def _json_read(path: Path) -> dict:
    return json.loads(path.read_text(encoding='utf-8'))

def _owner(job_id: str, username: str) -> dict:
    item = _json_read(_folder(job_id) / 'meta.json')
    if not username or item.get('username') != username:
        raise PermissionError('Bu işe erişim yetkiniz yok.')
    return item

def _alive(pid: int) -> bool:
    if not isinstance(pid, int) or pid <= 0:
        return False
    try:
        os.kill(pid, 0)
        return True
    except (ProcessLookupError, ValueError):
        return False
    except PermissionError:
        return True

def _launch(job_id: str) -> None:
    folder = _folder(job_id)
    log = open(folder / 'worker.log', 'ab', buffering=0)
    try:
        subprocess.Popen(
            [sys.executable, '-m', 'job_worker', job_id],
            cwd=str(Path(__file__).parent),
            stdin=subprocess.DEVNULL,
            stdout=log,
            stderr=subprocess.STDOUT,
            close_fds=True,
            start_new_session=True,
            env={**os.environ, 'PATENT_WORKER_MODE': '1', 'PATENT_WORKER_JOB_ID': job_id},
        )
    finally:
        log.close()

def start_job(*, username: str, workflow: str, payload: dict, upload_groups: dict[str, list]) -> str:
    if workflow != 'tarifname_create':
        raise ValueError('Arka plan işlemi şu anda yalnız tarifname oluşturma için etkin.')
    if not username.strip():
        raise ValueError('Kullanıcı oturumu bulunamadı.')
    job_id = secrets.token_hex(16)
    folder = _folder(job_id)
    folder.mkdir(mode=0o700)
    blobs = folder / 'inputs'
    blobs.mkdir(mode=0o700)
    specs = {}
    count = 0
    for group, uploads in upload_groups.items():
        specs[group] = []
        for item in uploads:
            count += 1
            if count > MAX_FILES:
                raise ValueError('Tek iş için yüklenebilecek dosya sayısı aşıldı.')
            content = bytes(item.getvalue())
            if len(content) > MAX_FILE_BYTES:
                raise ValueError('Bir dosya izin verilen boyutu aşıyor.')
            digest = hashlib.sha256(content).hexdigest()
            _atomic_write(blobs / f'{count:04d}-{digest}', content)
            specs[group].append({
                'path':f'inputs/{count:04d}-{digest}',
                'name':str(item.name),
                'type':str(item.type or 'application/octet-stream'),
                'sha256':digest,
            })
    meta = dict(id=job_id, username=username, workflow=workflow,
                reference=str(payload.get('reference','')), payload=payload, uploads=specs,
                created=time.time(), status='queued', stage='Bekliyor', percent=0,
                artifacts=[], error='', pid=None, updated=time.time())
    _json_write(folder / 'meta.json', meta)
    try:
        _launch(job_id)
    except Exception as exc:
        update_status(job_id, status='failed', error=f'Çalışan süreç başlatılamadı: {exc}')
        raise
    return job_id

def read_job(job_id: str, username: str) -> dict:
    meta = _owner(job_id,username)
    if meta.get('status') == 'running' and not _alive(meta.get('pid')):
        meta = update_status(job_id, status='interrupted',
                             error='Çalışan süreç durdu. Son kalıcı aşamadan devam edilebilir.')
    return meta

def list_jobs(username: str, workflow: str, limit: int = 8) -> list[dict]:
    jobs = []
    for path in _root().iterdir():
        if not path.is_dir() or len(path.name) != 32:
            continue
        try:
            item = _json_read(path / 'meta.json')
        except (OSError, ValueError):
            continue
        if item.get('username') == username and item.get('workflow') == workflow:
            jobs.append(item)
    return sorted(jobs, key=lambda x:x.get('created',0), reverse=True)[:limit]

def update_status(job_id: str, **changes) -> dict:
    folder = _folder(job_id)
    data = _json_read(folder / 'meta.json')
    data.update(changes)
    data['updated'] = time.time()
    _json_write(folder / 'meta.json', data)
    return data

def restart_job(job_id: str, username: str) -> None:
    job = read_job(job_id, username)
    if job['status'] not in ('failed','interrupted'):
        raise ValueError('Bu iş yeniden başlatılamaz; zaten çalışıyor veya tamamlanmış.')
    update_status(job_id, status='queued', pid=None, error='',stage='Kaydedilen aşamadan devam bekliyor')
    _launch(job_id)

def _stage_path(workflow: str, signature: str, stage: str) -> Path:
    job_id = os.getenv('PATENT_WORKER_JOB_ID','')
    if not job_id or workflow != 'tarifname_create':
        raise ValueError('Kalıcı checkpoint yalnız arka plan işinde kullanılabilir.')
    key = hashlib.sha256(f'{workflow}\0{signature}\0{stage}'.encode()).hexdigest()
    folder = _folder(job_id) / 'checkpoints'
    folder.mkdir(mode=0o700,exist_ok=True)
    return folder / f'{key}.pickle'

def checkpoint_get(workflow: str, signature: str, stage: str) -> Any:
    path = _stage_path(workflow,signature,stage)
    if not path.is_file():
        return None
    # Only the trusted worker writes these private filesystem snapshots.
    with path.open('rb') as f:
        return pickle.load(f)

STAGE_PERCENT={'source_package':12,'extracted':27,'literature':39,
               'final_draft_audit':84,'docx':92,'figures':98}

def checkpoint_set(workflow: str, signature: str, stage: str, value: Any) -> None:
    path = _stage_path(workflow,signature,stage)
    _atomic_write(path,pickle.dumps(value,protocol=pickle.HIGHEST_PROTOCOL))
    job_id = os.environ['PATENT_WORKER_JOB_ID']
    update_status(job_id,stage=f'Tamamlandı: {stage}',percent=STAGE_PERCENT.get(stage,0))

def metric_get(workflow: str, signature: str) -> list[dict]:
    return checkpoint_get(workflow,signature,'__metrics__') or []

def metric_append(workflow: str, signature: str, metric: dict) -> None:
    rows = metric_get(workflow,signature)
    rows.append(dict(metric))
    path = _stage_path(workflow,signature,'__metrics__')
    _atomic_write(path,pickle.dumps(rows,protocol=pickle.HIGHEST_PROTOCOL))

def save_artifact(job_id: str, filename: str, data: bytes, label: str,
                  artifact_type: str = "tarifname", checks: dict | None = None) -> None:
    if '/' in filename or '\\' in filename or not filename.endswith('.docx'):
        raise ValueError('Geçersiz çıktı adı.')
    folder = _folder(job_id) / 'outputs'
    folder.mkdir(mode=0o700,exist_ok=True)
    _atomic_write(folder / filename,data)
    meta = _json_read(_folder(job_id)/'meta.json')
    existing = [a for a in meta.get('artifacts',[]) if a['filename'] != filename]
    existing.append({'filename':filename,'label':label,'artifact_type':artifact_type,
                     'checks':dict(checks or {})})
    update_status(job_id,artifacts=existing)

def get_artifact(job_id: str, filename: str, username: str) -> bytes:
    job = _owner(job_id,username)
    if job.get('status') != 'completed':
        raise PermissionError('Nihai kalite kapıları tamamlanmadan indirme açılamaz.')
    if filename not in [a['filename'] for a in job.get('artifacts',[])]:
        raise ValueError('Dosya bu işin onaylı çıktısı değil.')
    if '/' in filename or '\\' in filename:
        raise ValueError('Geçersiz dosya.')
    return (_folder(job_id) / 'outputs' / filename).read_bytes()

class SnapshotUpload:
    def __init__(self, folder: Path, spec: dict):
        self.name = spec['name']
        self.type = spec['type']
        path = folder / spec['path']
        self._data = path.read_bytes()
        if hashlib.sha256(self._data).hexdigest() != spec['sha256']:
            raise ValueError('Kaynak dosya bütünlüğü bozuldu; iş durduruldu.')
    def getvalue(self):
        return self._data


def load_job_inputs(job_id: str) -> tuple[dict,list,list,list,list]:
    folder = _folder(job_id)
    meta = _json_read(folder/'meta.json')
    groups={name:[SnapshotUpload(folder,spec) for spec in specs]
            for name,specs in meta['uploads'].items()}
    if len(groups['bbf']) != 1:
        raise ValueError('Tek BBF dosyası gereklidir.')
    return (meta['payload'],groups['bbf'][0],groups['extra_technical_files'],
            groups['example_files'],groups['figure_files'])
