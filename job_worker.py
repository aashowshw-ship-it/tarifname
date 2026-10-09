"""Detached workflow executor; invoked only by durable_jobs.start_job."""
from __future__ import annotations

import fcntl
import os
import sys
import traceback
from contextlib import nullcontext

import durable_jobs as jobs

class _Progress:
    def __init__(self, job_id: str):
        self.job_id = job_id
    def progress(self, value, text=None):
        jobs.update_status(self.job_id, percent=max(0,min(100,int(value))),
                           stage=str(text or 'İşleniyor'))
        return self

class JobUI:
    """Only presentation methods. Reuses actual domain pipeline and quality gates."""
    def __init__(self, job_id: str):
        self.job_id = job_id
        self.session_state={}
    def progress(self,value=0, text=None):
        return _Progress(self.job_id).progress(value,text)
    def expander(self,*args,**kwargs):
        return nullcontext()
    def download_button(self,label,*,data,file_name,mime=None,**kwargs):
        jobs.save_artifact(self.job_id,file_name,bytes(data),label)
        return True
    def exception(self,exc):
        raise exc
    def error(self,message,*args,**kwargs):
        raise RuntimeError(str(message))
    def stop(self):
        raise RuntimeError('İşlem durduruldu (zorunlu kural kapısı).')
    def __getattr__(self,name):
        if name in ('info','warning','success','caption','write','markdown','dataframe','text','table','metric'):
            return lambda *args,**kwargs: None
        raise AttributeError(f'Arka plan görünümünde desteklenmeyen işlem: {name}')

def main(job_id):
    folder=jobs._folder(job_id)
    with open(folder/'runner.lock','a+b') as lock:
        try:
            fcntl.flock(lock,fcntl.LOCK_EX|fcntl.LOCK_NB)
        except BlockingIOError:
            return  # second worker cannot concurrently modify checkpoints
        jobs.update_status(job_id,status='running',pid=os.getpid(),error='',stage='Kaydedilen aşamadan işleniyor')
        try:
            payload,bbf,extras,examples,figures=jobs.load_job_inputs(job_id)
            import app
            app.st=JobUI(job_id)
            # Preserve the SINGLE canonical compliance gate, including on re-download.
            def compliant_artifact(label, *, data, output_name, default_name, artifact_type, checks, audit_sources=None, **kwargs):
                from release_evidence import issue_delivery_evidence
                from rules import canonical_output_name
                proposed = canonical_output_name(output_name, default_name, artifact_type)
                proof = issue_delivery_evidence(artifact_type, bytes(data), audit_sources or [], checks, proposed)
                final_name = app.final_compliance_gate(
                    artifact_type, data=data, output_name=output_name,
                    default_name=default_name, checks=checks, audit_receipt=proof,
                )
                jobs.save_artifact(job_id, final_name, bytes(data), label, artifact_type, checks)
                return True
            app.compliant_download_button = compliant_artifact
            app.execute_tarifname_job(
                bbf, payload['reference'],payload['language_choice'],payload['claim_choice'],
                extras,examples,payload['separate_figures'],figures,
                payload['literature'],payload['lit_count'],payload['jurisdiction'],payload['extra_instruction'],
            )
            meta=jobs._json_read(folder/'meta.json')
            if not any('Tarifname_' in a['filename'] for a in meta.get('artifacts',[])):
                raise RuntimeError('Tarifname kalite kapıları tamamlanmadı; onaylı Word üretilmedi.')
            jobs.update_status(job_id,status='completed',percent=100,stage='Tamamlandı',pid=None)
        except BaseException as exc:
            traceback.print_exc()
            jobs.update_status(job_id,status='failed',pid=None,
                               error=f'{type(exc).__name__}: {exc}',stage='Hata: kaldığı yerden devam edilebilir')

if __name__=='__main__':
    if len(sys.argv)!=2:
        raise SystemExit('Usage: python -m job_worker <job_id>')
    os.environ['PATENT_WORKER_MODE']='1'
    os.environ['PATENT_WORKER_JOB_ID']=sys.argv[1]
    main(sys.argv[1])
