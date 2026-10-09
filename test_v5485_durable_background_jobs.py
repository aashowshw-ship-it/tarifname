"""v5.4.85 persistent job and resume protection."""
import hashlib
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import durable_jobs as j

class Upload:
    name='bbf.txt'
    type='text/plain'
    def getvalue(self): return b'test input bytes'

class JobsTest(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory()
        self.patch=patch.object(j,'ROOT',Path(self.tmp.name))
        self.patch.start()
    def tearDown(self):
        self.patch.stop()
        self.tmp.cleanup()
    @patch.object(j,'_launch')
    def test_input_snapshot_and_ownership(self,mock_launch):
        jid=j.start_job(username='alice',workflow='tarifname_create',payload={'reference':'101'},
                        upload_groups={'bbf':[Upload()],'extra_technical_files':[],
                                       'example_files':[],'figure_files':[]})
        self.assertEqual(j.read_job(jid,'alice')['status'],'queued')
        with self.assertRaises(PermissionError): j.read_job(jid,'bob')
        vals=j.load_job_inputs(jid)
        self.assertEqual(vals[1].getvalue(),b'test input bytes')
        self.assertEqual(j.list_jobs('alice','tarifname_create')[0]['id'],jid)
        mock_launch.assert_called_once_with(jid)
    @patch.object(j,'_launch')
    def test_resume_and_checkpoints(self,mock_launch):
        jid=j.start_job(username='alice',workflow='tarifname_create',payload={},
                        upload_groups={'bbf':[Upload()],'extra_technical_files':[],
                                       'example_files':[],'figure_files':[]})
        with patch.dict(os.environ,{'PATENT_WORKER_JOB_ID':jid}):
            j.checkpoint_set('tarifname_create','s','extracted',{'facts':[1,2]})
            self.assertEqual(j.checkpoint_get('tarifname_create','s','extracted'),{'facts':[1,2]})
        j.update_status(jid,status='failed',error='API failure')
        j.restart_job(jid,'alice')
        self.assertEqual(j.read_job(jid,'alice')['status'],'queued')
        self.assertEqual(mock_launch.call_count,2)
        with patch.dict(os.environ,{'PATENT_WORKER_JOB_ID':jid}):
            self.assertEqual(j.checkpoint_get('tarifname_create','s','extracted'),{'facts':[1,2]})
    @patch.object(j,'_launch')
    def test_download_only_after_completion(self,_):
        jid=j.start_job(username='alice',workflow='tarifname_create',payload={},
                        upload_groups={'bbf':[Upload()],'extra_technical_files':[],
                                       'example_files':[],'figure_files':[]})
        j.save_artifact(jid,'Tarifname_101.docx',b'docx','tarifname')
        with self.assertRaises(PermissionError): j.get_artifact(jid,'Tarifname_101.docx','alice')
        j.update_status(jid,status='completed')
        self.assertEqual(j.get_artifact(jid,'Tarifname_101.docx','alice'),b'docx')
        with self.assertRaises(PermissionError):j.get_artifact(jid,'Tarifname_101.docx','other')

if __name__=='__main__': unittest.main()

class WorkerIntegrationTest(unittest.TestCase):
    def test_detached_worker_resumes_and_publishes_only_compliant_word(self):
        import sys, types
        from unittest.mock import patch
        import job_worker
        with tempfile.TemporaryDirectory() as tmp:
            with patch.object(j,'ROOT',Path(tmp)), patch.object(j,'_launch'):
                jid=j.start_job(username='user',workflow='tarifname_create',
                                payload={'reference':'101','language_choice':'Türkçe','claim_choice':"BBF'ye göre otomatik belirle",
                                         'separate_figures':False,'literature':False,'lit_count':0,'jurisdiction':'',
                                         'extra_instruction':''},
                                upload_groups={'bbf':[Upload()],'extra_technical_files':[],
                                               'example_files':[],'figure_files':[]})
                fake=types.ModuleType('app')
                from rules import final_compliance_gate, FINAL_COMPLIANCE_REQUIRED_CHECKS
                fake.final_compliance_gate=final_compliance_gate
                template=(Path(__file__).resolve().parent/'Tarifname_181176_template.docx').read_bytes()
                def pipeline(*args):
                    j.checkpoint_set('tarifname_create','sig','source_package',{'source':'full BBF'})
                    fake.compliant_download_button('Tarifname Word dosyasını indir',data=template,
                        output_name='Tarifname_101.docx',default_name='Tarifname_101.docx',
                        artifact_type='tarifname',
                        checks={key:True for key in FINAL_COMPLIANCE_REQUIRED_CHECKS['tarifname']},
                        audit_sources=[args[0]])
                fake.execute_tarifname_job=pipeline
                with patch.dict(os.environ,{'PATENT_WORKER_MODE':'1','PATENT_WORKER_JOB_ID':jid}), \
                     patch.dict(sys.modules,{'app':fake}):
                    job_worker.main(jid)
                item=j.read_job(jid,'user')
                self.assertEqual(item['status'],'completed')
                self.assertEqual(j.get_artifact(jid,'Tarifname_101.docx','user'),template)
                with patch.dict(os.environ,{'PATENT_WORKER_JOB_ID':jid}):
                    self.assertEqual(j.checkpoint_get('tarifname_create','sig','source_package'),{'source':'full BBF'})
