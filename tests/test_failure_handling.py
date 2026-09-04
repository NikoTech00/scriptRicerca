import importlib.util
import logging
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import Mock, patch

spec = importlib.util.spec_from_file_location('medici_under_test', Path(__file__).resolve().parents[1] / 'scriptMedici.py')
m = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = m
spec.loader.exec_module(m)


class FailureHandlingTests(unittest.TestCase):
    def client(self, provider='serper'):
        with patch.dict(m.os.environ, {'SERPER_API_KEY': 'test-secret', 'BRAVE_SEARCH_API_KEY': 'test-brave'}):
            return m.SearchClient(provider, 8, 0, 0)

    def test_http400_is_error_and_secret_is_redacted(self):
        response = Mock(status_code=400)
        response.json.return_value = {'message': 'Invalid test-secret'}
        with patch.object(m, 'requests', Mock(post=Mock(return_value=response))):
            with self.assertRaises(m.SearchRequestError) as caught:
                self.client().search('query')
        self.assertIn('400', str(caught.exception))
        self.assertNotIn('test-secret', str(caught.exception))

    def test_timeout_is_not_empty_result(self):
        client = self.client()
        with patch.object(client, '_search_provider', side_effect=TimeoutError('timeout')):
            with self.assertRaises(m.SearchRequestError):
                client.search('query')

    def test_valid_empty_response_stays_empty(self):
        client = self.client()
        with patch.object(client, '_search_provider', return_value=[]):
            self.assertEqual(client.search('query'), [])

    def test_auto_fallback_can_recover(self):
        client = self.client('auto')
        hit = m.WebHit('title', 'https://example.org', 'snippet', 'brave')
        with patch.object(client, '_search_provider', side_effect=[m.SearchRequestError('400'), [hit]]):
            self.assertEqual(client.search('query'), [hit])

    def test_locked_output_preserves_complete_recovery_file(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'result.xlsx'
            target.write_bytes(b'previous')
            workbook = Mock()
            workbook.save.side_effect = lambda path: path.write_bytes(b'complete-new-workbook')
            with patch.object(m.os, 'replace', side_effect=PermissionError('locked')), patch.object(m.time, 'sleep'):
                with self.assertRaisesRegex(RuntimeError, 'recupero'):
                    m.atomic_save(workbook, target)
            self.assertEqual(target.read_bytes(), b'previous')
            recovery = list(Path(directory).glob('.result_*.xlsx'))
            self.assertEqual(len(recovery), 1)
            self.assertEqual(recovery[0].read_bytes(), b'complete-new-workbook')

    def test_transient_lock_retries_and_replaces(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'result.xlsx'
            workbook = Mock()
            workbook.save.side_effect = lambda path: path.write_bytes(b'new')
            real_replace = m.os.replace
            calls = []
            def replace(src, dst):
                calls.append(src)
                if len(calls) == 1:
                    raise PermissionError('temporary lock')
                real_replace(src, dst)
            with patch.object(m.os, 'replace', side_effect=replace), patch.object(m.time, 'sleep'):
                m.atomic_save(workbook, target)
            self.assertEqual(target.read_bytes(), b'new')
            self.assertEqual(list(Path(directory).glob('.result_*.xlsx')), [])

    def test_cleanup_does_not_mask_save_error(self):
        with tempfile.TemporaryDirectory() as directory:
            workbook = Mock()
            workbook.save.side_effect = ValueError('original save failure')
            with patch.object(Path, 'unlink', side_effect=PermissionError('locked')):
                with self.assertRaisesRegex(ValueError, 'original save failure'):
                    m.atomic_save(workbook, Path(directory) / 'result.xlsx')


if __name__ == '__main__':
    logging.disable(logging.CRITICAL)
    unittest.main()
