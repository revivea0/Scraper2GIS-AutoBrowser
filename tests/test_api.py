import unittest
from unittest.mock import patch

from core.api import ParserAPI


class APITests(unittest.TestCase):
    def test_duplicate_start_does_not_replace_active_parser(self):
        api = ParserAPI()
        with patch('core.api.threading.Timer'):
            api.start_parsing('Город', 'Кафе', 'city', '', '', '')
            first = api.parser
            api.start_parsing('Город', 'Кафе', 'city', '', '', '')
            self.assertIs(api.parser, first)

    def test_whitespace_input_is_rejected(self):
        api = ParserAPI()
        with patch('core.api.threading.Timer'):
            api.start_parsing('  ', 'Кафе', 'city', '', '', '')
            self.assertIsNone(api.parser)

    def test_ui_failure_does_not_interrupt_parser(self):
        from unittest.mock import Mock
        api = ParserAPI()
        api.window = Mock()
        api.window.evaluate_js.side_effect = RuntimeError('window closed')
        api.log_to_ui('Получены данные')
        api.update_stats(1, 1)

    def test_closing_active_run_requests_stop_before_destroy(self):
        from unittest.mock import Mock
        api = ParserAPI()
        api.parser = Mock(is_running=True)
        api._run_lock.acquire()
        self.addCleanup(api._run_lock.release)
        self.assertFalse(api.on_closing())
        self.assertFalse(api.parser.is_running)

    def test_unsaved_rows_are_not_replaced_by_new_run(self):
        from core.parser import TwoGisParser
        api = ParserAPI()
        api.parser = TwoGisParser(lambda _: None, lambda *_: None)
        api.parser.results['123'] = {'ID': '123', 'Название': 'Не сохранено'}
        first = api.parser
        api._unsaved = True
        with patch('core.api.threading.Timer'):
            api.start_parsing('Москва', 'Кафе', 'city', '', '', '')
        self.assertIs(api.parser, first)
        self.assertIn('123', api.parser.results)

    def test_config_with_wrong_json_shape_is_ignored(self):
        import tempfile
        from pathlib import Path
        api = ParserAPI()
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder, 'config.json')
            path.write_text('[]')
            api._config_file = path
            self.assertIsInstance(api._load_config(), str)

    def test_validation_failure_releases_lock_after_previous_run(self):
        from core.parser import TwoGisParser
        api = ParserAPI()
        api.parser = TwoGisParser(lambda _: None, lambda *_: None)
        api._run_args = ('Москва', 'Кафе', 'city', '', '', '')
        api.start_parsing('', 'Кафе', 'city', '', '', '')
        self.assertFalse(api._run_lock.locked())

    def test_retry_save_preserves_in_memory_rows(self):
        import tempfile
        from pathlib import Path
        from core.parser import TwoGisParser
        api = ParserAPI()
        api.parser = TwoGisParser(lambda _: None, lambda *_: None)
        api.parser.results['123'] = {'ID': '123', 'Название': 'Кафе'}
        api._run_args = ('Москва', 'Кафе', 'city', '', '', '')
        api._unsaved = True
        with tempfile.TemporaryDirectory() as folder:
            path = api.retry_save(folder)
            self.assertTrue(Path(path).is_file())
        self.assertFalse(api._unsaved)
        self.assertIn('123', api.parser.results)

    def test_closing_handler_does_not_wait_for_javascript(self):
        from unittest.mock import Mock
        api = ParserAPI()
        api.window = Mock()
        api.parser = Mock(is_running=True)
        api._run_lock.acquire()
        self.addCleanup(api._run_lock.release)
        self.assertFalse(api.on_closing())
        api.window.evaluate_js.assert_not_called()
