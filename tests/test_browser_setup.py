import unittest
from types import SimpleNamespace
from unittest.mock import AsyncMock, patch

from playwright.async_api import Error
from core.parser import TwoGisParser


class BrowserSetupTests(unittest.IsolatedAsyncioTestCase):
    async def test_existing_chrome_needs_no_download(self):
        parser = TwoGisParser(lambda _: None, lambda *_: None)
        browser = object()
        chromium = SimpleNamespace(launch=AsyncMock(side_effect=[Error('missing Edge'), browser]))
        with patch('core.parser.Path.is_file', return_value=False):
            self.assertIs(await parser._launch_browser(SimpleNamespace(chromium=chromium)), browser)

    async def test_missing_browsers_download_then_launch(self):
        parser = TwoGisParser(lambda _: None, lambda *_: None)
        browser = object()
        chromium = SimpleNamespace(executable_path='/missing/chromium',
            launch=AsyncMock(side_effect=[Error('missing Edge'), Error('missing Chrome'), browser]))
        process = SimpleNamespace(communicate=AsyncMock(return_value=(b'', None)), returncode=0)
        with patch('core.parser.Path.is_file', return_value=False), patch(
                'asyncio.create_subprocess_exec', new=AsyncMock(return_value=process)) as install:
            self.assertIs(await parser._launch_browser(SimpleNamespace(chromium=chromium)), browser)
            self.assertEqual(install.call_args.args[-2:], ('install', 'chromium'))
            self.assertNotIn('channel', chromium.launch.call_args.kwargs)

    async def test_cached_chromium_is_reused_without_download(self):
        parser = TwoGisParser(lambda _: None, lambda *_: None)
        browser = object()
        chromium = SimpleNamespace(executable_path='/cached/chromium',
            launch=AsyncMock(side_effect=[Error('missing Edge'), Error('missing Chrome'), browser]))
        with patch('core.parser.Path.is_file', return_value=True), patch(
                'asyncio.create_subprocess_exec', side_effect=AssertionError('Unexpected download')):
            self.assertIs(await parser._launch_browser(SimpleNamespace(chromium=chromium)), browser)

    async def test_failed_download_reports_error(self):
        parser = TwoGisParser(lambda _: None, lambda *_: None)
        chromium = SimpleNamespace(executable_path='/missing/chromium', launch=AsyncMock(side_effect=Error('missing')))
        process = SimpleNamespace(communicate=AsyncMock(return_value=(b'network unavailable', None)), returncode=1)
        with patch('core.parser.Path.is_file', return_value=False), patch(
                'asyncio.create_subprocess_exec', new=AsyncMock(return_value=process)):
            with self.assertRaisesRegex(RuntimeError, 'Chromium'):
                await parser._launch_browser(SimpleNamespace(chromium=chromium))
