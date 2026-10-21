"""UI smoke check: python tests/check_ui.py (installed Chromium required)."""
from pathlib import Path
from playwright.sync_api import sync_playwright

root = Path(__file__).resolve().parents[1]
output = root / 'output' / 'design'
output.mkdir(parents=True, exist_ok=True)
with sync_playwright() as p:
    browser = p.chromium.launch()
    try:
        page = browser.new_page(viewport={'width': 1100, 'height': 760})
        page.add_init_script("""
            window.calls = [];
            window.pywebview = {api: {
                get_saved_dir: async () => '/Users/test/Results',
                select_folder: async () => '/Users/test/Results',
                select_backup: async () => '/Users/test/STOPPED_page_1.xlsx',
                start_parsing: async (...args) => window.calls.push(['start', ...args]),
                stop_parsing: async () => window.calls.push(['stop']),
                open_excel: async path => window.calls.push(['open', path]),
                retry_save: async folder => { window.calls.push(['retry', folder]); return folder + '/recovered.xlsx'; }
            }};
        """)
        page.goto((root / 'gui/index.html').as_uri())
        page.evaluate("window.dispatchEvent(new Event('pywebviewready'))")
        page.get_by_label('Город', exact=True).fill('Красноярск')
        page.get_by_label('Рубрика или название', exact=True).fill('Кофейня')
        page.locator('h1').click()
        page.screenshot(path=str(output / 'desktop.png'), full_page=True)
        page.select_option('#scopeSelect', 'street')
        assert page.locator('#detailInput').is_visible()
        page.locator('#detailInput').fill('Мира')
        page.locator('#startBtn').click()
        assert page.locator('#cityInput').is_disabled()
        assert page.locator('#stopBtn').is_visible()
        assert page.evaluate('window.calls[0]') == ['start', 'Красноярск', 'Кофейня', 'street', 'Мира', '/Users/test/Results', '']
        page.evaluate("updateStats(12, 2); addLog('+ Кофейня <img src=x>');")
        assert page.locator('#statCount').inner_text() == '12'
        assert page.locator('#logs img').count() == 0
        assert not page.locator('#logsEmpty').is_visible()
        assert page.locator('#stopBtn').bounding_box()['y'] + page.locator('#stopBtn').bounding_box()['height'] <= 760
        page.screenshot(path=str(output / 'running.png'), full_page=True)
        page.locator('#stopBtn').click()
        assert page.evaluate('window.calls[1]') == ['stop']
        assert page.locator('#stopBtn').is_disabled()
        page.evaluate("parsingStopped(); showCompletionDialog('/Users/test/Results/result.xlsx');")
        assert page.locator('#cityInput').is_enabled()
        page.get_by_role('button', name='Открыть Excel', exact=True).click()
        assert page.evaluate('window.calls[2]') == ['open', '/Users/test/Results/result.xlsx']
        page.evaluate("showCompletionDialog('/Users/test/Results/result.xlsx');")
        page.screenshot(path=str(output / 'saved.png'), full_page=True)
        page.evaluate('closeModal()')
        page.locator('#cityInput').fill(' ')
        page.locator('#startBtn').click()
        assert len(page.evaluate('window.calls')) == 3
        assert page.locator('#cityInput').get_attribute('aria-invalid') == 'true'
        for width, height, name in [(820, 700, 'compact'), (390, 844, 'mobile')]:
            page.set_viewport_size({'width': width, 'height': height})
            page.screenshot(path=str(output / f'{name}.png'), full_page=True)
            assert page.evaluate('document.documentElement.scrollWidth <= window.innerWidth')
        page.evaluate('showSaveError()')
        assert page.locator('#startBtn').is_disabled()
        page.get_by_role('button', name='Сохранить в другую папку', exact=True).click()
        page.wait_for_function("document.getElementById('completionPath').textContent === 'recovered.xlsx'")
        assert page.locator('#startBtn').is_enabled()
        assert page.evaluate('window.calls.at(-1)') == ['retry', '/Users/test/Results']
        print('OK: save recovery, form, filters, start/stop, bridge arguments, stats, safe logs, export, validation, responsive layouts')
    finally:
        browser.close()
