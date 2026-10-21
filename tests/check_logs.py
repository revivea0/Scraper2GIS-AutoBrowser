"""Browser regression check: python tests/check_logs.py (installed Chromium required)."""
from pathlib import Path

from playwright.sync_api import sync_playwright

script = (Path(__file__).resolve().parents[1] / 'gui/index.html').read_text().split('<script>')[-1].split('</script>')[0]
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    try:
        page = browser.new_page()
        page.set_content('<div id="logs"></div>')
        page.add_script_tag(content=script)
        message = '<img src=x onerror="window.injected=true"> =1+1 & Кафе'
        page.evaluate('(message) => addLog(message)', message)
        assert page.locator('#logs img').count() == 0
        assert message in page.locator('#logs').inner_text()
        print('OK: log messages render as literal text')
    finally:
        browser.close()
