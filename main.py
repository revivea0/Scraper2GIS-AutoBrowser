import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

import webview

from core.api import ParserAPI
from core.paths import user_data_dir


if __name__ == '__main__':
    log_dir = user_data_dir()
    log_dir.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s', handlers=[
        RotatingFileHandler(log_dir / 'scraper.log', maxBytes=1_000_000, backupCount=2, encoding='utf-8')
    ])
    api = ParserAPI()
    window = webview.create_window(
        title='Scraper 2GIS',
        url=str(Path(__file__).resolve().parent / 'gui' / 'index.html'),
        js_api=api,
        width=1100,
        height=800,
        min_size=(760, 640),
        background_color='#f4f5f7',
    )
    api.set_window(window)
    window.events.closing += api.on_closing
    webview.start(debug=False)
