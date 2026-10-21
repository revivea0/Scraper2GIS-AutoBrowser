import asyncio
import json
import logging
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import threading

import webview

from core.parser import TwoGisParser
from core.paths import default_save_dir, user_data_dir


class ParserAPI:
    def __init__(self):
        self.window = None
        self.parser = None
        self._run_lock = threading.Lock()
        self._close_requested = False
        self._unsaved = False
        self._run_args = None
        self._config_file = user_data_dir() / 'config.json'
        self.save_dir = self._load_config()

    def _load_config(self):
        path = Path(self._config_file)
        if not path.exists():
            path = Path('scraper_config.json')
        try:
            config = json.loads(path.read_text(encoding='utf-8'))
            if isinstance(config, dict) and isinstance(config.get('save_dir'), str) and config['save_dir']:
                return config['save_dir']
        except (OSError, ValueError):
            pass
        return str(default_save_dir())

    def _save_config(self, path):
        temp_path = None
        try:
            target = Path(self._config_file)
            target.parent.mkdir(parents=True, exist_ok=True)
            with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=target.parent, delete=False) as temp:
                temp_path = temp.name
                json.dump({'save_dir': path}, temp, ensure_ascii=False)
            os.replace(temp_path, target)
            temp_path = None
        except OSError as error:
            self.log_to_ui(f'Ошибка сохранения настроек: {error}')
        finally:
            if temp_path:
                Path(temp_path).unlink(missing_ok=True)

    def _evaluate(self, script):
        if self.window:
            try:
                return self.window.evaluate_js(script)
            except Exception:
                logging.getLogger(__name__).debug('Окно интерфейса недоступно', exc_info=True)

    def get_saved_dir(self):
        return self.save_dir

    def set_window(self, window):
        self.window = window

    def log_to_ui(self, message):
        logging.getLogger(__name__).info('%s', message)
        self._evaluate(f'addLog({json.dumps(str(message))})')

    def update_stats(self, count, page):
        self._evaluate(f'updateStats({int(count)}, {int(page)})')

    def select_folder(self):
        result = self.window.create_file_dialog(webview.FileDialog.FOLDER)
        if result:
            self.save_dir = result[0]
            self._save_config(self.save_dir)
            return self.save_dir
        return ''

    def select_backup(self):
        result = self.window.create_file_dialog(webview.FileDialog.OPEN, allow_multiple=False,
                                                file_types=('Excel Files (*.xlsx)',))
        return result[0] if result else ''

    def stop_parsing(self):
        if self.parser:
            self.parser.is_running = False
            self.log_to_ui('Сигнал остановки получен. Сохраняем прогресс...')

    def on_closing(self):
        if self._run_lock.locked():
            self._close_requested = True
            if self.parser:
                self.parser.is_running = False
            # This callback runs on the GUI thread; evaluate_js here would deadlock it.
            logging.getLogger(__name__).info('Закрытие окна: ожидаем сохранения данных.')
            return False
        if self._unsaved and self.window:
            return self.window.create_confirmation_dialog(
                'Данные не сохранены', 'Не удалось сохранить Excel. Закрыть окно и потерять несохранённые данные?')
        return True

    def open_excel(self, filepath):
        try:
            path = Path(filepath).resolve()
            if path.suffix.lower() != '.xlsx' or not path.is_file():
                raise ValueError('Выберите существующий файл .xlsx')
            if sys.platform == 'win32':
                os.startfile(str(path))
            else:
                subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', str(path)], check=True)
        except (OSError, ValueError, subprocess.CalledProcessError) as error:
            self.log_to_ui(f'Ошибка открытия файла: {error}')

    def start_parsing(self, city, query, scope, detail, save_dir, backup_path):
        if not self._run_lock.acquire(blocking=False):
            return False
        started = False
        try:
            if self._unsaved:
                self._evaluate('parsingStopped(); showSaveError()')
                return False
            if not all(isinstance(value, str) for value in (city, query, scope, detail, save_dir, backup_path)):
                raise ValueError('Некорректные параметры поиска')
            city, query, detail = city.strip(), query.strip(), detail.strip()
            if not city or not query or scope not in ('city', 'district', 'street') or (scope != 'city' and not detail):
                raise ValueError('Укажите город, рубрику и название выбранного района или улицы')
            self.parser = TwoGisParser(self.log_to_ui, self.update_stats)
            self._run_args = (city, query, scope, detail, save_dir or self.save_dir, backup_path)
            threading.Timer(0, self._run_thread_isolated, args=self._run_args).start()
            started = True
            return True
        except Exception as error:
            self.log_to_ui(f'Ошибка запуска: {error}')
            self._evaluate('parsingStopped()')
            return False
        finally:
            # The worker owns the lock only after the timer starts successfully.
            if not started:
                self._run_lock.release()

    def _run_thread_isolated(self, city, query, scope, detail, save_dir, backup_path):
        loop = None
        filepath = ''
        try:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            self.log_to_ui(f'Инициализация: {city} → {query}')
            filepath = loop.run_until_complete(self.parser.run(city, query, scope, detail, save_dir, backup_path))
        except Exception as error:
            logging.getLogger(__name__).exception('Сбор завершился с ошибкой')
            self.log_to_ui(f'Ошибка сбора: {error}')
        finally:
            self._unsaved = bool(self.parser and self.parser.results and not filepath)
            self._evaluate('parsingStopped()')
            if filepath:
                self._evaluate(f'showCompletionDialog({json.dumps(filepath)})')
            elif self._unsaved:
                self._evaluate('showSaveError()')
            if loop:
                loop.close()
            self._run_lock.release()
            if self._close_requested and not self._unsaved and self.window:
                self.window.destroy()
            self._close_requested = False

    def retry_save(self, save_dir):
        if not self._run_lock.acquire(blocking=False):
            return ''
        try:
            if not self._unsaved or not self.parser or not self._run_args:
                return ''
            directory = Path(save_dir or self.save_dir).resolve()
            directory.mkdir(parents=True, exist_ok=True)
            city, query = self._run_args[:2]
            filepath = self.parser._export_to_excel(city, query, str(directory), current_page=self.parser.current_page,
                                                    prefix_override='STOPPED_')
            if filepath:
                self._unsaved = False
                self._evaluate(f'showCompletionDialog({json.dumps(filepath)})')
            else:
                self._evaluate('showSaveError()')
            return filepath
        except OSError as error:
            self.log_to_ui(f'Ошибка сохранения: {error}')
            return ''
        finally:
            self._run_lock.release()
