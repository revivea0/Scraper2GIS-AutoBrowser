import asyncio
import tempfile
import unittest
from pathlib import Path

import openpyxl

from core.parser import TwoGisParser


class ReliabilityTests(unittest.TestCase):
    def setUp(self):
        self.logs = []
        self.parser = TwoGisParser(self.logs.append, lambda *_: None)
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def write_backup(self, rows):
        path = Path(self.tmp.name, 'STOPPED_page_5_Москва_Кафе.xlsx')
        wb = openpyxl.Workbook()
        wb.active.append(['ID', 'Название', 'Телефоны'])
        for row in rows:
            wb.active.append(row)
        wb.save(path)
        wb.close()
        return path

    def test_duplicate_backup_keeps_nonempty_contacts(self):
        path = self.write_backup([['123_old', 'Кафе', '+79990000000'], ['123_new', 'Кафе', None]])
        self.parser._load_backup(str(path))
        self.assertEqual(self.parser.results['123']['Телефоны'], '+79990000000')

    def test_duplicate_backup_keeps_both_nonempty_phone_numbers(self):
        path = self.write_backup([['123_old', 'Кафе', '+79990000001'], ['123_new', 'Кафе', '+79990000002']])
        self.parser._load_backup(str(path))
        self.assertEqual(self.parser.results['123']['Телефоны'], '+79990000001, +79990000002')

    def test_rounded_numeric_id_is_rejected(self):
        path = self.write_backup([[70000001012345678, 'Кафе', '']])
        with self.assertRaisesRegex(ValueError, 'ID'):
            self.parser._load_backup(str(path))

    def test_metadata_restores_page_after_file_rename(self):
        self.parser.results['123'] = {'ID': '123', 'Название': 'Кафе'}
        self.parser._search_context = {'city': 'Москва', 'query': 'Кафе', 'scope': 'city', 'detail': ''}
        self.parser._page_complete = True
        path = self.parser._export_to_excel('Москва', 'Кафе', self.tmp.name, is_backup=True, current_page=5)
        renamed = Path(self.tmp.name, 'renamed.xlsx')
        Path(path).rename(renamed)
        restored = TwoGisParser(self.logs.append, lambda *_: None)
        self.assertEqual(restored._load_backup(str(renamed), self.parser._search_context), 6)
        with self.assertRaisesRegex(ValueError, 'параметр'):
            restored._load_backup(str(renamed), {'city': 'Казань', 'query': 'Кафе', 'scope': 'city', 'detail': ''})

    def test_next_page_does_not_match_page_twenty(self):
        hrefs = ['/moscow/search/cafe/page/20', '/moscow/search/cafe/page/2?m=1']
        self.assertEqual(self.parser._next_page_url(hrefs, 'https://2gis.ru/moscow/search/cafe', 2), 'https://2gis.ru/moscow/search/cafe/page/2?m=1')

    def test_card_ids_are_deduplicated(self):
        hrefs = ['/moscow/firm/123', '/moscow/firm/123_session?x=1', '/firm/456', '/firm/not-a-firm']
        self.assertEqual(self.parser._card_ids(hrefs), ['123', '456'])

    def test_advertising_redirect_is_not_a_search_card(self):
        hrefs = ['http://link.2gis.ru/ad?https://2gis.ru/firm/999', '/moscow/firm/123?stat=x']
        self.assertEqual(self.parser._card_ids(hrefs), ['123'])

    def test_final_export_does_not_overwrite_previous_result(self):
        self.parser.results['123'] = {'ID': '123', 'Название': 'Первый'}
        first = self.parser._export_to_excel('Москва', 'Кафе', self.tmp.name)
        before = Path(first).read_bytes()
        self.parser.results['123']['Название'] = 'Второй'
        second = self.parser._export_to_excel('Москва', 'Кафе', self.tmp.name)
        self.assertNotEqual(first, second)
        self.assertEqual(Path(first).read_bytes(), before)

    def test_null_contacts_do_not_drop_following_items(self):
        class Response:
            status = 200
            url = 'https://catalog.api.2gis.ru/3.0/items/byid?id=123,456'

            async def json(self):
                return {'result': {'items': [
                    {'id': '123', 'name': 'Первое', 'contact_groups': None, 'rubrics': None},
                    {'id': '456', 'name': 'Второе', 'contact_groups': [{'contacts': [{'type': 'phone', 'value': '+7999'}]}]}
                ]}}

        asyncio.run(self.parser._intercept_api_response(Response()))
        self.assertEqual(set(self.parser.results), {'123', '456'})
        self.assertEqual(self.parser.results['456']['Телефоны'], '+7999')

    def test_successful_response_marks_firm_processed(self):
        class Response:
            status = 200
            url = 'https://catalog.api.2gis.ru/3.0/items/byid?id=123'

            async def json(self):
                return {'result': {'items': [{'id': '123', 'name': 'Кафе'}]}}

        asyncio.run(self.parser._intercept_api_response(Response()))
        self.assertIn('123', self.parser.processed_firm_ids)
