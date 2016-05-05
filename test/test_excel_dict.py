
from contextlib import contextmanager
import os
import tempfile
import unittest

import pyexcel

from plover import system
from plover.config import DEFAULT_SYSTEM_NAME
from plover.registry import registry

from plover_excel_dictionary import ExcelDictionary


INITIAL_CONTENTS = [
    (('TEFT', '-D'), 'tested'),
    (('TEFGT',), 'testing'),
    (('S-P',), ''),
    (('S*P',), ''),
    (('R-R',), '{^}\n{^}{-|}'),
    (('R*R',), 'not a tab: `\\t`'),
    (('SPAS',), ' space '),
    (('START',), 'start'),
    (('EUPB', 'SAOEUT', '-FL'), 'insightful'),
    (('SHR*UG', 'SHR*UG'), '¯\\\\_(ツ)_/¯'),
]
MODIFIED_CONTENTS = [
    ('Sheet1', [
        ['TEFT/-D', 'tested', 'insightful comment'],
        ['S-P', '', ''],
        ['S*P', 'not space!', 'blah blah'],
        ['R-R', '{^}\n{^}{-|}', '\\n → newline'],
        ['R*R', 'not a tab: `\\t`', ''],
        ['SPAS', ' space ', ''],
    ]),
    ('Sheet3', [
        ['START', 'start'],
        ['EUPB/SAOEUT/-FL', 'insightful'],
    ]),
    ('Sheet4', [
        ['SHR*UG/SHR*UG', '¯\\\\_(ツ)_/¯'],
    ]),
    ('NEW', [
        ['TEFT/-G', 'testing'],
    ]),
]


@contextmanager
def temp_dict(contents, extension):
    tf = tempfile.NamedTemporaryFile(delete=False, suffix='.'+extension)
    try:
        tf.write(contents)
        tf.close()
        yield tf.name
    finally:
        os.unlink(tf.name)


class ExcelDictionaryTestCase(object):

    FORMAT = None

    @classmethod
    def setUpClass(cls):
        registry.update()
        system.setup(DEFAULT_SYSTEM_NAME)

    def test_1(self):
        d_path = os.path.join(os.path.dirname(__file__), 'test.' + self.FORMAT)
        d = ExcelDictionary.load(d_path)
        self.assertEqual(list(d.items()), INITIAL_CONTENTS)
        d[('S*P',)] = 'not space!'
        del d[('TEFGT',)]
        d[('TEFT', '-G')] = 'testing'
        with temp_dict(b'blah!', self.FORMAT) as savename:
            d.path = savename
            d.save()
            book = pyexcel.get_book_dict(file_name=savename)
            self.assertEqual(list(book.items()), MODIFIED_CONTENTS)

class OdsDictionaryTestCase(ExcelDictionaryTestCase, unittest.TestCase):
    FORMAT = 'ods'

class XlsxDictionaryTestCase(ExcelDictionaryTestCase, unittest.TestCase):
    FORMAT = 'xlsx'
