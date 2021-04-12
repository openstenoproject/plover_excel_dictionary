
from contextlib import contextmanager
from pathlib import Path
import os
import tempfile

from pyexcel_io.exceptions import SupportingPluginAvailableButNotInstalled
import pyexcel
import pytest

from plover import system
from plover.config import DEFAULT_SYSTEM_NAME
from plover.registry import registry

from plover_excel_dictionary import ExcelDictionary
import plover_excel_dictionary


TEST_DIR = Path(__file__).parent
TEST_FILES = {
    'ods': TEST_DIR / 'test.ods',
    'xlsx': TEST_DIR / 'test.xlsx',
}
TEST_FORMATS = 'ods xlsx'.split()

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
        ['S*P', 'not space!'],
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

def setup_module(cls):
    registry.update()
    system.setup(DEFAULT_SYSTEM_NAME)

@pytest.mark.parametrize('dict_format', TEST_FORMATS)
def test_format(dict_format):#, preferred_reader, preferred_writer):
    d_path = TEST_FILES[dict_format]
    d = ExcelDictionary.load(str(d_path))
    assert list(d.items()) == INITIAL_CONTENTS
    d[('S*P',)] = 'not space!'
    del d[('TEFGT',)]
    d[('TEFT', '-G')] = 'testing'
    with temp_dict(b'blah!', dict_format) as savename:
        d.path = savename
        d.save()
        book = pyexcel.get_book_dict(file_name=savename)
        assert list(book.items()) == MODIFIED_CONTENTS

@pytest.mark.parametrize('testcase', (
    'reader ods pyexcel-ods3',
    'writer ods pyexcel-ods3',
    'reader xlsx pyexcel-xlsx',
    'writer xlsx pyexcel-xlsx',
))
def test_preferred_readers_writers_detection(testcase):
    plugin_type, dict_format, preferred_plugin = testcase.split()
    attr = 'PREFERRED_' + plugin_type.upper()
    ext = '.' + dict_format
    assert getattr(plover_excel_dictionary, attr)[ext] == preferred_plugin

@pytest.mark.parametrize('dict_format', TEST_FORMATS)
def test_preferred_reader_is_used(dict_format, monkeypatch):
    monkeypatch.setattr('plover_excel_dictionary.PREFERRED_READER',
                        {'.' + dict_format: 'pouet'})
    with pytest.raises(SupportingPluginAvailableButNotInstalled):
        ExcelDictionary.load(str(TEST_FILES[dict_format]))

@pytest.mark.parametrize('dict_format', TEST_FORMATS)
def test_preferred_writer_is_used(dict_format, monkeypatch):
    monkeypatch.setattr('plover_excel_dictionary.PREFERRED_WRITER',
                        {'.' + dict_format: 'pouet'})
    d = ExcelDictionary.load(str(TEST_FILES[dict_format]))
    with temp_dict(b'', dict_format) as savename:
        d.path = savename
        with pytest.raises(SupportingPluginAvailableButNotInstalled):
            d.save()
