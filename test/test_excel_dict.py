from pathlib import Path
import itertools

from pyexcel_io.exceptions import SupportingPluginAvailableButNotInstalled
import pyexcel
import pytest

from plover_build_utils.testing import dictionary_test, make_dict

from plover_excel_dictionary import ExcelDictionary
import plover_excel_dictionary


TEST_DIR = Path(__file__).parent
TEST_FILES = {
    'ods': TEST_DIR / 'test.ods',
    'xlsx': TEST_DIR / 'test.xlsx',
}
TEST_FORMATS = 'ods xlsx'.split()
# Note: preferred first.
TEST_READERS = {
    'ods': ('pyexcel-ods', 'pyexcel-ods3'),
    'xlsx': ('pyexcel-xlsx',),
}
TEST_WRITERS = {
    'ods': ('pyexcel-ods', 'pyexcel-ods3'),
    'xlsx': ('pyexcel-libxlsxw', 'pyexcel-xlsx'),
}

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


FORMAT_TESTS = list(itertools.chain(*(
    itertools.product((dict_format,),
                      TEST_READERS[dict_format],
                      TEST_WRITERS[dict_format])
    for dict_format in TEST_FORMATS
)))

@pytest.mark.parametrize('dict_format, preferred_reader, preferred_writer', FORMAT_TESTS)
def test_format(tmp_path, dict_format, preferred_reader, preferred_writer, monkeypatch):
    monkeypatch.setattr('plover_excel_dictionary.PREFERRED_READER',
                        {'.' + dict_format: preferred_reader})
    monkeypatch.setattr('plover_excel_dictionary.PREFERRED_WRITER',
                        {'.' + dict_format: preferred_writer})
    d_path = TEST_FILES[dict_format]
    d = ExcelDictionary.load(str(d_path))
    assert list(d.items()) == INITIAL_CONTENTS
    d[('S*P',)] = 'not space!'
    del d[('TEFGT',)]
    d[('TEFT', '-G')] = 'testing'
    with make_dict(tmp_path, b'blah!', extension=dict_format) as savename:
        d.path = str(savename)
        d.save()
        book = pyexcel.get_book_dict(file_name=str(savename))
        assert list(book.items()) == MODIFIED_CONTENTS

@pytest.mark.parametrize('plugin_type, dict_format, preferred_plugin', (
    ('reader', 'ods', TEST_READERS['ods'][0]),
    ('writer', 'ods', TEST_WRITERS['ods'][0]),
    ('reader', 'xlsx', TEST_READERS['xlsx'][0]),
    ('writer', 'xlsx', TEST_WRITERS['xlsx'][0]),
))
def test_preferred_readers_writers_detection(plugin_type, dict_format, preferred_plugin):
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
def test_preferred_writer_is_used(tmp_path, dict_format, monkeypatch):
    monkeypatch.setattr('plover_excel_dictionary.PREFERRED_WRITER',
                        {'.' + dict_format: 'pouet'})
    d = ExcelDictionary.load(str(TEST_FILES[dict_format]))
    with make_dict(tmp_path, b'', extension=dict_format) as savename:
        d.path = str(savename)
        with pytest.raises(SupportingPluginAvailableButNotInstalled):
            d.save()


class _TestDictionary:

    DICT_CLASS = ExcelDictionary
    DICT_REGISTERED = True
    DICT_SAMPLE = 'test'
    DICT_LOAD_TESTS = [
        lambda: (
            'test',
            '\n'.join(
                '%r: %r,' % ('/'.join(k), v)
                for k, v in INITIAL_CONTENTS
            )
        ),
    ]
    DICT_SAVE_TESTS = [
        lambda: (t()[1], None)
        for t in DICT_LOAD_TESTS
    ]

    @classmethod
    def make_dict(cls, contents):
        if isinstance(contents, bytes):
            return contents
        path = Path(__file__).parent / (contents + '.' + cls.DICT_EXTENSION)
        return path.read_bytes()

@dictionary_test
class TestOdsDictionary(_TestDictionary):

    DICT_EXTENSION = 'ods'

@dictionary_test
class TestXlsxDictionary(_TestDictionary):

    DICT_EXTENSION = 'xlsx'
