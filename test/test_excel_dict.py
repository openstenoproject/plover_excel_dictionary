
from contextlib import contextmanager
import os
import tempfile

import pyexcel
import pytest

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

@pytest.mark.parametrize('dict_format', 'ods xlsx'.split())
def test_1(dict_format):
    d_path = os.path.join(os.path.dirname(__file__), 'test.' + dict_format)
    d = ExcelDictionary.load(d_path)
    assert list(d.items()) == INITIAL_CONTENTS
    d[('S*P',)] = 'not space!'
    del d[('TEFGT',)]
    d[('TEFT', '-G')] = 'testing'
    with temp_dict(b'blah!', dict_format) as savename:
        d.path = savename
        d.save()
        book = pyexcel.get_book_dict(file_name=savename)
        assert list(book.items()) == MODIFIED_CONTENTS
