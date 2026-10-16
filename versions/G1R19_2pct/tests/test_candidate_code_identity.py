"""Allow Excel identifier casing without accepting a changed program."""
from pathlib import Path
import importlib.util
import unittest

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('book_builder', ROOT/'tools/build_candidate_book.py')
builder = importlib.util.module_from_spec(spec)
spec.loader.exec_module(builder)


class SerializedCodeIdentityTests(unittest.TestCase):
    def test_identifier_case_is_ignored(self):
        self.assertEqual(builder.code_key('Call RPX_Clear\r\nx = 1#\r\n'),
                         builder.code_key('call rpx_clear\nx = 1#\n'))

    def test_string_and_escaped_quote_case_is_preserved(self):
        self.assertNotEqual(builder.code_key('x = "A""B"'), builder.code_key('x = "a""B"'))

    def test_comments_are_preserved(self):
        self.assertNotEqual(builder.code_key("Call X ' Proof"), builder.code_key("Call X ' proof"))

    def test_numeric_value_and_actual_calls_are_preserved(self):
        self.assertNotEqual(builder.code_key('x = 0.02'), builder.code_key('x = 0.2'))
        self.assertNotEqual(builder.code_key('Call Clear'), builder.code_key('Clear:'))


if __name__ == '__main__':
    unittest.main()
