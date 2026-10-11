from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
for name in ['RPX_LDL','RPX_SchurLDL','RPX_SeepUI','RPX_Order','RPX_ColumnOrder','RPX_Native']:
    text=(ROOT/'src/vba'/f'{name}.bas').read_text(encoding='utf-8')
    data=text.replace('\n','\r\n').encode('cp932',errors='strict')
    assert data.decode('cp932').replace('\r\n','\n')==text
    (ROOT/'import'/f'{name}.bas').write_bytes(data)
print('PASS 6 CP932/CRLF modules, lossless')
