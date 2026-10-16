"""Portable static contracts; no Excel, worksheet mutation, or numerical claim."""
from pathlib import Path
import hashlib
import json
import re

ROOT = Path(__file__).resolve().parents[1]
OLD = ROOT.parent / 'G1R20_Fix2'


def procedure(text, name):
    match = re.search(r'(?im)^(?:Public |Private )?(?:Sub|Function) ' + name + r'\b[^\n]*\n', text)
    assert match, name
    end = re.search(r'(?im)^End (?:Sub|Function)\s*$', text[match.end():])
    assert end, name
    return text[match.start():match.end()+end.end()]


def ordered(text, *needles):
    previous = -1
    for needle in needles:
        current = text.find(needle, previous+1)
        assert current > previous, needle
        previous = current


def main():
    water = (ROOT/'src/vba/RPX_Water.bas').read_text(encoding='utf8')
    seep = (ROOT/'src/vba/RPX_Seep.bas').read_text(encoding='utf8')
    ui = (ROOT/'src/vba/RPX_SeepUI.bas').read_text(encoding='utf8')
    checks = []
    for name in ['RPX_Water.bas', 'RPX_Seep.bas']:
        source = (ROOT/'src/vba'/name).read_text(encoding='utf8')
        imported = (ROOT/'import'/name).read_text(encoding='cp932')
        assert source == imported, name+' src/import mismatch'
    checks.append('utf8_cp932_source_identity')
    ordered(procedure(water, 'RPX_SourceRecordWrite'),
            'For Each part In selected: previous.Add part',
            'CustomXMLParts.Add(doc.XML)', 'SourceReadPart added',
            'previous.item(i).Delete', 'RPX_SourceRecordRead(recordNamespace')
    ordered(procedure(water, 'RPX_SourceRecordRead'),
            'parts.count = 0', 'parts.count <> 1', 'SourceReadPart')
    assert 'ProhibitDTD' in procedure(water, 'SourceDOM')
    assert 'bytes = value' in procedure(water, 'SourceBase64')
    assert 'decoded = bytes' in procedure(water, 'SourceUnbase64')
    checks.append('lossless_journal_fail_closed_ordering')
    ordered(procedure(seep, 'RPX_SeepWriteResults'),
            'RPX_WaterResultWriteBegin', 'RPX_SeepSourceBegin', 'RPX_SeepSourceAssert',
            'ws.cells.Clear', 'RPX_SeepSourceCommit(identifier)',
            'ws.Range("K9").Value2 = identifier', 'ws.Range("K2").Value2 = SConverged',
            'RPX_WaterResultWriteCommit(identifier)')
    ordered(procedure(water, 'LoadSeepResult'),
            'WaterResultWriteAssert(stamp)', 'SeepResultSourceAssert(stamp)', 'If stamp = wStamp')
    ordered(procedure(ui, 'RPX_SeepRun'),
            'RPX_SeepInvalidateResult', 'RPX_SeepSourceBegin', 'RPX_ReadInputs',
            'RPX_SeepSourceAssert', 'Select Case mode', 'RPX_WTWriteMode mode',
            'RPX_SeepSourceAssert', 'Finish:', 'RPX_SeepSourceEnd')
    checks.append('producer_and_consumer_source_checks_precede_eligibility_cache')
    context = procedure(water, 'RPX_SeepSourceContext')
    for address in ['A7:C1006', 'E7:H1006', 'A5:A1004', 'F5:F1004',
                    'A7:C206', 'E7:K206', 'P7:Q206', 'M7:N206', 'C210:C215', 'C216']:
        assert '"'+address+'"' in context, address
    for excluded in ['B5:B1004', 'J7:T1006', 'A4:D100', 'C209', 'C217', 'D7:D206', 'L7:L206', 'O7:O206', 'R7:R206']:
        assert '"'+excluded+'"' not in context, excluded
    ordered(procedure(water, 'RPX_WaterCommonInputBegin'),
            'WaterInputContext()', 'waterInputBound = True', 'WCCommonInputLevel')
    ordered(procedure(water, 'RPX_WaterCommonInputAssert'),
            'waterInputBound', 'WaterInputContext()', 'Not commonInputBound')
    checks.append('all_modes_scope_and_relevant_input_footprint')
    # The numerical algorithms/load assembly are unchanged by provenance work.
    old_seep = (OLD/'src/vba/RPX_Seep.bas').read_text(encoding='utf8')
    old_water = (OLD/'src/vba/RPX_Water.bas').read_text(encoding='utf8')
    solve = procedure(seep, 'RPX_SeepSolve').replace('        Call RPX_SeepSourceAssert\n', '')
    assert solve == procedure(old_seep, 'RPX_SeepSolve')
    assert procedure(water, 'RPX_WaterApply') == procedure(old_water, 'RPX_WaterApply')
    for name in ['BuildGrid', 'bary', 'RPX_WaterHeadAt']:
        assert procedure(water, name) == procedure(old_water, name)
    checks.append('seep_and_water_numerical_formulas_unchanged')
    print(json.dumps(dict(status='PASS',scope='portable static contract only',checks=checks,
                          sources={n:hashlib.sha256((ROOT/'src/vba'/n).read_bytes()).hexdigest()
                                   for n in ['RPX_Water.bas','RPX_Seep.bas']}),indent=2))


if __name__ == '__main__':
    main()
