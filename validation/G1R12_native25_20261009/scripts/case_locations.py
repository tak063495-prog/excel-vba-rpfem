"""Resolve real evidence roots; scheduler reservations are not FEM results."""
import json
from prepare_inputs import ROOT

def evidence_root(case):
    transfer=ROOT/'queue_transfer.json'
    if transfer.exists() and case in json.loads(transfer.read_text(encoding='utf-8'))['cases']:
        return ROOT/'queued_continuation'
    return ROOT

def case_folder(case):return evidence_root(case)/'cases'/case
