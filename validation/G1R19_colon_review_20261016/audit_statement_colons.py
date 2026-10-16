"""Read-only scan of the merged G1R19/2pct VBA source; never starts Excel.

The runtime observation that a leading ``ZeroArgSub:`` is parsed as a label
comes from the separately owned Excel compile/micro work. This scan identifies
that exact ambiguity. It is not a VBA parser or a proof about all colon syntax.
"""
from pathlib import Path
import argparse
import ast
import hashlib
import json
import re

REPO = Path(__file__).resolve().parents[2]
DECL = re.compile(r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?Sub\s+(\w+)\s*(?:\((.*)\))?\s*$', re.I)
LEADING = re.compile(r'^\s*([A-Za-z_]\w*)\s*:(.*)$')


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def logical_lines(text):
    pending = ''
    for line in text.splitlines():
        pending += line
        if pending.rstrip().endswith(' _'):
            pending = pending.rstrip()[:-1]
            continue
        yield pending
        pending = ''
    if pending:
        yield pending


def split_args(text):
    # Declarations in these sources use ordinary Optional/ParamArray arguments.
    parts, depth, quoted, first, i = [], 0, False, 0, 0
    while i < len(text):
        char = text[i]
        if char == '"':
            if quoted and i+1 < len(text) and text[i+1] == '"':
                i += 2
                continue
            quoted = not quoted
        elif not quoted:
            if char == '(':
                depth += 1
            elif char == ')':
                depth -= 1
            elif char == ',' and depth == 0:
                parts.append(text[first:i].strip())
                first = i+1
        i += 1
    parts.append(text[first:].strip())
    return [part for part in parts if part]


def source_files():
    files = {p.stem: p for p in (REPO/'versions/G1R19/src/vba').glob('*.bas')}
    files.update({p.stem: p for p in (REPO/'versions/G1R19_2pct/src/vba').glob('*.bas')})
    if len(files) != 62:
        raise ValueError('Merged source must contain exactly 62 modules')
    return dict(sorted(files.items()))


def code_key_checks():
    path = REPO/'versions/G1R19_2pct/tools/build_candidate_book.py'
    parsed = ast.parse(path.read_text(encoding='utf8'))
    function = next(n for n in parsed.body if isinstance(n, ast.FunctionDef) and n.name == 'code_key')
    namespace = {'hashlib': hashlib}
    exec(compile(ast.Module(body=[function], type_ignores=[]), str(path), 'exec'), namespace)
    key = namespace['code_key']
    comparisons = {
        'identifier_case_ignored': ('Call CacheClear', 'call cacheclear', True),
        'string_case_protected': ('a = "Input_ID"', 'a = "input_id"', False),
        'escaped_quote_protected': ('a = "x""Y"', 'a = "x""y"', False),
        'apostrophe_comment_protected': ("a=1 'Clear", "a=1 'clear", False),
        'numeric_value_protected': ('a=1.000003', 'a=1.000001', False),
        'call_vs_label_protected': ('Call Clear: a=1', 'Clear: a=1', False),
        'newline_format_ignored': ('Call Clear\r\n', 'Call Clear\n', True),
        'rem_comment_case_not_protected': ('Rem Input_ID', 'Rem input_id', True),
    }
    result = {}
    for name, (left, right, expected_equal) in comparisons.items():
        equal = key(left) == key(right)
        if equal != expected_equal:
            raise AssertionError(name)
        result[name] = dict(equal=equal, expected_equal=expected_equal)
    return dict(builder_path=path.relative_to(REPO).as_posix(), builder_sha256=sha(path), checks=result)


def audit():
    files = source_files()
    texts = {name: path.read_text(encoding='utf8') for name, path in files.items()}
    no_arg, optional = {}, {}
    for module, text in texts.items():
        for line in logical_lines(text):
            match = DECL.match(line)
            if not match:
                continue
            name, arguments = match.group(1), split_args(match.group(2) or '')
            if not arguments:
                no_arg[name.lower()] = (module, name)
            elif all(re.match(r'^(Optional|ParamArray)\b', part, re.I) for part in arguments):
                optional[name.lower()] = (module, name)
    callable_without_args = no_arg | optional
    findings = []
    for module, text in texts.items():
        for number, line in enumerate(text.splitlines(), 1):
            match = LEADING.match(line)
            if not match or match.group(1).lower() not in callable_without_args:
                continue
            name = match.group(1)
            declaration_module, declaration_name = callable_without_args[name.lower()]
            # A genuinely intentional label may share a procedure name, so
            # record branch targets separately instead of automatically editing.
            references = [i for i, value in enumerate(text.splitlines(), 1)
                          if re.search(r'\b(?:GoTo|GoSub|Resume)\s+'+re.escape(name)+r'\b', value, re.I)]
            findings.append(dict(path=files[module].relative_to(REPO).as_posix(), line=number,
                                 text=line.strip(), procedure=declaration_name,
                                 declaration_module=declaration_module,
                                 optional_arguments=name.lower() in optional,
                                 branch_reference_lines=references))
    return dict(status='READ_ONLY_COMPLETE', solver_calls=0, excel_started=False,
                source_modules=len(files), zero_argument_sub_names=len(no_arg),
                optional_only_sub_names=len(optional), findings=findings,
                source_sha256={p.relative_to(REPO).as_posix():sha(p) for p in files.values()},
                code_key=code_key_checks())


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    result = audit()
    if args.output:
        args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf8')
    print(json.dumps({k:result[k] for k in ['status','source_modules','zero_argument_sub_names',
                                         'optional_only_sub_names','findings']}, ensure_ascii=False))
