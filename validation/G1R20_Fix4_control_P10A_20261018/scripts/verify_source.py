import re,struct
def canonical(text):
    text='\n'.join(s for s in text.splitlines() if not s.startswith('Attribute VB_'))
    result=[]
    for match in re.finditer(r'"(?:""|[^"])*"|\x27[^\n]*|[A-Za-z_][A-Za-z_0-9]*|(?:\d+\.\d*|\.\d+|\d+)(?:[eE][-+]?\d+)?[#@!%&]?|[^\s]',text):
        token=match.group()
        if token.startswith("'"):continue
        if token.startswith('"'):result.append(token);continue
        if re.fullmatch(r'(?:\d+\.\d*|\.\d+|\d+)(?:[eE][-+]?\d+)?[#@!%&]?',token) and ('.' in token or 'e' in token.lower() or token.endswith('#')):
            try:token='DOUBLE:'+struct.pack('<d',float(token.rstrip('#'))).hex()
            except ValueError:pass
        result.append(token.lower())
    return result
