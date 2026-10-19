import json,sys,glob,os
for f in sorted(glob.glob(sys.argv[1]+'/*/log.jsonl')):
    print('==',os.path.basename(os.path.dirname(f)))
    for l in open(f):
        e=json.loads(l);ev=e['event']
        if ev=='step':
            print(f" {e['side'][0]}{e['round']:>2} r{e['result']} el={e.get('elements','-')} fs={e.get('fs','-') if not isinstance(e.get('fs'),float) else round(e['fs'],4)} gap={e['gap']*100:.2f}% solves={e.get('solves','-')} max={e.get('max_solve_s','-')} cones={e.get('cones','-')} t={e['t']} {e.get('reason','')}")
        elif ev=='end': print(' END',{k:(round(v,5) if isinstance(v,float) else v) for k,v in e.items() if k!='event'})
