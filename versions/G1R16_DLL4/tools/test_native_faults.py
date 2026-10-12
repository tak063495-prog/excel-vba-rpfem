import json,subprocess
from prepare import ROOT,REPO,write,sha

def run():
 result={}
 for name in ['order','lu','job']:
  command=ROOT/'tests'/('build_failure_test.cmd' if name=='order' else 'build_'+name+'_failure_test.cmd')
  r=subprocess.run(['cmd','/c',str(command)],cwd=REPO,capture_output=True)
  raw=r.stdout+r.stderr;(ROOT/'results'/('native_'+name+'_fault_build.log')).write_bytes(raw)
  text=raw.decode('cp932',errors='replace');messages=[x for x in text.splitlines() if x.startswith('PASS')]
  result[name]={'exit_code':r.returncode,'messages':messages,'test_source_sha256':sha(ROOT/'tests'/(name+'_failure_injection.cpp')),'source_hashes':{p.name:sha(p) for p in (ROOT/'native').iterdir() if p.suffix in ['.c','.cpp','.h','.asm']},'actual_compile_and_execute':True}
  write(ROOT/'results/native_faults.json',result);print(name,r.returncode,messages,flush=True);assert r.returncode==0 and messages
 print('NATIVE fault builds/runs PASS',flush=True)
if __name__=='__main__':run()
