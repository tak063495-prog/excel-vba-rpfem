import sys,platform,subprocess,importlib.metadata,json,datetime
from prepare import ROOT,REPO,SOURCE,DLL,sha,write
hardware=subprocess.check_output(['powershell.exe','-NoProfile','-Command',
    "@{cpu=Get-CimInstance Win32_Processor | Select-Object Name,NumberOfCores,NumberOfLogicalProcessors;os=Get-CimInstance Win32_OperatingSystem | Select-Object Caption,Version,TotalVisibleMemorySize;excel_processes=Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object Id,StartTime} | ConvertTo-Json -Depth 4"],text=True,encoding='utf-8',errors='replace')
data={'captured_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'platform':platform.platform(),
    'python':sys.version,'packages':{name:importlib.metadata.version(name) for name in ['pywin32','numpy','scipy','oletools']},
    'hardware':json.loads(hardware),'source_commit':(ROOT/'SOURCE_COMMIT.txt').read_text().strip(),
    'source_sha256':sha(SOURCE),'dll_sha256':sha(DLL),'execution':'Sequential native Excel, one analysis at a time',
    'macro_security':'Recorded without modification in each case status',
    'VBA_measurement':'No new VBA full timing per prior user instruction',
    'timing_boundary':'Main: run_end.elapsed_s; additionally COM RPX_Run, water preparation and initial RPX_CreateMesh measured separately; saving/audits/Excel startup excluded'}
write(ROOT/'results/environment.json',data);print('Environment recorded')
