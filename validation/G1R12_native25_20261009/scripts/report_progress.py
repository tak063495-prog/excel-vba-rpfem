from pathlib import Path
import json,collections,csv,hashlib
from datetime import datetime,timezone
from prepare_inputs import ROOT,SOURCE,SOURCE_SHA
from case_locations import case_folder,evidence_root
from grade_results import grade
def collect():
    rows=[]
    for p in sorted((ROOT/'inputs').glob('*.json')):
        s=case_folder(p.stem)/'status.json'
        row=grade(json.loads(s.read_text(encoding='utf-8'))) if s.exists() else {'case':p.stem,'status':'NOT_RUN'}
        logs=list((case_folder(p.stem)/'RPFEM_logs').glob('*_analysis.tsv'))
        row['native_run_started']=bool(logs)
        row['evidence_root']=str(evidence_root(p.stem).relative_to(ROOT))
        rows.append(row)
    return rows
def report():
    rows=collect();counts=dict(collections.Counter(r['status'] for r in rows))
    completed=[r for r in rows if r['status'] not in ['RUNNING','PREPARING','NOT_RUN']]
    started=sum(r['native_run_started'] for r in rows);obtained=sum(r.get('fs_obtained',False) for r in rows)
    stamp=datetime.now(timezone.utc).isoformat()
    lines=['# G1R12 — 実Excel 25ケースのFs検証', '',f'更新: {stamp}。間隙水圧D08Aを除く25ケース。', '',f'実解析開始 {started}/25、終了 {len(completed)}/25、独立監査を通過したFs取得 {obtained}/25。', '',
        '## 実行条件','',
        '- 原本: RPFEM_20261009_G1R12_P06N_IssueFix.xlsm。変更せず、ケースごとにコピーを作成。',
        '- TARGET=1536、ADAPT=1、CYCLES=3、最終要求QUADRATURE=8、FS_TOL=0.0001、GAP=0.01。G1R12の既存政策・カーネル・監査・許容誤差を維持。',
        '- 粗探索のq=4など実装内の既存段階は維持し、実際に採用した上下の場のqは個別結果へ記録。カタログM0/M1/M2のメッシュ収束試験とは区別。',
        '- AはASSOCIATED、NはDAVIS_EQUIVALENT。非関連弾塑性FEMそのものの検証とは区別。',
        '- 各入力は元MDとJSONを照合。多層の共有境界、底面固定・側面水平拘束、P07の右側拘束なし、D07のx=2～8の上載圧20kPaを明示。',
        '- D04のE、D09のνは剛塑性の入力変数に含まれない。実際にFsを求解するが、弾性・圧縮性の感度を検証したとは扱わない。','',
        '## 結果台帳','',
        '|Case|状態|実解析|下界Fs|上界Fs|区間幅[%]|目標達成|秒|エラー|',
        '|---|---|---|---:|---:|---:|---|---:|---|']
    def num(x):return f'{x:.9g}' if isinstance(x,(int,float)) else ''
    for r in rows:
        gap=r.get('relative_gap');err=str(r.get('error','')).replace('|','/').replace('\n',' ')[:160]
        lines.append('|'+ '|'.join([r['case'],r['status'],'開始済' if r['native_run_started'] else '未実施',num(r.get('lower_fs')),num(r.get('upper_fs')),num(100*gap) if isinstance(gap,(int,float)) else '',str(r.get('target_met','')),num(r.get('seconds')),err])+'|')
    lines+=['','FS_AVAILABLE_PARTIALは、上下のFsと独立に監査した場を取得できたものの、GAP1%またはFs根探索条件が未達の状態です。PASS_TARGETと分けています。RUNNINGやNOT_RUNを合格へ転記しません。Fs根幅の判定は`FS_TOL × max(1, 根ブラケット中点)`です。絶対根幅をFS_TOLへ直接比較した初期の集計判定は、製品コードの定義に合わせて修正しました。生のworker記録と改訂判定は区別しています。', '',
        '## 証拠と再現','',
        '`inputs/`は設定と明示した境界条件、`cases/<ID>/`と`queued_continuation/cases/<ID>/`は実解析のXLSM・ログ・場・状態、`results/`は集計です。`queue_transfer.json`に記載した20ケースでは後者が実証拠の場所です。元の待ち行列への予約は解析失敗として数えません。`prepare_inputs.py`、`run_native25.py`、`continue_queue.py`、`audit_batch.py`で再現します。BatchEvidenceは読み取り・場の保存用の追加モジュールで、製品46モジュールを変更していません。', '',
        '最初の一括入力ではPython COMの範囲指定により材料が投入されず、実行時エラー5が発生しました。修正後は全入力範囲を読み戻して一致確認し、実定義の点・領域・面数・面積・政策も検査しています。準備段階の失敗は解析実行やFsの結果へ数えていません。', '',
        f'原本SHA-256: `{SOURCE_SHA}`。現在も一致: {hashlib.sha256(SOURCE.read_bytes()).hexdigest()==SOURCE_SHA}。','']
    (ROOT/'REPORT.md').write_text('\n'.join(lines),encoding='utf-8',newline='\n')
    (ROOT/'results/live_summary.json').write_text(json.dumps({'updated_utc':stamp,'started':started,'completed':len(completed),'fs_obtained':obtained,'counts':counts},ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'started':started,'completed':len(completed),'fs_obtained':obtained,'counts':counts},ensure_ascii=False))
if __name__=='__main__':report()
