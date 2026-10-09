"""Generate final report only after all 25 real native runs terminate."""
from pathlib import Path
import json,csv,hashlib,collections
from datetime import datetime,timezone
from prepare_inputs import ROOT,SOURCE,SOURCE_SHA
from report_progress import collect
from case_locations import case_folder,evidence_root

def run():
    rows=collect();assert len(rows)==25
    assert all(r['native_run_started'] for r in rows),'Not all cases entered native FEM'
    assert all(r['status'] not in ['RUNNING','PREPARING','NOT_RUN'] for r in rows),'Native runs still pending'
    costs=json.loads((ROOT/'results/solve_costs.json').read_text(encoding='utf-8'))
    counts_by_case={r['case']:r['finished_solves'] for r in costs['cases']}
    assert set(counts_by_case)=={r['case'] for r in rows}
    assert all(counts_by_case[r['case']]>0 for r in rows),'Input-only attempts are not native FEM runs'
    assert hashlib.sha256(SOURCE.read_bytes()).hexdigest()==SOURCE_SHA
    for r in rows:
        folder=case_folder(r['case'])
        assert hashlib.sha256((ROOT/'inputs'/f'{r["case"]}.json').read_bytes()).hexdigest()==r['input_sha256']
        if r.get('workbook'):
            src=json.loads((folder/'saved_sources_check.json').read_text(encoding='utf-8'));assert src['pass']
            assert hashlib.sha256((evidence_root(r['case'])/r['workbook']).read_bytes()).hexdigest()==r['workbook_sha256']
        if r.get('fs_obtained'):
            saved=json.loads((folder/'saved_output_check.json').read_text(encoding='utf-8'));assert saved['pass']
            assert r['independent_audit']['pass']
            r['saved_output_pass']=True
        r['native_evidence_folder']=str(folder)
        r['accepted_witness_q']=r.get('witness_q') if r.get('fs_obtained') else None
        r['accepted_element_counts']=r.get('element_counts') if r.get('fs_obtained') else None
    counts=dict(collections.Counter(r['status'] for r in rows))
    fs=sum(r.get('fs_obtained',False) for r in rows);targets=sum(r.get('target_met',False) for r in rows)
    summary={'completed_utc':datetime.now(timezone.utc).isoformat(),'all25_native_executed':True,'fs_obtained':fs,'target_met':targets,'counts':counts,'production_sha256':SOURCE_SHA,'results':rows}
    (ROOT/'results/final_summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
    keys=['case','status','native_worker_status','fs_obtained','target_met','target_unmet_reasons','global_incomplete_note','lower_fs','upper_fs','relative_gap','root_width','root_search_incomplete','accepted_witness_q','accepted_element_counts','solve_seconds','seconds','error_number','error','saved_output_pass','native_evidence_folder']
    with (ROOT/'FINAL_RESULTS.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.DictWriter(f,fieldnames=keys,extrasaction='ignore');w.writeheader();w.writerows(rows)
    def n(v):return f'{v:.9g}' if isinstance(v,(int,float)) else '—'
    text=['# G1R12 — 間隙水圧を除く25ケースの実Excel検証','',
          f'25ケースすべてを実Excelで実行しました。独立監査を通った上下界Fsの取得は **{fs}/25**、GAP 1%とFs根探索目標の達成は **{targets}/25** です。',
          '',f'集計: `{counts}`。終了日時UTC: {summary["completed_utc"]}。','',
          '## 検証条件','',
          '原本は`RPFEM_20261009_G1R12_P06N_IssueFix.xlsm`です。ケースごとにコピーし、TARGET=1536、ADAPT=1、CYCLES=3、最終要求QUADRATURE=8、FS_TOL=0.0001、GAP=0.01で実行しました。G1R12の既存政策・カーネル・構成則・監査・許容誤差は変更していません。粗探索は実コードのCLngにより目標538、q=4、既存のFs許容幅0.005です。最終評価では要求されたFs許容幅0.0001とq=8へ戻します。追加細分化の上限1000も維持し、実採用要素数とqは個別結果へ記録しました。カタログのM0/M1/M2同一メッシュ収束試験とは区別します。',
          '',
          'Aは関連モデル、NはDavis等価関連モデルです。D08Aは実行していません。D04AのE、D09Aのνは剛塑性モデルの入力変数に含まれないため、実際のFs計算を行いましたが、弾性係数・圧縮性の感度を検証したとは扱いません。',
          '',
          '各入力は元MDと照合し、Excel投入後に座標・材料・領域・境界を読み戻しました。多層の共有境界、底面固定・垂直側面の水平拘束、P07の右側拘束なし、D07の上載圧区間を明示しています。原本は保全し、保存ブックの製品46モジュールと文書10モジュールを照合しました。追加したBatchEvidenceは証拠保存用です。','',
          '## ケース別結果','',
          '|Case|判定|下界Fs|上界Fs|GAP [%]|採用q 下/上|要素数 下/上|実解析秒|エラー|',
          '|---|---|---:|---:|---:|---|---|---:|---|']
    for r in rows:
        gap=100*r['relative_gap'] if isinstance(r.get('relative_gap'),(int,float)) else None
        error=str(r.get('error','')).replace('|','/').replace('\n',' ')[:180]
        text.append('|'+ '|'.join([r['case'],r['status'],n(r.get('lower_fs')),n(r.get('upper_fs')),n(gap),str(r['accepted_witness_q']) if r['accepted_witness_q'] else '—',str(r['accepted_element_counts']) if r['accepted_element_counts'] else '—',n(r.get('solve_seconds')),error])+'|')
    text+=['',
          '`PASS_TARGET`は、独立した場の監査に加え、保存した各所有者の有効な根ブラケットと完了状態、GAP 1%を確認した結果です。根幅の条件は`FS_TOL × max(1, 根ブラケット中点)`で、絶対幅との直接比較ではありません。集計側の最初の絶対幅判定は修正し、生のworker判定もCSVへ残しています。',
          '',
          '`FS_AVAILABLE_PARTIAL`は、有効な上下界を取得できましたが、GAPまたは根探索の目標が未達です。区間中点だけを確定したFsとして扱いません。`FAILED`は最終Fsを取得できなかった結果です。',
          '',
          'CSVの`target_unmet_reasons`と`global_incomplete_note`が未達の理由です。`root_search_incomplete`は既存の全体フラグRFsIncompleteの保存名であり、GAP未達でもTrueになります。個別根の有効性・完了フラグとは区別します。要素数上限1000など既存の計算予算は維持したため、GAP未達ケースについては追加細分化の効果を別途検証する必要があります。',
          '',
          '全材料c=0では、安定側の下界がUNBOUNDED、上界がINFEASIBLEの証明を使うことがあります。一方、`RPX_ActualRootBracket`は両端OPTIMALだけを有効とする契約です。D01Aの個別根フラグはFalseで、この証明種別の扱いも点検対象です。保存された最良所有者の両端statusと証明場は未exportなので、これだけを唯一の原因とは断定しません。将来の修復では証明種別と厳密所有者を保存・監査してから根幅を受理し、status文字列の置換だけで合格にはしません。`results/pure_c0_root_review.json`に原ログと静的契約を記録しました。',
          '',
          '全材料c=0への対策候補として、D01Aの実メッシュでλ=1の等価な実行可能性問題をPython検証しました。下端FsではAlmostSolvedから物理監査PASSの単位荷重場を取得し、上端FsではPrimalInfeasibleの数値的証明を確認しました。元のλ最大化はNumericalErrorで、どのstatusも変更していません。[検証と実装要件](pure_c0_feasibility/DIAGNOSIS_JA.md)に記録しました。今回の候補ではP06の混在cモデルを対象外とし、VBAでの完走・速度向上も未検証です。',
          '',
          '## 失敗と対策','',
          'D07Aは`FS_LOWER_BRACKET_NOT_FOUND`で停止しました。同じ粗メッシュでは上載圧が切り替わる片方の端点(2,10)に2要素しか接続せず、内部面力連続と外周面力条件が荷重倍率を0に制限することを確認しました。独立したPythonの線形矛盾証明と局所分割候補、元の物理式による50桁の監査を[診断書](surcharge_diagnosis/DIAGNOSIS_JA.md)へまとめています。候補は固定Fs=1の実行可能性を確認したもので、VBAでの修復や全Fs探索の合格へ加算していません。',
          '',
          'D03Nは最終上界の求解中に`HSD_DD_PRODUCT_UNDERFLOW`で停止しました。直前の候補は参照荷重の仕事誤差が1e-7を超えて棄却されていました。積が0になるガードと一致しますが、実際の演算値と棄却場は未保存です。8種類の演算試験と、別ケースの保存場による14件の正規化・棄却試験をPythonで実施しました。候補場・演算値の取得、仕事の高精度評価と正規化、指数を保つ演算、物理・数値の再検査の順に進める対策を[診断書](underflow_diagnosis/DIAGNOSIS_JA.md)へ記録しました。修復・本ケースの完走は未確認です。',
          '',
          'D03Nの物理入力を別の実メッシュへ設定したPython試験でも、速度係数が約2.25e7へ増え、仕事の監査が失敗しました。速度場を小さくする選択は、そのFsでは全てFAIL、別のFs=1.30では監査済みの小さな場を取得できました。[6計算の肯定・否定結果](pure_c0_upper_selection/DIAGNOSIS_JA.md)を記録しています。停止時の実メッシュや正確なFsの再生ではなく、候補をそのままVBA採用できる結果とは扱いません。',
          '',
          '最初の準備段階で発生した実行時エラー5「材料を入力してください」は、Python COMの範囲指定による投入失敗でした。明示したセル範囲への投入と全表の読み戻し検査へ修正し、25件の本試験は修正後の入力で実行しました。準備の失敗・待ち行列の予約・小規模試験を25ケースの実績へ加算していません。',
          '',
          '## 証拠と再現','',
          '[FINAL_RESULTS.csv](FINAL_RESULTS.csv)と`results/final_summary.json`が集計です。各ケースの原ログ、入力JSON、保存場、結果ブック、保存表とソースの検査を収録します。Python独立監査は幾何・材料別面積・重量・境界荷重・釣合い・降伏・速度流れ条件・仕事・端点方向を確認し、下界は元単位でも既存の1e-7許容差を確認しました。静的ソース照合、テンプレートのVBEコンパイル、小規模の証拠保存試験、25件の実解析を区別しています。',
          '',
          '複数の独立したExcelを最大13プロセスで実行しました。並行実行の計算時間を単独実行の高速化率とは扱いません。DIAGNOSTIC_MODE=1で実行し、所要時間には診断ログ・進捗保存の負荷も含まれます。ケースごとの実解析秒とログの求解時間内訳を記録しています。',
          '',
          'P02AとD04A・D09Aの上下保存場はバイト単位で一致しました。E・νを使わない同じ有効入力での再現性確認です。D06Aの平行移動によるFs差は約1e-12、D05Aの相似変換による相対差は約0.01%で、監査済み区間が重なりました。D05Aは998要素、基準は1000要素なので、同一メッシュの厳密差分とは扱いません。詳細は`results/invariants.json`に記録しています。',
          '',f'原本SHA-256: `{SOURCE_SHA}`。原本一致確認: PASS。','']
    (ROOT/'FINAL_REPORT.md').write_text('\n'.join(text),encoding='utf-8',newline='\n')
    print(json.dumps({k:summary[k] for k in ['all25_native_executed','fs_obtained','target_met','counts']},ensure_ascii=False))

if __name__=='__main__':run()
