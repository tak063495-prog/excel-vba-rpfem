from pathlib import Path
import csv,collections,json,hashlib
R=Path(__file__).resolve().parent
paths=[Path('C:/Users/link_/Desktop/RPFEM_logs/20261006_183452_26544_analysis.tsv'),Path('C:/Users/link_/Desktop/RPFEM_logs/20261007_054059_19848_analysis.tsv')]
def detail(s):
    return dict(p.split('=',1) for p in s.split(';') if '=' in p)
def load(p):
    raw=p.read_bytes();enc='utf-16' if raw[:2] in [b'\xff\xfe',b'\xfe\xff'] else 'utf-8-sig'
    rows=list(csv.DictReader(raw.decode(enc).splitlines(),delimiter='\t'))
    def events(e):return [r for r in rows if r['event']==e]
    end=events('run_end')[-1];adapt=detail(events('adapt_end')[-1]['detail'])
    solves=events('solve_end');audits=events('audit')
    phases={detail(r['detail'])['phase']:detail(r['detail']) for r in events('run_phase')}
    buckets={detail(r['detail'])['bucket']:detail(r['detail']) for r in events('hopt0') if detail(r['detail']).get('scope')=='run'}
    groups={}
    for r in solves:
        g=groups.setdefault(r['candidate'],{'calls':0,'solve_seconds':0,'lower_seconds':0,'upper_seconds':0,'iteration_sum':0})
        d=detail(r['detail']);t=float(d['wall_seconds']);g['calls']+=1;g['solve_seconds']+=t;g[r['mode']+'_seconds']+=t;g['iteration_sum']+=int(d['last_iteration_index'])
    coarse=[(r['mode'],r['strength_factor'],{k:v for k,v in detail(r['detail']).items() if k in ['status','objective','last_iteration_index','primal','dual','gap']}) for r in solves if r['candidate']=='coarse_0']
    return {'path':str(p),'sha256':hashlib.sha256(raw).hexdigest(),'elapsed_seconds':float(end['elapsed_s']),'outcome':detail(end['detail']),'adapt_end':adapt,'solve_statuses':dict(collections.Counter(detail(r['detail']).get('status') for r in solves)),
            'audit_statuses':dict(collections.Counter(detail(r['detail']).get('status') for r in audits)),
            'audit_category_false':sum(detail(r['detail']).get('category_pass')=='False' for r in audits),'error_origins':events('error_origin'),'cycles':[detail(r['detail']) for r in events('adapt_cycle')],
            'groups':groups,'phases':phases,'buckets':buckets,'coarse_points':coarse,'settings':detail(events('settings')[0]['detail']),
            'material_records':sorted(set(r['detail'] for r in events('hsd_material') if r['candidate']=='coarse_0')),'final_audits':[detail(r['detail']) for r in audits if r['candidate']=='final']}
old,new=map(load,paths)
delta={'elapsed_increase_percent':100*(new['elapsed_seconds']/old['elapsed_seconds']-1),'gap_reduction_percent':100*(1-float(new['adapt_end']['gap'])/float(old['adapt_end']['gap'])),
       'lower_change_percent':100*(float(new['adapt_end']['lower'])/float(old['adapt_end']['lower'])-1),'upper_change_percent':100*(float(new['adapt_end']['upper'])/float(old['adapt_end']['upper'])-1),
       'initial_18_coarse_solve_results_exact_printed_match':old['coarse_points']==new['coarse_points'],'logged_settings_equal':old['settings']==new['settings'],'logged_materials_equal':old['material_records']==new['material_records']}
combined_lower=float(old['adapt_end']['lower']);combined_upper=float(new['adapt_end']['upper']);delta['conditional_cross_run_gap']=2*(combined_upper-combined_lower)/(combined_upper+combined_lower)
(R/'results/full_run_comparison.json').write_text(json.dumps({'old':old,'new':new,'comparison':delta,'scope':'LOG_REVIEW_ONLY_NO_SAVED_FIELD_OR_COMPLETE_INPUT_IDENTITY_REAUDIT'},ensure_ascii=False,indent=2),encoding='utf-8')
def duration(s):
    s=round(s);return f'{s//3600}時間{s%3600//60}分{s%60}秒'
table='\n'.join(f"|{k}|{v['calls']}|{v['solve_seconds']/3600:.2f}時間|{v['lower_seconds']/3600:.2f}時間|{v['upper_seconds']/3600:.2f}時間|" for k,v in new['groups'].items())
report=f'''# G1R10 P06N全解析ログの評価

解析は完走し、全61求解がOPTIMALです。監査記録は{new['audit_statuses']}、category_pass=Falseは{new['audit_category_false']}件、error_originは{len(new['error_origins'])}件でした。ただし最終状態はPARTIAL_CERTIFIED / ADAPT_GAP_TARGET_NOT_METであり、GAP=1%未達です。

## 前回との比較

|指標|G1R9|G1R10|
|---|---:|---:|
|全時間|{duration(old['elapsed_seconds'])}|{duration(new['elapsed_seconds'])}|
|Fs下界|{old['adapt_end']['lower']}|{new['adapt_end']['lower']}|
|Fs上界|{old['adapt_end']['upper']}|{new['adapt_end']['upper']}|
|GAP|{100*float(old['adapt_end']['gap']):.2f}%|{100*float(new['adapt_end']['gap']):.2f}%|
|最終要素数|{old['adapt_end']['elements']}|{new['adapt_end']['elements']}|
|求解数|61|61|

GAPは相対{delta['gap_reduction_percent']:.1f}%減りましたが、時間は{delta['elapsed_increase_percent']:.1f}%増えました。下界は{abs(delta['lower_change_percent']):.2f}%低下、上界は{abs(delta['upper_change_percent']):.2f}%低下です。上界の改善を確認しましたが、高速化の成果としては不合格です。上下界幅の減少だけで、下界も改善したとは評価しません。

## 原因の切り分け

設定・材料ログが一致し、初期601要素の18求解について、Fs・目的値・残差・最適性指標・反復数が印字値で全て一致しました。初期段階の時間も{old['groups']['coarse_0']['solve_seconds']:.1f}秒と{new['groups']['coarse_0']['solve_seconds']:.1f}秒で近く、追加の費用は主として適応経路の差です。完全な入力snapshot・保存場の同一性を再監査した証拠とは区別します。

G1R10は601→902→1348→1536要素へ直接細分化しました。前回は601要素で再配置してから781要素へ細分化しています。前回の良い再配置済みメッシュを出発点とせず、上界の速度場による指標で要素を増やしたため、下界の応力場に有利な要素配置を十分に得られなかった可能性があります。これはメッシュ・保存場の追加比較で確認すべき仮説です。後半ほど各求解が重くなりました。新経路内では下界・上界が各段階で単調に改善しており、ログから数値核の不具合を示す証拠は見つかっていません。異なる2経路のメッシュは入れ子ではないため、要素が多い方の下界が高くなる保証はありません。

重要な事前検証との差は、Pythonで検証した994要素が、前回の再配置済み781要素から細分化したメッシュだったことです。今回実運転した1536要素は、初期601要素から直接細分化した別メッシュです。固定点のPython/VBA検証が合格でも、今回の適応順序の性能を保証するものではありませんでした。今回の全ログでこの差を確認できました。

|段階|求解数|求解時間合計|下界|上界|
|---|---:|---:|---:|---:|
{table}

最終8求解だけで約{new['groups']['final']['solve_seconds']/3600:.2f}時間を費やしています。ただし最終精度確認であり、一律に省略する対策は取りません。numeric_ldlは{float(new['phases']['numeric_ldl']['seconds'])/3600:.2f}時間、DD演算の計測bucketは{float(new['buckets']['dd_operator']['exclusive_seconds'])/3600:.2f}時間です。包含関係のある計測値を二重加算しません。メッシュ操作自体は{new['phases']['mesh']['seconds']}秒で、メッシュ処理のループだけを高速化しても総時間はほとんど改善しません。

## 次の改訂方針

1. **従来の再配置・良いメッシュを保持してから追加細分化する。** 初期601要素から直接1536へ進む経路を既定とする判断を見直す。再配置済み781要素→約1000要素の経路をPythonで、Fs=0.655/0.66下界とFs=0.77上界の保存場・原式監査まで再検証する。
2. **下界も考慮して配分する。** 上界の運動学指標に加え、下界の降伏接近度や応力変化・釣合い補正の指標を比較する。重みを試験で決め、単なる全体増加ではなく弱層・境界・先端のどこを増やすと両側が改善するか検証する。既知数値失敗の復帰、未知例外・監査失敗・中止の停止は維持する。
3. **監査済みの最良下界を候補間・同一入力の再実行で保持する。** 前回の0.630355を新メッシュの0.613174で単純に捨てない。ただしFsの数値だけをコピーしない。入力・荷重・Davis政策の同一性と、旧場の原式監査を確認する。条件を満たせば旧下界と新上界の区間は0.630355～0.711569、GAP約{100*delta['conditional_cross_run_gap']:.2f}%となる。現時点では条件付き提案であり、新しい認証結果としては扱わない。
4. **不要に遠いFs試行を削減する。** 有効な上界があるのに下界でFs=1.0864、1.2133まで拡張し、約921秒・1143秒の求解をしています。多層ではUPPER_TRIALがmaterialsで無効化されているため、そのゲートを無条件に解除せず、監査済み同一モデルの上界を安全なブラケット上限に使う方法をPythonで検証する。数値失敗を物理的不可能と読み替えない。
5. **最終確認の費用を抑える。** 同じメッシュでのcoarse→final再組立・根幅確認について、安全な端点再利用と必要な追加点だけを評価する方法を検証する。q=8、FS_TOL、GAP、物理監査は維持する。

まず再配置済み約1000要素でP06Nを検証し、今回の1536要素より下界が高く、時間が短いかを確認する方針を優先します。TARGETを小さくするだけでは経路の問題が残るため、単純な設定変更だけで改善を保証しません。

## 証拠と制約

比較元: `{paths[0].name}`。今回: `{paths[1].name}`。
集計コード: evaluate_full_run.py。機械可読結果: results/full_run_comparison.json。
今回はログの評価であり、VBA改訂・全ケース再実行・保存場の新たな50桁再監査は行っていません。以前の固定点検証と今回の全解析結果を混同しません。
'''
(R/'FULL_RUN_REVIEW.md').write_text(report,encoding='utf-8')
print(json.dumps({'comparison':delta,'solve_statuses':new['solve_statuses'],'audit_statuses':new['audit_statuses'],'audit_false':new['audit_category_false'],'error_origins':len(new['error_origins']),'groups':new['groups']},ensure_ascii=False,indent=2))
