# Issue #28：D05AのPython比較データ（2026-10-17）

[Issue #28](https://github.com/tak063495-prog/excel-vba-rpfem/issues/28)の「HSDはOPTIMALなのに、後段の原dual復元失敗でケース全体が終了する」報告を、Pythonの元問題・変数単位・dual逆変換から調べた資料です。製品VBA・DLL・監査の閾値は変更していません。

[専用ZIP](D05A_PYTHON_ANALYSIS_20261017.zip)に、再検査コード・正規入力・メッシュ・生ベクトル・初回失敗を含む実行記録をまとめています。

## 分かったこと

- Fs=1・同じ214要素構成で、元スケールP02Aと長さ1000倍／γ1000分の1のD05Aを6条件で新規求解しました。最終6子プロセスは9.275秒で正常終了しました。
- D05Aの通常／affine経路は原物理監査で不合格。目的関数の正規化だけでは、dualが合格してもprimalの仕事量誤差が残りました。
- 正確な速度単位変換とcone単位の正スケーリングでは、元の式へ戻した仕事量誤差6.66e-16・raw stationarity6.140e-8で、原物理場と原dualの門を通過しました。
- 過去のD05A物理対も保存データから再監査し、GAP1.7206300416%を再確認しました。過去のfull dual2本は、一方の数値合格と他方の数値不合格をそのまま再現しています。

新実験は固定Fs=1の比較です。目的値約2.583は上界witnessの閾値より大きく、**Fsの根やGAP2%の完走結果ではありません**。Issueの400要素・HSD／DLL実行そのものを再現した結果でもありません。詳しくは[原因の考察](ANALYSIS_JA.md)を参照してください。

## 保存ベクトルを再検査する（新規最適化0回）

Python3.13の環境で、展開したディレクトリから実行します。出力は未作成ファイルを指定してください。

```powershell
python -m pip install -r requirements.txt
python -B -X utf8 VERIFY_MICRO.py --output ../D05A_MICRO_REPLAY_NEW.json
python -B -X utf8 VERIFY_HISTORICAL.py --output ../D05A_HISTORICAL_REPLAY_NEW.json
```

`VERIFY_MICRO.py`は新実験6条件の元q8問題を組み立て直し、生ベクトルの合格・棄却を照合します。`VERIFY_HISTORICAL.py`は過去の物理対と2本の原dualを再監査します。どちらも新しい最適化を実行しません。

## 固定点6条件を新しく求解する

既定の1求解予算は20秒、各BLAS／solverは1 threadです。producerは一式内への相対パスを記録するため、出力先にはこの一式の中の新規ディレクトリを指定します。一式外の出力先は使用できません。

```powershell
python -B -X utf8 run_fixed_point_experiment.py --seconds 20 --output results/fixed_point_new
python -B -X utf8 VERIFY_MICRO.py --run results/fixed_point_new --output ../D05A_FIXED_POINT_NEW_REAUDIT.json
```

環境によって状態や所要時間が変わるため、原監査に落ちた候補を自動的に合格へ変える処理はありません。値が得られてもFs根・2%完走と表示しません。

## 収録内容と限界

- `python_snapshot/`、`inputs/`、`SOURCE_DATA_MANIFEST.json`：新実験で固定した元ソースと入力・214要素の相似メッシュ。
- `results/fixed_point_02/`：最終6条件の全ベクトル・原監査・終了状態。`VERIFY_MICRO.py`で再検査します。
- `results/fixed_point_01/`、`reports/harness_initial_v1.py`：初回wrapperのreduced dual API誤認2件を含む原記録。数値coreを変更せずwrapperを直して全6条件を再実行しました。両回の計12求解を開示しています。
- `historical/`、`HISTORICAL_MANIFEST.json`：過去の原データ・ソースを72ファイルのSHAで固定。これまで公開98MB ZIPに無かった2本のfull vectorも今回収録しています。
- `reports/`：新実験の集計、過去対・raw dualの再監査、23箇所の原文・SHA、報告本文の保存、ソース上の原因考察。

過去の2%対は下界1.3028266247689342／上界1.3254379799492289、両1622要素です。元workerのOS終了コードが記録されなかった点はnullのまま保持し、正常cold終了時間を保証する結果にはしません。全反対端点の根符号を今回フルベクトルから再生成したという主張も含めません。
