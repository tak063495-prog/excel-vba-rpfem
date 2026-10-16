# 同一 Fs 上界 LP / 最大2回 QP 回復の隔離試作

TerminalRoot の凍結済み `RPX_AccuracyNative` から派生した1モジュールです。製品既存ソース、実行中ブック、旧証拠は変更しません。`ACCURACY_DLL_UPPER_RECENTER=0` が既定で、従来の上界・全下界の処理本文を維持します。G1R20 の UI 設定は integration が管理します。

1 にした上界のみ、同一の原 q8 / reference_total / Fs / 入力 epoch / モデル・方針・組立 context に対して、新規原 LP →正の仕事正規化→元 PrimalAudit と FieldGate を行います。不採用なら `P=.01 I, q=-.01 center` と元目的の上限 `.999998` を持つ補助 QP を解き、仕事正規化して同じ監査を行います。さらに不採用なら、その候補を中心にした QP をもう1回だけ解きます。全3回を通じて元式・元目的・Fs を変更しません。全補助双対は破棄し、根の符号に使いません。

DirectBound では、既存原 LP と原 dual/helper の全てが不採用だった場合にだけこの回復を HSD より先に試します。同じ exported original program の LP candidate が有限なら再利用し、LP の重複実行を省きます。欠落または不正な candidate の場合だけ fresh LP を行い、別ログを残します。元 dual/helper の positive sign が採用された場合は回復に入りません。回復成功後にも元監査を再実行し、既存 `original_primal` commit と Restore に合流します。回復中の入力変更・context変更・取消等の例外は Restore 後に伝播させます。公開 RecoverPrimal の直接入口は常に fresh LP 開始です。

仕事の正規化は `E_work x / b_work` が有限かつ正の場合の全変数の除算です。適用条件は、元等式右辺の非零が厳密に1行、全 cone 右辺が厳密に0であることです。正の倍率は他の同次等式と cone を保存しますが、元々不適合な場を自動的に証明へ変えません。正規化の後でも元 raw 等式・cone と物理 gate の全採否が必要です。仕事が非正または非有限の候補は採用せず、有限の元 candidate だけは次の QP の中心として使用できます。未正規化の場を返しません。

Python の `validation/G1R19_python26_upper_goal_recipe_20261016/upper_goal_recipe.py` の `normalized_candidate` と最大2回 QP の順序を移植しています。新 QP は Python の既定 `cone_scaling=False` と同じく、元の行と cap を直接使用します。DLL 内部 equilibration は既存明示設定です。Python の original LP は affine reduction / primal polish を使用するため、VBA の literal original LP と全反復・配列が同一であるとの主張はしません。配列一致の実証は同一 LP を与えた別試験が必要です。

新しい場を返すことと root 完了は別です。関数は場だけを返し、既存呼出側が元の endpoint 監査と符号付き根区間契約を満たした時だけ root を完了します。設定と予算の契約は `SETTINGS_CONTRACT.json` を参照してください。30秒 LP + 各90秒の QP 最大2回は、ケース全体が210秒以内に終わる保証ではありません。

Direct の30秒 LP +30秒 dual に加えて、QP各90秒最大2回、LP center が無い場合はさらに30秒 fresh LP という最悪270秒の solve 上限もあり得ます。ケース全体の監視予算は別です。

`build_overlay.py --initialize` で新規生成し、その後は引数なしで baseline と限定 patch / UTF8 と CP932 / manifest を照合します。`python -m unittest discover -s tests -v` は solver なしの式・拒否条件・元関数無変更の試験です。実 Excel compile、DLL 実行、P06 再現・速度・GAP 達成は、この隔離 source 単独ではまだ確認していません。
