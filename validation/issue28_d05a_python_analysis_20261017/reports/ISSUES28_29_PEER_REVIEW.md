# Issues #28・#29のPython証拠：独立レビュー

対象起点：main `1ab383bbc1764249a4320ca28ed3ac6682c8fbac`、確認日2026-10-17（日本時間）。製品ソース・Git・Excel・solverは変更／実行していない。今回のpeer作業はコード・保存記録・生ベクトルの構成・SHA照合・構文検査であり、新たな求解や第三の実行再監査を行ったという記録ではない。対応するJSONに実際に読んだファイルのSHA・確認範囲を残す。

## 判定

高確信の数学的な誤り、製品成功範囲の過大表示、失敗候補の隠蔽は見つからなかった。公開にあたって、下記の限界を現在の説明どおり保持することが前提である。現在のFix4問題が修正済み、通常cold計算が2%以内、という結論にはならない。

## Issue #29：P06Nの保存対

- `VERIFY.py` は82必須ファイルのサイズ・SHAを照合し、明示した読取用対応表で旧絶対パスを解決する。全82ファイルのSHAをpeerでも再計算し、不一致0を確認した。MANIFESTのSHA `09c7b5324edb1bc17030953b326c8e9d8b9881f748aaa5d75fb33f02b36ab880` は独立展開の実行結果と一致する。
- 正規入力を元builderのcanonical JSONハッシュ・方針で検査し、下5966／上8739メッシュを再構築、メッシュID・物理モデルID・q8・Fs・自己根記録の端点／幅／符号を照合する。保存された応力・速度は新しく組み立てた原q8問題と独立物理監査で検査する。
- 上界反対端点では元のfull dualを補正前後とも再監査する。補正前のstationarity棄却、補正後の合格と原cone変数のバイト不変を確認する手順である。helperのdualを原dualと取り違える実装ではない。
- 下界反対端点のfull vector未保存という制限は、コード・出力・README・MANIFESTに明記されている。記録された自己根の同一性・符号・幅を確認することと、その反対端点を独立に再生成することを区別している。
- `clarabel.DefaultSolver` を監査前に禁止し、保存場の原式再監査というoptimizer0の範囲を明示している。今回のpeerでは関連ソース内のoptimizer入口も確認し、呼出しはこの属性を介する構成だった。
- README／ANALYSISは、歴史的な複合実行と通常cold経路、既知形状の利用と全形状生成未解決、Issue本文・コメントによる実測報告と未取得Triela生ログを分けている。SpeedとG20改善の代入抽出は静的証拠であり、実行機の設定全体の証拠とは扱っていない。
- 50個の当時の数値ソース＋7wrapper、計57Pythonファイルの必須ファイル内訳を確認した。VERIFYの構文検査は合格。専用ZIPはこのレビュー資料を追加する対象にせず、凍結した一式を維持する。

## Issue #28：6回の固定点実験

`run_fixed_point_experiment.py`、`VERIFY_MICRO.py`、24個のSOURCE_DATA_MANIFEST必須ファイル、`results/fixed_point_02`の6個の結果・全体記録、保存した変換前後ベクトルのキー・対応を確認した。マニフェスト24ファイルのサイズ・SHA不一致0、3個の実行／監査スクリプトの構文不具合0を確認した。

変数変換の数式は正しい。`x = k*x'`、正の行倍率R、目的倍率aについて、`A' = R*A`、`rhs' = R*rhs/k`、`c' = a*k*c` とした後、原dualは `R*z'/(a*k)` で戻す。目的値は変換後値をaで割る。`cone_row_scales` は各SOCブロック内で共通の正倍率であり、coneを変更しない。元問題のdigestを求解前後で照合し、元物理門・絶対dual門で再監査する構成だった。

P02A→D05Aでは、三角形接続・材料IDが一致し、座標×1000・gamma÷1000・c/phi/psi不変を記録している。原Eほぼ同一、bとstrain Gは1/1000、jump Gは同一、costは1000倍という全係数比較とも整合する。対応するprimal `x/1000`、等式・jump dual×1000、strain dual×1e6の対応は正しい。相似則は入力や座標の値を変更した新モデルを勝手に採用する操作ではなく、元の正規D05Aに戻して残差と物理場を再確認する検査として使われている。

6試行を読み直した分類：

| 試行 | solver status | 原物理場 | 逆側原dual符号 | この結果の意味 |
|---|---|---|---|---|
| P02A none／affine | Solved | 合格 | 合格 | Fs=1の固定点。upper witnessは不合格（値>1） |
| D05A none | AlmostSolved | 不合格 | 不合格 | dual目的値約0.224で分離値未達。statusを採用していない |
| D05A affine | AlmostSolved | 不合格 | 不合格 | dual目的値約0.115、stationarityも門外 |
| D05A normalized_objective | AlmostSolved | 不合格 | 合格 | 原単位へ戻したdualだけ離散符号に合格。物理場を合格にしていない |
| D05A velocity_units | AlmostSolved | 合格 | 合格 | 原物理／dual門に合格。ただし値約2.583なのでupper witnessではない |

全てFs=1であり、FOS上下限・根・GAP2%達成を新しく示す実験ではない。Issue機の400要素・HSD／DLL helperを再現したものでもない。原dualに合格しても、それを連続体のprimal boundと扱わないという説明は適切である。

新しいREADME／ANALYSISの6行の丸め表も各result JSONと照合し、記載の物理合否・dual guide合否・値・時間範囲に相違はなかった。全6 NPZのFs/mode/q、primal／dualの次元と有限性、変換経路2件の保存ベクトル間の逆変換を別途読み取りで照合し、保存された数値の対応が一致した。これは原問題の再組立監査を第三に実行したという意味ではない。

### 正規領域面積の丸め処理

loaderは正規recordの領域面積と実座標からの面積和を先に比較し、浮動小数点の加算規模に応じたguard以内であることを要求した後、元builderの固定1e-9 assertionにだけ同じ実面積和を渡している。確認した誤差はP02A `5.684341886080802e-14`、D05A `5.960464477539063e-8`。この局所的なconstruction recordの置換は `region_areas` だけである。

元builderでは `region_areas` は整合性assertに使い、物理モデルIDは元 `source_case` から作る。座標・三角形・材料ID・材料値・物理入力は元データのままなので、このguard対応は要素演算子や原監査閾値を変えていない。原プログラムdigest・相似係数比較・元の物理／rawdual再監査も保存されている。一般の悪条件幾何を許可する製品仕様の変更として説明してはならず、この214要素の読み取り用対処という現在の範囲を保持する。

### 初回harness失敗の保持

`results/fixed_point_01` は `EXPERIMENT_HAS_ERRORS` のまま保存され、P02A/D05Aのaffine2件はreduced dualを元dualと誤認したshape errorを保持している。初回コード `reports/harness_initial_v1.py` とtracebackがある。最終コードは既存coreの3番目の戻り値を `restore_dual` に明示的に渡し、全6件を再実行している。初回を含む12optimizer呼出しの開示と、最終6回の正常終了を分ける説明は適切である。

## Issue #28：歴史的D05A再監査

`VERIFY_HISTORICAL.py` と72個の必須ファイルを確認し、サイズ・SHA不一致0を確認した。保存された1622要素の下上物理場のGAP1.7206300416%と自己根記録の契約を確認する一方、反対端点のfull vectorによる根再生成ではないことを出力している。cold workerの終了コード未記録をnullで維持しており、wrapper約327秒を今回のVBA速度とは扱っていない。

保存されたattempt5は元dualのstationarity約3.41e-10で数値門に合格、attempt65は原物理場には合格するが元dualのstationarity約3.895e-7で数値門には不合格、という両方を再現する仕様だった。後者を隠さず、原門を緩めてもいない。座標dualから元dualへの逆変換を実ベクトルでバイト一致照合する。historicalな元dualの合否と、今回の固定点の逆符号専用guideを混同していない。

## 残る制約

生のIssue #28/#29の実行機ベクトル・全行列をpeerは持っていない。今回のデータからDLLの不具合、今回のHSD停止の唯一原因、通常VBA経路の2%成功を確定できない。#28固定点は原問題のconditioningと単位復元を調べる有効な小実験だが、原helper/HSDの修正確認には実問題と同一性を合わせた再現が必要である。

このレポート自身の範囲はread-only peer確認であり、親エージェントが実行したpositive replaysの第三の再実行ではない。ソース／manifest／結果の固定後に一式を梱包し、最終ZIPを別フォルダーへ展開した再監査の結果を別に残す。
