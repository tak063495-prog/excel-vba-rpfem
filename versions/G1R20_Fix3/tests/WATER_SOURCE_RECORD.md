# Fix3 の水理出典記録

`RPX_Water` 内の公開共通APIは次のとおりです。

- `RPX_SourceCell(value) As String`: Variant型、文字列長、値を保持します。Double/SingleはIEEE表現を使用します。
- `RPX_SourceRange(cells As Range) As String`: シート名・アドレス・各セルのValue2/Formulaを型付きで保持します。単一セルにも対応します。
- `RPX_SourceRecordWrite(recordNamespace, identifier, context)`: ブックのCustomXMLPartsへ記録します。
- `RPX_SourceRecordRead(recordNamespace, ByRef identifier, ByRef context) As Boolean`: 欠落時だけFalseを返します。多重、破損、未知schema、空IDは例外です。

XMLは渡されたnamespaceの`rpxSource`ルート、`version="1"`、`identifier`と`context`の順の2子要素です。内容はVBA文字列のUTF16LEバイト列をbase64化したものです。日本語、改行、NUL、サロゲートを失いません。CustomDocumentPropertiesの文字列長制限を使いません。Writeは旧partの参照を固定してから新partを追加し、そのpartを読戻照合した後で旧partを削除し、namespace全体を再照合します。途中例外は成功扱いにしません。

水理出典のnamespaceは`urn:rpfem:source:seep:v1`、IDは浸透結果K9の計算IDです。contextは`RPX_SeepSourceContext()`の完全な型付き文字列です。形状A:C/領域E:H、材料ID/剛体ID、mode/γwと、そのmodeが利用する水理表・水理解法設定を保持します。mode0/3は透水係数、水理境界、krmin/β/緩和係数/収束判定/最大反復/剛体透水比、mode3は観測点/逆計算上限も含みます。mode1/2は地下水面表を含みます。TARGET・q・土の強度・機械荷重・GAP・DLL/ソルバー時間・表示や入力チェック結果は含みません。

`RPX_SeepRun`は失効後・ReadInputs前に`RPX_SeepSourceBegin`で再計算/捕捉し、生成後、結果モード出力後にAssert、FinishでEndを実行します。順解析の各反復と結果writerでもAssertします。writerはsourceをK9へ束縛して保存した後にK9/K2/COMMITTEDを出力します。直接の公開`RPX_SeepWriteResults`呼出しは入口で現contextを捕捉し、末尾まで一致を要求します。直接writerによる手動配列は従来と同じく呼出側の責任であり、このAPI自体が数値場の求解を証明するものではありません。

Loadは全modeでsourceの存在、ID一致、現入力一致を要求します。Fix2以前のsourceのない結果は、K2=Trueと旧ジャーナルがあっても再利用せず、再浸透解析を要求します。土解析の入力scopeは全modeと連成OFFのsource/連成・サクション/結果資格を保持し、途中変更時は解析を停止します。水平mode2の既存の特別経路と数値式はそのままです。

`WaterSourceFix3Probe.bas`の直接writerケースは数値的なFEM結果の証明ではなく、出典・永続化・拒否状態の検査です。公開producerのmode0/1/2/3、全modeの保存再開、実際の通常入力編集による拒否を別途Excelで実行します。portable contractスクリプトは呼出順序/依存範囲/src-import一致を検査する静的検査で、VBA実行や数値回帰の代わりではありません。
