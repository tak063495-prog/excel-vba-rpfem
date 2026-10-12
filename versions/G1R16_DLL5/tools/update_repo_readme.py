import json
from prepare import ROOT,REPO

def main():
 c=json.loads((ROOT/'results/comparison.json').read_text(encoding='utf8'));t=c['seconds']
 section=f'''## G1R16 DLL5 — Schur・HSD修復・疎構造・一括演算の追加最適化（最新版）

2026年10月12日。主要な追加最適化をDLL5にまとめました。堤防例題は **DLL4 {t['DLL4_fresh']:.2f}秒 → DLL5 {t['DLL5_fresh']:.2f}秒（{c['reduction_percent_vs_DLL4']:.2f}%短縮、約{c['speedup_vs_DLL4']:.2f}倍）**。両版を同一PCで順番に再測定しています。DLLなしの既存実測{t['VBA_existing']:.2f}秒比では約{c['speedup_vs_VBA_reference']:.2f}倍ですが、VBAの長時間測定は再実行していません。

- **[対応XLSM・DLL・ソース・検証資料の一式ZIP](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/downloads/RPFEM_G1R16_DLL5_20261012.zip)**
- [対応XLSM](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/versions/G1R16_DLL5/delivery/RPFEM_20261012_G1R16_DLL5_Levee_Checked.xlsm) / [DLL5](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/versions/G1R16_DLL5/delivery/rpfem_native_x64.dll)
- [使い方](versions/G1R16_DLL5/README.md) / [検証報告・DLLなし基準の速度表](versions/G1R16_DLL5/REPORT.md) / [実装範囲](versions/G1R16_DLL5/IMPLEMENTATION_JA.md) / [導入・復旧](versions/G1R16_DLL5/IMPORT_ROLLBACK_JA.md) / [VBA差分](versions/G1R16_DLL5/changes.diff) / [native差分](versions/G1R16_DLL5/native_changes.diff)

通常内点法のSchur内部処理、HSDのDD組立とGMRES修復、疎パターン準備、入力比較、時間区間ごとの一括処理、DD/疎Double演算の細かい呼出しを追加移植しました。物理モデル・Fs定義・監査・許容誤差・救済順序は維持しています。堤防のFs下界{c['lower']:.15g}・上界{c['upper']:.17g}、53求解・603監査がVBA/DLL3/DLL4記録と一致しました。固定FsのC5/P06A/P06N計6組は全解成分とメッシュがVBA/DLL4/DLL5でビット一致しています。

主回帰22グループ、HSD既知解6組、追加の障害注入24条件、入力・世代・所有権・DLL不在/旧版復帰などを検証しました。P06Aの粗メッシュはFs≥0.125の15点だけ、P06N/D03N/D07Aは限定した粗メッシュ終端までです。粗メッシュP06NのFs未取得はDLL4と同じであり、回帰一致とFs取得を区別しています。フルP06A/P06N、極小Fs探索、以前中断した26ケース全試験は再開していません。

XLSMとDLLを組で同じフォルダーに置き、交換前にExcelを終了してください。DLL不在・非対応時はVBAで計算できます。詳細な実測条件・未実施範囲は報告書を参照してください。

'''
 path=REPO/'README.md';s=path.read_text(encoding='utf8');marker='## G1R16 DLL4 — 通常内点法・HSDの線形求解と精度補正をまとめてDLL化（最新版）'
 assert marker in s and '## G1R16 DLL5' not in s
 s=s.replace(marker,section+marker.replace('（最新版）','（前版）'),1)
 path.write_text(s,encoding='utf8',newline='\n');print('Repository README updated',flush=True)
if __name__=='__main__':main()
