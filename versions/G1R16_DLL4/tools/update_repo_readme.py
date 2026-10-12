import json
from prepare import ROOT,REPO
def main():
 c=json.loads((ROOT/'results/comparison.json').read_text(encoding='utf-8'))
 section=f'''## G1R16 DLL4 — 通常内点法・HSDの線形求解と精度補正をまとめてDLL化（最新版）

2026年10月12日。線形求解・反復補正を状態付きのDLL内処理にし、HSDのLU因子をDLLに保持します。残差・方向・境界・状態更新も追加移植しました。堤防例題は **DLL_3 {c['DLL3_seconds']:.2f}秒 → DLL4 {c['DLL4_seconds']:.2f}秒（{c['reduction_percent_vs_DLL3']:.2f}%短縮、約{c['speedup_vs_DLL3']:.2f}倍）**。同一PC・同一入力で両版を順番に再測定しています。

- **[対応XLSM・DLL・ソース・検証資料の一式ZIP](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/downloads/RPFEM_G1R16_DLL4_20261012.zip)**
- [対応XLSM](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/versions/G1R16_DLL4/delivery/RPFEM_20261012_G1R16_DLL4_Levee_Checked.xlsm) / [DLL4](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/versions/G1R16_DLL4/delivery/rpfem_native_x64.dll)
- [使い方](versions/G1R16_DLL4/README.md) / [検証報告・速度表](versions/G1R16_DLL4/REPORT.md) / [実装範囲](versions/G1R16_DLL4/IMPLEMENTATION_JA.md) / [導入・復旧](versions/G1R16_DLL4/IMPORT_ROLLBACK_JA.md) / [VBA差分](versions/G1R16_DLL4/changes.diff) / [native差分](versions/G1R16_DLL4/native_changes.diff)

堤防のFs下界{c['lower']:.15g}・上界{c['upper']:.17g}、全{c['DLL3_log_comparison']['solves']}求解・{c['DLL3_log_comparison']['audits']}監査がVBA/DLL_3記録と一致。HSD6組、q8のC5・多層P06A/P06N固定点6組、P06N/D03N/D07Aの粗メッシュ終端比較、P06AのFs≥0.125の15点、LU360・境界5070・取消/入力変更/メモリ不足/未知例外等を比較しました。粗メッシュの一部ではDLL_3でもFsを取得できず、失敗を含む回帰一致と安全率取得を区別しています。P06Aフル適応解析はユーザー指示で中断しました。追加の粗メッシュ探索ではDLL_3が極小Fsで未収束となり、ユーザーが選択した15点に比較範囲を絞りました。その未収束のDLL4再現は未実施です。P06Nフル適応解析は開始していません。`PARTIAL_CERTIFIED`と目標GAP達成も区別します。26ケース全再実行ではありません。

XLSMとDLLを組で同じフォルダーに置き、交換時にはExcelを終了します。DLL不在・非対応版ではVBAに切替します。固定Fsの非線形反復制御、Fs探索、監査、救済、メッシュは既存VBAを維持しています。詳細な検証条件・未実施範囲は報告書を参照してください。

'''
 path=REPO/'README.md';s=path.read_text(encoding='utf-8');marker='## G1R16 DLL_3 — 通常内点法・HSDの追加DLL化（最新版）';assert marker in s and '## G1R16 DLL4' not in s
 s=s.replace(marker,section+marker.replace('（最新版）','（前版）'),1);path.write_text(s,encoding='utf-8',newline='\n');print('Repository README updated',flush=True)
if __name__=='__main__':main()
