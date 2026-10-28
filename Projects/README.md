# Excel VBA Projects

このフォルダーには、業務用Excelマクロのプロジェクト、UserFormのサンプルとCSV編集ツール、新規プロジェクト用テンプレートがあります。各プロジェクトの目的と管理対象を以下にまとめます。READMEの説明とエクスポート済みソースが一致しない場合や、ブックの実際の動作を確かめる場合は、対象ブックと`Source`内のコードを確認してください。

## 目的から選ぶ

| したいこと | 使用するプロジェクト |
| --- | --- |
| 創意工夫テーマ管理シートのデータを抽出・転送する | [Project CSV Data Transfer](Project%20CSV%20Data%20Transfer/README.md) |
| 創意工夫履歴ブックのメモ移動やセル表示を操作する | [Project Project ImprovementHistory](Project%20Project%20ImprovementHistory/README.md) |
| TL管理シートのExcelマクロを使う | [Project TL-ManagementSheet](Project%20TL-ManagementSheet/README.md) |
| 交通安全ルーレットやリスト用ルーレットを使う | [Project TrafficRoulette](Project%20TrafficRoulette/README.md) |
| UserFormの情報をCSVで確認・編集する | [Project UserForm](Project%20UserForm/README.md) と[CSV Exporterツール](../.tools/UserFormCsvExporter/README.md) |
| 新しいExcel VBAプロジェクトを始める | [_Template_Excel](_Template_Excel/README.md) |

## UserForm CSV ツール

`../.tools/UserFormCsvExporter/`には、UserFormのコントロール情報をCSVへ出力する処理と、編集したCSVの一部プロパティを既存フォームへ反映するPowerShellスクリプトがあります。タスク定義はワークスペース直下の`.vscode/tasks.json`にあります。VS Codeでは`Workspace.code-workspace`またはワークスペース直下を開いて実行してください。利用前に[ツールのREADME](../.tools/UserFormCsvExporter/README.md)で専用アドインの準備、Trust Center設定、反映対象と制限を確認してください。

## VBAコードの管理

マクロの実行用コードは各`.xlsm`ブック内にあります。`Source`にある`.bas`、`.cls`、`.frm`などは、差分確認や履歴管理のために保存するソースです。ソースファイルだけを編集しても、ブック内のコードは変わりません。

VBAを変更するときは、対象プロジェクトのREADMEに従ってブックとソースを更新してください。通常は、VBEでコンパイルと動作確認を行い、ブックを保存したうえで、対応するエクスポート済みソースも更新します。UserFormでは`.frm`と関連する`.frx`を一緒に扱います。新規プロジェクトでの標準的な管理手順は[_Template_ExcelのREADME](_Template_Excel/README.md)にあります。

## プロジェクトを選ぶときの注意

- 各プロジェクトは独立したフォルダーです。実行やGit操作は、対象プロジェクトのREADMEを確認してからそのフォルダーで行ってください。
- プロジェクトによって、ブック本体をGit管理するかどうかが異なります。共通設定と決めつけず、個別の`.gitignore`、READMEを確認してください。
- CSV Data TransferとTL-ManagementSheetのREADMEは、現状ではテンプレートの説明です。Project ImprovementHistoryのREADMEにも、現在の`Source`内では確認できないマクロの説明があります。TrafficRouletteのREADMEは、抽選方式、使用範囲、ブック間の違い、現状の制約を説明しています。
- Project UserFormのREADMEはVBAソース管理のガイドです。UserForm CSVツールの具体的なセットアップ・実行方法は[ツールREADME](../.tools/UserFormCsvExporter/README.md)を確認してください。
- この一覧は入口です。セル範囲やマクロ名など、実行に必要な条件はブックと現在の`Source`内コードで確認してください。
