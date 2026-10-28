# Excel VBA Project Template

Excel VBA（VBE）を使うプロジェクトのひな形です。新しいプロジェクトを作るときは、このフォルダーをコピーしてプロジェクト名に変更し、コピー先で作業してください。

## このテンプレートの内容

```text
ProjectName/
├─ README.md          プロジェクトの目的と使い方
├─ .gitignore         Gitに登録しないファイルの指定
├─ Source/            VBEから書き出した本番用VBAソースを置く
├─ Sandbox/           試作・検証用コードを置く（詳細はSandbox/README.md）
├─ ProjectName.xlsm   Excelで作成・保存するマクロ有効ブック
└─ .git/              コピー後に作成するプロジェクト専用Git履歴
```

テンプレートには`.git`とプロジェクト固有の`.xlsm`を含めません。これらはコピー先のプロジェクトで用意します。

## 新しいプロジェクトの準備

1. このフォルダーをコピーし、`Workspace\\Projects`の下に`Project プロジェクト名`として配置します。
2. コピー先にExcelブックを作成し、マクロ有効ブック（`.xlsm`）として保存します。
3. コピー先のプロジェクトフォルダーをVS Codeで開きます。
4. VS Codeのターミナルで、そのプロジェクト専用のGitリポジトリを作成します。

```powershell
git init -b main
```

## VBAコードの管理

VBAコードの本体はExcelブック（`.xlsm`）内に保存されます。履歴や差分をGitで確認しやすくするため、VBEから標準モジュール、クラス、フォームをエクスポートし、`Source`以下に保存します。

```text
Source/
├─ ブック名/                         VBAソース
│  ├─ モジュール名.bas
│  ├─ クラス名.cls
│  └─ フォーム名.frm と関連する .frx
└─ .excel-vba-sync-backups/
   └─ ブック名/                      同期管理用のJSON
      └─ モジュール名.syncbase.json
```

取得時は、同じブック名のフォルダーが `Source/` 直下と `Source/.excel-vba-sync-backups/` 以下にそれぞれ作成されます。VBAソースは種類別の `Modules/`、`Classes/`、`Forms/` には分けず、`Source/ブック名/` に保存します。ファイルの種類は対象ブックと取得方法によって異なり、シートや `ThisWorkbook` のコードが `.bas` として出力される場合もあります。

`.syncbase.json` には、ソースファイル名、コードのSHA-256ハッシュ、取得元、記録日時などが保存されます。VBAコード本体ではなく、JSONだけではコードを復元できません。フォルダー名が `.excel-vba-sync-backups` でも、添付例のJSONをVBAソースのバックアップそのものとして扱わないでください。

エクスポートしたファイルを編集しただけでは、ブック内のVBAコードは変わりません。編集内容をExcelで使うには、VBEへインポートするか、VBE上で編集した後に改めてエクスポートしてください。ブックとエクスポートしたソースの両方を更新して管理します。

## Gitで変更を保存

`.gitignore`を確認してから、プロジェクトのファイルをGitに登録します。

```powershell
git status
git add .
git commit -m "Initial project setup"
```

Gitコマンドは必ずコピー先のプロジェクトフォルダーで実行してください。Excelの一時ファイルや秘密情報が登録対象になっていないことも確認します。
