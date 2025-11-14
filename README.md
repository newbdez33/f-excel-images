# excel_fuck_image

Excel の画像を扱うための Windows 向けユーティリティです。PowerShell の Excel COM を用いて、以下を自動化します。

- 画像の一括挿入（テンプレート行を複製し、指定セルに画像をフィット・中央配置）
- 画像の一括抽出（各シートごとに PNG で書き出し）
- 図形の一括削除（画像以外の図形を削除、グループ解除に対応）

## 前提条件
- Windows 10/11
- Microsoft Excel がインストール済み（COM 経由で操作します）
- Node.js（推奨: v18 以上）
- PowerShell 実行ポリシーが `Bypass` で実行可能（スクリプト内部で指定）

## インストール
特別な依存はありません。リポジトリを取得後、ディレクトリに移動してください。

```
cd excel_fuck_image
```

## 使い方
スクリプトは `npm run` または `node` で実行できます。`npm run` で引数を渡す場合は、`--` の後に付けてください。

### 1) 画像を Excel に挿入
- 画像ディレクトリ内の画像をソートして、テンプレート行を複製しながら指定セルに配置します。
- 画像はセルサイズに収まるように縦横比を維持してスケーリングし、中央に配置します。
- 任意で別列にファイル名を記録します。
- 実行前に対象 Excel を同ディレクトリへタイムスタンプ付きでバックアップします。

例（npm 経由）:
```
npm run insert -- --excel "C:\path\to\workbook.xlsx" --sheet "農政局" --dir "C:\path\to\images" --templateRow 3 --imageCol 6 --recordCol 7
```

例（node 直接実行）:
```
node insert_images_to_excel.js --excel "C:\path\to\workbook.xlsx" --sheet "農政局" --dir "C:\path\to\images" --templateRow 3 --imageCol 6 --recordCol 7
```

主なオプション:
- `--excel`: 対象 Excel ファイルのパス（必須に近い。指定が無い場合はスクリプト内の例が使用されます）
- `--sheet`: 対象シート名（見つからなければ先頭シート）
- `--dir`: 画像ディレクトリ（必須）
- `--templateRow`: 複製元となるテンプレート行番号（既定: 1）
- `--imageCol`: 画像を配置する列番号（既定: 1）
- `--recordCol`: 画像ファイル名を記録する列番号（省略可）
- `--dryRun`: 変更せずに設定と対象を出力して終了

動作仕様:
- シートの使用済み範囲が 4 行以上ある場合、4 行目以降をクリア
- 行>=4 にある図形を削除してから挿入処理
- マージセルは解除してセル寸法を取得
- 図形の `LockAspectRatio=true`、`Placement=2` を試行

### 2) Excel から画像を抽出
- 指定シートまたは全シートの画像を PNG で書き出します。
- 直接エクスポートが失敗する場合、コピー＆ペーストや Chart 経由のエクスポートにフォールバックします。

例（npm 経由）:
```
npm run extract -- --excel "C:\path\to\workbook.xlsx" --sheet "農政局" --out "C:\path\to\out"
```

例（node 直接実行）:
```
node extract_images_from_excel.js --excel "C:\path\to\workbook.xlsx" --sheet "" --out "C:\path\to\out"
```

主なオプション:
- `--excel`: 対象 Excel ファイルのパス
- `--sheet`: 対象シート名。空文字なら全シート対象
- `--out`/`--outDir`: 出力ディレクトリ（必須）。シートごとにサブフォルダ作成
- `--dryRun`: 変更せずに設定を出力

動作仕様:
- 画像タイプ（例: `msoPicture`, `msoLinkedPicture`）を対象
- グループ図形内の画像にも対応
- 可能な限り PNG で保存し、幅・高さの正当性を確認

### 3) 図形を一括削除（画像以外）
- 画像（ピクチャ）を除く図形を削除します。
- グループ図形は最大 5 ラウンドまで解除を試みてから削除します。
- 実行前に対象 Excel をバックアップします。

例（npm 経由）:
```
npm run delete-shapes -- --excel "C:\path\to\workbook.xlsx" --sheet "農政局"
```

例（node 直接実行）:
```
node delete_shapes_in_excel.js --excel "C:\path\to\workbook.xlsx" --sheet "農政局"
```

主なオプション:
- `--excel`: 対象 Excel ファイルのパス
- `--sheet`: 対象シート名。未指定なら全シート
- `--dryRun`: 変更せずに設定を出力

## 注意事項・既知の制限
- Excel が操作可能な拡張子（`xls`/`xlsx` など）である必要があります。
- マージセル解除によりレイアウトが変化する可能性があります。
- 画像はセルの中央に配置し、セルサイズに合わせて縮小・拡大されます。
- COM 経由の操作のため、Excel のバージョンや環境によって挙動が異なる場合があります。
- 実行前に自動バックアップが作成されますが、重要なファイルは手動でもバックアップしてください。

## トラブルシューティング
- PowerShell 実行ポリシーで失敗する場合: 管理者権限の PowerShell で `Set-ExecutionPolicy RemoteSigned` を検討してください。
- Excel の応答が無い/失敗する: 一度 Excel を完全に終了してから再実行してください。
- 権限関連のエラー: 出力/画像ディレクトリに対して書き込み権限があるか確認してください。

## ライセンス
`ISC` ライセンス。

