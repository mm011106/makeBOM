# makeBOM

## EagleCADの出力するBOMからP板.comの部品実装に必要な部品表を作成します。
## 環境
- Python 3
- pandas
- 部品辞書（後述）ファイル

## 仕組み
- Eagleが出力する部品表から、部品番号、部品名、シート番号を取り出す
- 部品名をキーにして、グルーピングする（使用されている部品ごとに、どの部品番号として使われているかをリストアップ）
- 部品名から部品辞書（MI)を引いて、正式名称やメーカなどの情報（部品実装に必要）を取り出す
- それらをマージして部品名ごとに１レコードとなるような部品表をつくる
- CSVファイルに出力する

## 例
上が入力データ、下が出力データ
![image](https://user-images.githubusercontent.com/9587359/112434301-c24ea480-8d86-11eb-91f5-4aa2bd8fe7e2.png)

上のテキストのValue欄の値を元に部品辞書（MI）を検索して必要なデータを付加し、下のようなデータになる。
