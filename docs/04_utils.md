# ユーティリティ

テンプレートの作成に役に立ちそうなサポートツール。
bin 下のもそうだが、

    python3 -i スクリプト オプション

として実行すると、処理終了後 python の対話モードに落ちるので、Writer
文書に python から対話的にアクセスできる。使い方は簡単ではないが。

inserter.document.document が文書に対応するオブジェクトである。

## `utils/user_var_setter.py`

    utils/user_var_setter.py CSV

CSVデータを読み込んで、CSVレコードとして定義されたものを現在アクティブ
なWriter文書にユーザフィールドとして追加する。

入力データが既に決まっている所からテンプレートを作成する際、ユーザフィー
ルドを先に定義しておける。

また作成中のテンプレートに対して実行すれば、CSVデータを差し込んでどう
なるか確認できる。ただし、画像系の差し替え動作などは行なわない。
