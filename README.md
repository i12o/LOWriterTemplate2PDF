# LOWriterTemplate2PDF

LibreOffice Writer で作成した帳票テンプレートに、CSVのデータを差し込み、
PDF としてエクスポートします。

`bin/generic_form_insertion.py --odt TEMPLATE.odt data.csv`

などとして実行すると、pdfout ディレクトリに data.csv(タブ切り) の一行
毎に対応する PDF が出力されます。

普通のテキストの差し込み、バーコード/QRコードの生成、データから対応す
る画像の差し込みなどが可能です。

帳票テンプレートは Writer で作成し、テキストを差し込む部分はフィールド
(フィールドタイプ:ユーザ蘭)を配置し、画像の差し込みは画像に特殊な名前
を付ける事でデータのカラムから指定の処理で生成した画像に置き換えます。

画像の差し込み機能として、生成したバーコードでの置換の他、チェックボッ
クスやラジオボタン風(複数のチェックが可能)に対応しています。チェックボッ
クス(Toggle型と呼称します)は何種類かの画像からどれかを選んで差し替える
ので、丸付け、ロゴ切り替えなどにも使えます。

あまり python の知識がない人でも使いやすいようには作っておりません。多
少のコードは書ける人向けです。

## システム要件

  - Python3
  - LibreOffice(6.1.5 で開発)
  - python3-uno
  - python3-barcode
  - python3-pyqrcode
  - python3-pandas

に依存します。LibreOffice の古いバージョンでの動作は確認していません。
UNOが使えれば、よほど古くなければ動く気はします。
Linux 上で開発していますので、それ以外の環境で動作するかは不明です。

私 python でのプログラミング経験はほぼないため、ロクでもないコードになっ
ていると思われます。とりあえずご容赦下さい。

## サンプル

とりあえず samples/ 以下に

  -
    - テンプレート `test_sasikomi.odt`
    - 差し込みデータ `test_sasikomi.csv`
  -
    - テンプレート `test_toggle_selection.odt`
    - 差し込みデータ `test_toggle_selection.csv`

他を用意してあります。

事前に LibreOffice にコマンド行オプションを付けてローカルホストのポー
ト2083で接続できるように起動しておきます。

    libreoffice --accept='socket,host=localhost,port=2083;urp;'

サンプルを使って

    bin/generic_form_insertion.py -t samples/test_sasikomi.odt \
        test_sasikomi.csv

などと実行して、pdfout に作成される PDF をご覧下さい。LibreOfficeに接
続するためのポートは、スクリプトの `--loport` オプションで変更できます。
サンプルのテンプレートには色々と解説も含めてあります。

書きかけですが多少の[ドキュメント](./docs/01_intro.md)があります。
