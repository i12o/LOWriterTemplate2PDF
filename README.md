# LOWriterTemplate2PDF

You can create PDF files from ODT template, filling it with data
read from CSV text.

  - [CSV Data Sample](./samples/test_image_generation.csv)
  - [ODT Template Sample](./samples/test_image_generation.odt)
  - [Resulting PDFs](./samples/test_image_generation.result/)

After checking out this repository, go to top of your working copy.
Then do these commands to create above sample, on your machine.

Start libreoffice with server functinality.  It's beyond my ability to
describing the meaning of arguments.  If you have running LO already,
first shutdown LO, and restart LO with these args.

    libreoffice --accept='socket,host=localhost,port=2083;urp;'

Create PDF under pdfout directory, with sample test_image_generation.

    bin/generic_form_insertion.py -t samples/test_image_generation.odt \
        test_image_generation.csv


LibreOffice Writer で作成した帳票テンプレートに、CSVのデータを差し込み、
PDF としてエクスポートします。

`bin/generic_form_insertion.py --odt TEMPLATE.odt data.csv`

などとして実行すると、pdfout ディレクトリに data.csv(タブ切り) の一行
毎に対応する PDF が出力されます(直接プリンタに出力する機能もあります
が、十分にはテストしていません)。

普通のテキストの差し込み、バーコード/QRコードの生成、データから対応す
る画像の差し込みなどが可能です。

帳票テンプレートは Writer で作成し、テキストを差し込む部分はフィールド
(フィールドタイプ:ユーザ蘭)を配置し、画像の差し込みは画像に特殊な名前
を付ける事でデータのカラムから指定の処理で生成した画像に置き換えます。

画像の差し込み機能として、生成したバーコードでの置換の他、チェックボッ
クスやラジオボタン風(複数のチェックが可能)に対応しています。チェックボッ
クス(Toggle型と呼称します)は何種類かの画像からどれかを選んで差し替える
ので、丸付け、ロゴ切り替えなどにも使えます。

作成済みのスクリプトを使うだけであれば、端末やコマンドプロンプトを操作
できる人には使えると思います。が、CLIにまったく馴染みがない人でも使い
やすいようには作っておりません。多少のコードは書ける人の方が活用できる
と思います。

[ドキュメント](./docs/01_intro.md)にある程度詳しい説明があります。

## システム要件

  - Python3
  - LibreOffice(6.1.5 で開発)
  - python3-uno
  - python3-barcode
  - python3-pyqrcode
  - python3-pandas

に依存します。LibreOffice の古いバージョンでの動作は確認していません。
UNOが使えれば、よほど古くなければ動かせると思います。Linux 上で開発し
ていますので、それ以外の環境で動作するかは不明です。

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

## 実用的なサンプル

年賀はがきの宛名印刷用のテンプレートと処理用のスクリプトがあります。
ダミーの住所を使った`samples/nenga_address_sample.csv`サンプルCSVがあ
りますので、

    bin/hagaki_atena_sasikomi.py \
       -t samples/nenga_hagaki_multi_pattern.odt \
       samples/nenga_address_sample.csv

として、'pdfout/' の PDF を見てみて下さい。

テンプレートの `samples/nenga_hagaki_multi_pattern.odt` を修正すれば、
通常のはがきやレイアウトの異なる年賀はがきにも対応できます。多分差出人
の表示位置の修正が必要でしょう。

縦書き部分の英数字の向きが気になるようでしたら、テンプレートを修正して
縦書きにも対応しているフォントを指定して下さい。多分フォントの変更だけ
でどうにかなる問題だと思います。
