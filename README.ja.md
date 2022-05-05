# Suwarin

Version 1.0.1

![demo](https://github.com/taidalog/Suwarin/blob/images/images/suwarin.gif)

[English README](README.md)

[GitHub](https://github.com/taidalog/Suwarin)

## 目次

1. [概要](#%e6%a6%82%e8%a6%81)
1. [特徴](#%e7%89%b9%e5%be%b4)
1. [使い方](#%e4%bd%bf%e3%81%84%e6%96%b9)
    1. [実行結果](#%e5%ae%9f%e8%a1%8c%e7%b5%90%e6%9e%9c)
1. [座席表ファイルの作り方](#%e5%ba%a7%e5%b8%ad%e8%a1%a8%e3%83%95%e3%82%a1%e3%82%a4%e3%83%ab%e3%81%ae%e4%bd%9c%e3%82%8a%e6%96%b9)
    1. [罫線を引く](#%e7%bd%ab%e7%b7%9a%e3%82%92%e5%bc%95%e3%81%8f)
        1. [罫線のルール](#%e7%bd%ab%e7%b7%9a%e3%81%ae%e3%83%ab%e3%83%bc%e3%83%ab)
        1. [使用できる形式例](#%e4%bd%bf%e7%94%a8%e3%81%a7%e3%81%8d%e3%82%8b%e5%bd%a2%e5%bc%8f%e4%be%8b)
        1. [使用できない形式例](#%e4%bd%bf%e7%94%a8%e3%81%a7%e3%81%8d%e3%81%aa%e3%81%84%e5%bd%a2%e5%bc%8f%e4%be%8b)
    1. [数式を入力する](#%e6%95%b0%e5%bc%8f%e3%82%92%e5%85%a5%e5%8a%9b%e3%81%99%e3%82%8b)
    1. [マクロを追加する](#%e3%83%9e%e3%82%af%e3%83%ad%e3%82%92%e8%bf%bd%e5%8a%a0%e3%81%99%e3%82%8b)
1. [その他の機能](#%e3%81%9d%e3%81%ae%e4%bb%96%e3%81%ae%e6%a9%9f%e8%83%bd)
1. [詳細設定](#%e8%a9%b3%e7%b4%b0%e8%a8%ad%e5%ae%9a)
1. [License](#License)


## 概要

Excel ファイル上で座席表を作成するマクロです。**席替え用のマクロではなく、既にできている名簿から座席表を作るマクロです。** 新学期の出席番号順の座席表を作成する場合や、課外授業や総合的な探究の時間用の一時的な座席表を複数枚作成する場合などに役立つと思います。


## 特徴

- 参加者の名前や番号などをまとめて貼り付けて、2回クリックするだけで座席表が作成可能
- コンテキストメニュー（右クリックメニュー）から使用可能
- 座席表の形式がある程度柔軟
- 座席の開始位置（左下・右下・左上・右上）と方向（縦・横）を指定可能
- 使用禁止席を指定可能


## 使い方

1. 座席表ファイルを作る（初回のみ）  
    「[座席表ファイルの作り方](#%e5%ba%a7%e5%b8%ad%e8%a1%a8%e3%83%95%e3%82%a1%e3%82%a4%e3%83%ab%e3%81%ae%e4%bd%9c%e3%82%8a%e6%96%b9)」参照  
    ![座席表ファイルのレイアウト例](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin02.png)
1. 座席表の右上のセルから2つ右のセルに、参加者を貼り付け or 入力する  
    座席表の範囲が `B3:M16` だった場合、一番右上は `L3` なので、その2つ右の `O3` から下、つまり`O3`, `O4`, `O5` ... に貼り付ける
    ![座席表の右上のセルから2つ右に参加者を貼り付ける](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin03.png)
1. 適当なセルで右クリックする
1. コンテキストメニュー（右クリックメニュー）の一番下辺りにある、座席表ファイルと同じ名前の項目にマウスを乗せる
1. 横に飛び出たメニューの「座席表を作成(M)」をクリックする  


### 実行結果

参加者が、それぞれの座席の左上のセルに順番に入力されます。

![実行結果](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin04.png)

![実行結果（拡大）](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin05.png)

座席の左上のセル以外には一切入力しないので、必要に応じて数式を入れるなどしてください。もう少し詳しいことは[後述](#%e6%95%b0%e5%bc%8f%e3%82%92%e5%85%a5%e5%8a%9b%e3%81%99%e3%82%8b)。

![座席の数式の例](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin06.png)


## 座席表ファイルの作り方

このマクロを使うにあたって、初回のみ、以下の3つのステップが必要です。

1. [罫線を引く](#%e7%bd%ab%e7%b7%9a%e3%82%92%e5%bc%95%e3%81%8f)
1. [数式を入力する](#%e6%95%b0%e5%bc%8f%e3%82%92%e5%85%a5%e5%8a%9b%e3%81%99%e3%82%8b)
1. [マクロを追加する](#%e3%83%9e%e3%82%af%e3%83%ad%e3%82%92%e8%bf%bd%e5%8a%a0%e3%81%99%e3%82%8b)


### 罫線を引く

![座席表の罫線の引き方](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin02.png)

座席と座席表の範囲は、罫線で識別しています。逆に言うと、罫線（と「参加者」の位置）以外のものはマクロの動作に影響しません。罫線の種類や太さ、色も何でもいいです。


#### 罫線のルール

以下のルールに従って罫線を引いてください。ルールに則っていれば、座席表の形式はある程度自由に作れますし、既存のファイルも使用できます。

- 座席表の外周を、途切れることなく罫線で囲ってください  
    ![座席表の外周に罫線](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin08.png)
- 座席ひとつひとつを、途切れることなく罫線で区切ってください  
    ![座席を罫線で区切る](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin09.png)
- 座席表の範囲を越えた余分な罫線を引かないでください
- 座席内に罫線を引かないでください
- 座席の行数や列数を揃えてください
- 座席と座席の間に空白の行や列を入れないでください
- 1枚のシートに座席表は1つにしてください（2つ目以降は無視します）  
    ![注意事項諸々](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin10.png)


#### 使用できる形式例

それぞれの座席の行数や列数が揃っていれば、座席のサイズは何行×何列でもいいです。

![使用できる形式例 (縦2行×横2列)](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin11.png)

![使用できる形式例 (縦2行×横1列)](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin12.png)

![使用できる形式例 (縦3行×横1列)](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin13.png)


#### 使用できない形式例

座席と座席の間に空白の行や列を入れないでください。

![使用できない形式例](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin14.png)


### 数式を入力する

「[実行結果](#%e5%ae%9f%e8%a1%8c%e7%b5%90%e6%9e%9c)」でも書きましたが、このマクロは、

1. 座席表の右上のセルから2つ右のセルに貼り付けた参加者を
1. それぞれの座席の左上のセルに
1. 順番に入力する

![マクロの動作のイメージ図](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin15.png)

ものです。それ以外のセルへの入力は一切行いません。ですので「出席番号を入れると氏名が出るようにしたい」といった場合は、別シートに名簿を用意しておき、座席の左上以外のセルに `VLOOKUP` 関数等を入れてお使いください。

![座席の数式の例](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin16.png)

![別シートの名簿の例](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin17.png)


### マクロを追加する

以下の手順で作業してください。

以下の記事を参考にしてください。  
https://taidalog.hatenablog.com/entry/2022/05/05/100000

1. https://github.com/taidalog/Suwarin/releases から最新版の "Source code (zip)" をダウンロードする
1. 「[座席表ファイルの作り方](#%e5%ba%a7%e5%b8%ad%e8%a1%a8%e3%83%95%e3%82%a1%e3%82%a4%e3%83%ab%e3%81%ae%e4%bd%9c%e3%82%8a%e6%96%b9)」で用意した座席表の Excel ファイルに、ダウンロードした `main.bas` をインポートする
1. 座席表ファイルの `ThisWorkbook` モジュールに以下のコードをコピーして貼り付ける  
    ```
    Private Sub Workbook_Open()
        Call AddToContextMenu
    End Sub
    ```
1. 座席表ファイルを「Excel マクロ有効ブック (*.xlsm)」形式で「名前を付けて保存」する  
1. ファイルを閉じて、新しく保存した `.xlsm` の方を開く  
    今後は新しい方 (`.xlsm`) を使ってください

コンテキストメニュー（右クリックメニュー）の一番下辺りに、座席表ファイルと同じ名前の項目があれば設定完了です。

![コンテキストメニュー](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin22.png)


## その他の機能

- 使用禁止席を指定する
- 座席の開始位置（左下・右下・左上・右上）と方向（縦・横）を指定する
- 最後列の端数の座席を寄せる方向（中央寄せ・先頭寄せ・末尾寄せ）を指定する
- 座席表シートを、枚数を指定してコピーする

座席の左上のセルに<u>半角小文字のエックス</u>を入力すると、その座席を使用禁止にできます。

![使用禁止席を指定](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin23.png)

座席の開始位置と方向を指定できます。初期設定では「左下・縦」です。設定の変更方法は「[詳細設定](#%e8%a9%b3%e7%b4%b0%e8%a8%ad%e5%ae%9a)」の通りです。

左下・縦

![左下・縦](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin01.png)

左下・横

![左下・横](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin24.png)

右上・縦

![右上・縦](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin25.png)

列ごとの人数が揃わず、端数が出た場合、以下の画像のように、その端数の配置を指定できます。初期設定では「中央寄せ」です。設定の変更方法は「[詳細設定](#%e8%a9%b3%e7%b4%b0%e8%a8%ad%e5%ae%9a)」の通りです。

左下・縦・中央寄せ

![左下・縦・中央寄せ](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin01.png)

左下・縦・先頭寄せ

![左下・縦・先頭寄せ](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin26.png)

左下・縦・末尾寄せ

![左下・縦・末尾寄せ](https://github.com/taidalog/Suwarin/blob/images/images/ja/suwarin27.png)


## 詳細設定

以下の設定は、`main.bas` モジュール内で変更できます。

```
Public Sub CallMakeSeatingChart()
```

という行から

```
End Sub
```

という行までの間の、`settingSearchDirection = SearchByColumn` のような部分の `=` の**右側**を書き換えます。

`Call MakeSeatingChart` で始まる箇所は触らないでください。


### settingSearchDirection

座席の左上のセルを探す方向

|値|意味|
|---|---|
|SearchByColumn|縦方向|
|SearchByRow|横方向|


### settingSeatStart

座席表の一つ目の席の位置

|値|意味|
|---|---|
|BottomLeft|左下|
|BottomRight|右下|
|TopLeft|左上|
|TopRight|右上|


### settingSeatDirection

座席表の列の方向

|値|意味|
|---|---|
|ByColumn|縦|
|ByRow|横|


### settingSeatAlignment

最後列の端数の座席を寄せる方向

|値|意味|
|---|---|
|ToCenter|中央寄せ|
|ToFirst|前寄せ|
|ToLast|後ろ寄せ|


### stringToSkip

使用禁止席を示す記号や文字列

`stringToSkip =` の右に、任意の記号や文字列を `"` でくくって入力してください。


## License

Copyright 2022 taidalog

Suwarin is licensed under the MIT License.
