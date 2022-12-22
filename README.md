# JXler
## 概要
JsonとExcelを相互に変換します。  
多層構造のJsonも再帰的に処理することで、すべての情報をExcelに変換します。  
Excelは所定のフォーマットで書き出され、編集後そのままJsonに再変換することができます。  

## 画面
![image](https://user-images.githubusercontent.com/75108963/208922911-f87d1607-2ea2-46b4-ab20-58b163dbe971.png)

## 使用技術
- 言語 : C#
- フレームワーク : .NET 6.0、WPF
- 主な使用ライブラリ : ClosedXML、NewtonSoft.Json

## 使い方

### インストール/起動
Zipをダウンロード解凍し、JXler.exeをダブルクリックします。

### 基本的な使い方
- 上段のDataGridにファイルをドラッグ&ドロップします。
- 読み込まれたデータの「Action」列をクリックし処理内容を選択します。
- 実行ボタンをクリックします。
- 確認画面で出力先を選択しOKをクリックします。

## 詳細

### Actionについて
各行のActionボタンをクリックすると処理が切り替わります。
|Action|処理|
|:--|:--|
|>>|JsonからExcelに変換します。|
|<<|ExcelからJsonに変換します。|
|<>|変換しません。|

### 右クリックメニュー
DataGrid上で右クリックすることで表示します。
|メニュー|処理|
|:--|:--|
|追加|変換設定を追加します。（サブ画面が開きます）|
|更新|選択中の変換設定の内容を変更します。（サブ画面が開きます）|
|コピー|選択中の変換設定をコピーします。|
|削除|選択中の変換設定を削除します。|
|移動|選択中の変換設定に指定されているフォルダを開きます。|
|リロード|変換設定を再読み込みします。|

### 確認画面/出力先選択
実行ボタンをクリックすると確認画面が表示されます。出力方法を選択しOKをクリックすることで、変換処理が実行されます。
|メニュー|処理|
|:--|:--|
|入力ファイルと同じ|入力ファイルと同じフォルダに出力します。ファイル名は拡張子以外が入力ファイルと同じものになります。(例：Sample.Json → Sample.Xlsx)|
|パスを指定|サブ画面で出力先フォルダを選択します。ファイル名は「入力ファイルと同じ」を選んだ時と同じです。|
|個別設定に従う|行ごとの変換設定に従い出力します。|

## 付録

### 環境設定
ヘッダメニューの環境設定ボタンをクリックすると設定画面が開きます。

|設定項目|内容|備考|
|:--|:--|:--|
|出力Excelにリンクを設定する|「する」とした場合、親要素のセルに子要素へのリンクを設定します。|
|インデックスシート|Excelに書き出し始めるシートを設定します。Json変換時の最初のシートにもなります。|
|基底ディレクトリ|確認画面/出力先に「個別設定に従う」を選択した際、パスが指定されていない場合の代替ディレクトリになります。|
|Excel書式/Text|Excel変換時のText型の書式を指定します。|設定例 : @|
|Excel書式/DateTime|Excel変換時のText型の書式を指定します。|設定例 : yyyy/mm/dd HH:mm:ss|
|Excel書式/Int|Excel変換時のText型の書式を指定します。|設定例 : 0000|
|Excel書式/Float|Excel変換時のText型の書式を指定します。|設定例 : #.###|

### Excel変換フォーマット

Jsonのサンプルとともに変換フォーマットを説明します。

#### Sample1

- Json
```json
{
    "prop1" : "valu1",
    "prop2" : "valu2",
    "prop3" : "valu3"
}
```

- Excel

![image](https://user-images.githubusercontent.com/75108963/208298487-d24b6817-14b8-46df-8db6-130be963ebb8.png)

設定要領は以下の通り  
A1セル → 「No」固定表示　（※注１）  
B1セル移行1行目 → Jsonプロパティ  
2行目以降 → Json値  

※ シート名は環境設定で指定した値で作成されます。
※ A1セルにメモで「Object」が設定されます。これはExcelからJsonへ変換する際に必要な情報なので削除しないようにしてください。

#### Sample2

- Json
```json
[
    {
        "prop1" : "valu1-1",
        "prop2" : "valu1-2",
        "prop3" : "valu1-3"   
    },
    {
        "prop1" : "valu2-1",
        "prop2" : "valu2-2",
        "prop3" : "valu2-3"   
    }
]
```

- Excel

![image](https://user-images.githubusercontent.com/75108963/208298572-4f6c73d4-72ca-46d2-9081-f0b86c8a1a4d.png)

A1セル → 「No」のメモに「Array」が設定されます。これもObjectと同等に削除しないようにしてください。

#### Sample3

- Json
```json
[
    "valu1-1",
    "valu1-2",
    "valu1-3"   
]
```

- Excel

![image](https://user-images.githubusercontent.com/75108963/208298953-00699261-15aa-4b98-a54c-239fb2b9cdfe.png)

値のみの配列の場合、プロパティ名に「List」が固定で設定されます。


#### Sample4

- Json
```json
{
    "prop1" : {
        "prop1-1" : "value1-1",
        "prop1-2" : "value1-2",
        "prop1-3" : "value1-3"
    },
    "prop2" : [
        {
            "prop2-1" : "value2-1-1",
            "prop2-2" : "value2-2-1",
            "prop2-3" : [
                {
                    "prop2-3-1" : "value2-3-1",
                    "prop2-3-2" : "value2-3-2"
                },
                {
                    "prop2-3-1" : "value2-3-3",
                    "prop2-3-2" : "value2-3-4"
                }
            ]
        },
        {
            "prop2-1" : "value2-1-2",
            "prop2-2" : "value2-2-2",
            "prop2-3" : [
                {
                    "prop2-3-1" : "value2-3-5",
                    "prop2-3-2" : "value2-3-6"
                },
                {
                    "prop2-3-1" : "value2-3-7",
                    "prop2-3-2" : "value2-3-8"
                }

            ]
        }
    ],
    "prop3" : "valu3"
}
```

- Excel

![image](https://user-images.githubusercontent.com/75108963/208299767-2ade4f70-4b2b-473f-beaf-00fead746c9b.png)

子要素がある場合、親要素のセルには{プロパティ名_No.X}の形で値が設定されます。子要素は親要素のプロパティ名のシートが作成され値が設定されます。この時、子要素のNo列は、親要素の”No.x”にグルーピングされます。

### JsonとExcelの型

JsonはNewtonsoftの[JtokenTypeクラス](https://www.newtonsoft.com/json/help/html/t_newtonsoft_json_linq_jtokentype.htm)、とExcelは[ColsedXML](https://github.com/ClosedXML/ClosedXML)のXLDataTypeクラスに準じ、以下のとおり対応します。

|Json型(JTokenType)|Excel型(XLDateType)|備考|
|:--|:--|:--|
|String|Text|Excel出力時は環境設定/Excel書式/Textの指定に従います。|
|Integer|Number|Excel出力時は環境設定/Excel書式/Intの指定に従います。|
|Float|Number|Excel出力時は環境設定/Excel書式/Floatの指定に従います。|
|Date|DateTime|Excel出力時は環境設定/Excel書式/DateTimeの指定に従います。|
|Null|Text|Excel出力時は文字列で"{null}"を設定|
|Boolean|Boolean||

## License

[[MIT]](https://github.com/m-ishikawa1980/JXler/blob/master/LICENSE.md)
