# ExcelTool
![ICON](https://om.namanime.com/Pictures/ExcelTool/ExcelTool.ico)  
プラグイン拡張機能、バッチ処理、およびExcelファイルへの結果の出力。  
[中文ReadMe](https://github.com/ZjzMisaka/ExcelTool/blob/main/README_zh-CN.md) | [English ReadMe](https://github.com/ZjzMisaka/ExcelTool/blob/main/README.md)  
[例](https://github.com/ZjzMisaka/AnalyzersForExcelTool)  
[実行中のgif画像](https://www.namanime.com/ZjzMisaka/ExcelTool/ExcelTool.gif?20220603)

### 多言語
- [x] 简体中文
- [x] 日本語
- [x] English

### メインインターフェース
- ドロップダウンボックスから検索情報と処理ロジック（プラグイン）を選択し、1つずつ対応させます。
- 着信パラメータの設定、検索ルートディレクトリとデフォルトの出力ディレクトリ、出力ファイル名の設定が可能
- 実行モードには、順次実行と同時実行の2つがあります。
- 上記のコンテンツは、ルールのドロップダウンボックスから保存されたルールを選択することで自動的に入力できます
    - ルールを選択した後、いくつかのフォルダとファイルを監視するように監視を設定し、変更があったときにこのルールを自動的に実行できます  

### 検索情報設定インターフェース
**指定されたパスで指定されたExcelファイルを見つける必要がある指定されたシートを設定するために使用されます**
- 検索方法は、すべて、完全一致、部分包含、正規表現を選択できます

### 処理ロジック（プラグインコーディング）インターフェース
**特定の種類のシートの処理ロジックと処理後の出力ロジックを設定するために使用されます**
- エディターでコードを書くと、操作中に順番に実行されます。  
- パラメータを設定でき、プラグインユーザーはメインインターフェイスでパラメータを編集し、実行時に使用するコードにそれらを渡すことができます。  

#### コーディング関連
- プロセス全体の自動完了と色付け、dllファイルを自分でDllsフォルダーに追加でき、追加後に直接参照できます
- コンテンツのエンコードは、[ClosedXML](https://github.com/ClosedXML/ClosedXML)オープンソースライブラリに依存します
- 追加の提供された関数とプロパティを使用して、オンザフライで実行できます
    - メインインターフェイスのログ領域にあるログのリアルタイム出力
    - ハングして待って、ユーザー入力を読む
    - 追加のExcelファイル操作処理
- コンパイルエラーまたは実行エラーが発生すると、関連するデバッグ情報がメインインターフェイスの下部にあるログ領域に表示されます

##### Logger (静的クラス)
```c#
// ---- 出力ログ機能 ----
// 出力ログの種類に応じて、色の違いが異なります. 
void Info(string info);
void Warn(string warn);
void Error(string error);
void Print(string str);

// ---- 関数が見つからない場合に警告するかどうか ----
bool IsOutputMethodNotFoundWarning { get => isOutputMethodNotFoundWarning; set => isOutputMethodNotFoundWarning = value; }
```

##### Scanner (静的クラス)
```c#
// ---- 入力を取得 ----
// パラメータは入力を取得するためのプロンプトであり、実行後、ユーザーが入力するまで待機します. 
// 入力を待機している他のスレッドがある場合、最初に前のスレッドが入力を取得するのを待ってから、このステートメントの内容を実行します.  
string GetInput();
string GetInput(string value);

// ---- 入力を待ち ----
// おそらく役に立たない機能です。他のスレッドが入力を実行している間、ユーザー入力が取得されるまで待つことができます。
// 最新のユーザー入力を返します。 
string WaitInput();

// ---- 最近の入力内容 ----
string LastInputValue { get => lastInputValue; set => lastInputValue = value; }
```

##### Output (静的クラス)
```c#
// ---- Excelファイルの操作 ----
// 新しいexcelファイルを作成する
XLWorkbook CreateWorkbook(string name);
// CreateWorkbookで作成されたExcelファイルを取得します
XLWorkbook GetWorkbook(string name);
// シート取得します
IXLWorksheet GetSheet(string workbookName, string sheetName);
IXLWorksheet GetSheet(XLWorkbook workbook, string sheetName);
// 作成されたすべてのexcelファイルを取得します
Dictionary<string, XLWorkbook> GetAllWorkbooks();
// 作成したすべてのExcelファイルをクリアする
void ClearWorkbooks();

// ---- デフォルトの出力ファイルを保存するかどうか ----
bool IsSaveDefaultWorkBook { get => isSaveDefaultWorkBook; set => isSaveDefaultWorkBook = value; }
```

##### Param (クラス)
```c#
// パラメータを取得する
List<string> Get(string key);
string GetOne(string key);
// パラメータキーのコレクションを取得する
IEnumerable<String> GetKeys();
// パラメータが含まれているかどうかを確認する
bool ContainsKey(string key);
```

##### RunBeforeAnalyze関数
|パラメータ|タイプ|説明|メモ|
|----|----|----|----|
|param|Param|着信パラメータ||
|globalObjects|Object|グローバルに存在し、現在の行番号など、他の呼び出しで使用する必要のあるデータを保存できます。||
|allFilePathList|List<string>|分析されるすべてのファイルパスのリスト||
|isExecuteInSequence|bool|順番実行するかどうか||

##### AnalyzeSheet関数
|パラメータ|タイプ|説明|メモ|
|----|----|----|----|
|param|Param|着信パラメータ||
|sheet|IXLWorksheet|分析するシート||
|result|ConcurrentDictionary<ResultType, Object>|現在のファイルの情報を保存する||
|globalObjects|Object|グローバルに存在し、現在の行番号など、他の呼び出しで使用する必要のあるデータを保存できます。||
|isExecuteInSequence|bool|順番実行するかどうか||
|invokeCount|int|この分析関数が呼び出された回数|最初の呼び出しでの値は1です|

##### RunBeforeSetResult関数
|パラメータ|タイプ|説明|メモ|
|----|----|----|----|
|param|Param|着信パラメータ||
|workbook|XLWorkbook|出力用のExcelファイル||
|globalObjects|Object|グローバルに存在し、現在の行番号など、他の呼び出しで使用する必要のあるデータを保存できます。||
|resultList|ICollection<ConcurrentDictionary<ResultType, Object>>|すべてのファイルに関する情報||
|allFilePathList|List<string>|分析されたすべてのファイルパスのリスト||
|isExecuteInSequence|bool|順番実行するかどうか||

##### SetResult関数
|パラメータ|タイプ|説明|メモ|
|----|----|----|----|
|param|Param|着信パラメータ||
|workbook|XLWorkbook|出力用のExcelファイル||
|result|ConcurrentDictionary<ResultType, Object>|現在のファイルの情報を保存する||
|globalObjects|Object|グローバルに存在し、現在の行番号など、他の呼び出しで使用する必要のあるデータを保存できます。||
|isExecuteInSequence|bool|順番実行するかどうか||
|invokeCount|int|この出力関数が呼び出された回数|最初の呼び出しでの値は1です|
|totalCount|int|出力関数を呼び出す必要がある合計回数|最後の呼び出しは、invokeCountがtotalCountと同じ場合です|

##### RunEnd関数
|パラメータ|タイプ|説明|メモ|
|----|----|----|----|
|param|Param|着信パラメータ||
|workbook|XLWorkbook|出力用のExcelファイル||
|globalObjects|Object|グローバルに存在し、現在の行番号など、他の呼び出しで使用する必要のあるデータを保存できます。||
|resultList|ICollection<ConcurrentDictionary<ResultType, Object>>|すべてのファイルに関する情報||
|allFilePathList|List<string>|分析されたすべてのファイルパスのリスト||
|isExecuteInSequence|bool|順番実行するかどうか||

# 使用されるオープンソースライブラリ
|オープンソースライブラリ|オープンソースプロトコル|
|----|----|
|[roslynpad/roslynpad](https://github.com/roslynpad/roslynpad)|MIT|
|[icsharpcode/AvalonEdit](https://github.com/icsharpcode/AvalonEdit)|MIT|
|[JamesNK/Newtonsoft.Json](https://github.com/JamesNK/Newtonsoft.Json)|MIT|
|[ClosedXML/ClosedXML](https://github.com/ClosedXML/ClosedXML)|MIT|
|[rickyah/ini-parser](https://github.com/rickyah/ini-parser)|MIT|
|[amibar/SmartThreadPool](https://github.com/amibar/SmartThreadPool)|MS-PL|
|[punker76/gong-wpf-dragdrop](https://github.com/punker76/gong-wpf-dragdrop)|BSD-3-Clause|
|[Kinnara/ModernWpf](https://github.com/Kinnara/ModernWpf)|MIT|
|[CommunityToolkit/WindowsCommunityToolkit (Microsoft.Toolkit.Mvvm)](https://github.com/CommunityToolkit/WindowsCommunityToolkit)|MIT|
|[microsoft/XamlBehaviorsWpf](https://github.com/microsoft/XamlBehaviorsWpf)|MIT|
|[ZjzMisaka/CustomizableMessageBox](https://github.com/ZjzMisaka/CustomizableMessageBox)|WTFPL|
