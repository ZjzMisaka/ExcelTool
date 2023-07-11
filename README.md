# ExcelTool
![ICON](https://raw.githubusercontent.com/ZjzMisaka/ExcelTool/main/ExcelTool/ExcelTool.ico)  
Perform batch reading, analysis, and output operations of Excel by executing pre-written c# scripts.  
[中文ReadMe](https://github.com/ZjzMisaka/ExcelTool/blob/main/README_zh-CN.md) | [日本語ReadMe](https://github.com/ZjzMisaka/ExcelTool/blob/main/README_ja-JP.md)  
[View example](https://github.com/ZjzMisaka/AnalyzersForExcelTool)  
[Gif image when executing](https://www.namanime.com/ZjzMisaka/ExcelTool/ExcelTool.gif?20220603)

### Multi-language
- [x] 简体中文
- [x] 日本語
- [x] English

### Main interface
- Select retrieval information and processing logic (plug-in) through the drop-down box, make sure they are correspond one by one. 
- Can set incoming parameters, set search root directory and default output directory, output file name. 
- There are two execution modes: sequential execution and simultaneous execution. 
- The above content can be automatically filled in by selecting a saved rule from the rule drop-down box. 
    - After selecting a rule, you can set monitoring to monitor some folders and files, and automatically execute this rule when there is a change. 

### Search information setting interface
**Used to set the specified Sheet that needs to find the specified Excel file in the specified path**
- The search method can be selected all, complete match, partial contain and regular expression. 

### Processing logic (plug-in coding) interface
**Used to set the processing logic for a certain type of Sheet and the output logic after processing**
- Write code in the editor, and it will be executed in sequence during operation.
- Parameters can be set, plug-in users can edit the parameters in the main interface, and pass them to the code to use at runtime.  
- Parameter description and running log output can be multilingualized. 

#### Coding related
- Automatic completion and coloring throughout the process, you can add dll files to the Dlls folder by yourself, and you can directly reference them after adding. 
- Encoding content depends on the [ClosedXML](https://github.com/ClosedXML/ClosedXML) open source library. 
- Can use additional provided functions and properties to perform on-the-fly. 
    - Real-time output of Log in the Log area of the main interface. 
    - Hang and wait, read user input. 
    - Additional Excel file operations. 
- When a compilation error or a running error occurs, the relevant debugging information will appear in the log area at the bottom of the main interface. 

##### GlobalDic (static class)
```c#
// ---- Save data function ----
// Save data.
void SetObj(string key, object value);
// Check if key is exist.
bool ContainsKey(string key)
// Get data.
object GetObj(string key)
// Reset all data.
void Reset()
```

##### Logger (static class)
```c#
// ---- Output Log function ----
// Depending on the output log type, there will be different coloring distinctions.
void Info(string info);
void Warn(string warn);
void Error(string error);
void Print(string str);

// ---- Whether to warn when a function is not found ----
bool IsOutputMethodNotFoundWarning { get => isOutputMethodNotFoundWarning; set => isOutputMethodNotFoundWarning = value; }
```

##### Scanner (static class)
```c#
// ---- Get the input function ----
// The parameter is the prompt to get the input, and after execution, it will wait until the user enters it.  
// If there are other threads waiting for input, it will first wait for the thread in front to get it, and then execute the content of this statement.  
string GetInput();
string GetInput(string value);

// ---- Waiting for input function ----
// Possibly a useless function. Can wait until user input is obtained while other threads are performing input. 
// Returns the most recent user input. 
string WaitInput();

// ---- Recently entered content ----
string LastInputValue { get => lastInputValue; set => lastInputValue = value; }
```

##### Output (static class)
```c#
// ---- Excel file operations ----
// Create a new excel file
XLWorkbook CreateWorkbook(string name);
// Get an excel file created via CreateWorkbook
XLWorkbook GetWorkbook(string name);
// Get a sheet
IXLWorksheet GetSheet(string workbookName, string sheetName);
IXLWorksheet GetSheet(XLWorkbook workbook, string sheetName);
// Get all excel files created
Dictionary<string, XLWorkbook> GetAllWorkbooks();
// Clear all created excel files
void ClearWorkbooks();

// ---- Whether to save the default output file ----
bool IsSaveDefaultWorkBook { get => isSaveDefaultWorkBook; set => isSaveDefaultWorkBook = value; }
// ---- The location of the output file ----
string OutputPath { get => outputPath; set => outputPath = value; }
```

##### Param (class)
```c#
// Get parameters
List<string> Get(string key);
string GetOne(string key);
// Get a collection of parameter keys
IEnumerable<String> GetKeys();
// Check whether parameter is included
bool ContainsKey(string key);
```

##### Running (static class)
```c#
// ---- operating status ----
// When executing time-consuming logic, it can be judged whether the user has stopped running, so as to exit the logic in time.
bool UserStop { get => userStop; set => userStop = value; }
// Determine whether the current logic is being executed. 
bool NowRunning { get => nowRunning; set => nowRunning = value; }
```

##### RunBeforeAnalyze Function
|Parameter|Type|Description|Remarks|
|----|----|----|----|
|param|Param|Parameters passed in||
|globalObjects|Object|Global existence, can save data that needs to be used in other calls, such as the current line number, etc.||
|allFilePathList|List\<string>|A list of all file paths that will be analyzed||
|globalizationSetter|GlobalizationSetter|Get internationalization string|globalizationSetter.Find("Code");|
|isExecuteInSequence|bool|Whether to execute sequentially||

##### AnalyzeSheet Function
|Parameter|Type|Description|Remarks|
|----|----|----|----|
|param|Param|Parameters passed in||
|sheet|IXLWorksheet|Sheet to be analyzed||
|filePath|string|File path||
|globalObjects|Object|Global existence, can save data that needs to be used in other calls, such as the current line number, etc.||
|globalizationSetter|GlobalizationSetter|Get internationalization string|globalizationSetter.Find("Code");|
|isExecuteInSequence|bool|Whether to execute sequentially||
|invokeCount|int|The number of times this analysis function was called|The value is 1 on the first call|

##### RunBeforeSetResult Function
|Parameter|Type|Description|Remarks|
|----|----|----|----|
|param|Param|Parameters passed in||
|workbook|XLWorkbook|Excel file for output||
|globalObjects|Object|Global existence, can save data that needs to be used in other calls, such as the current line number, etc.||
|allFilePathList|List\<string>|List of all file paths analyzed||
|globalizationSetter|GlobalizationSetter|Get internationalization string|globalizationSetter.Find("Code");|
|isExecuteInSequence|bool|Whether to execute sequentially||

##### SetResult Function
|Parameter|Type|Description|Remarks|
|----|----|----|----|
|param|Param|Parameters passed in||
|workbook|XLWorkbook|Excel file for output||
|filePath|string|File path||
|globalObjects|Object|Global existence, can save data that needs to be used in other calls, such as the current line number, etc.||
|globalizationSetter|GlobalizationSetter|Get internationalization string|globalizationSetter.Find("Code");|
|isExecuteInSequence|bool|Whether to execute sequentially||
|invokeCount|int|The number of times this output function was called|The value is 1 on the first call|
|totalCount|int|The total number of times the output function needs to be called|The last call is when invokeCount is the same as totalCount|

##### RunEnd Function
|Parameter|Type|Description|Remarks|
|----|----|----|----|
|param|Param|Parameters passed in||
|workbook|XLWorkbook|Excel file for output||
|globalObjects|Object|Global existence, can save data that needs to be used in other calls, such as the current line number, etc.||
|allFilePathList|List\<string>|List of all file paths analyzed||
|globalizationSetter|GlobalizationSetter|Get internationalization string|globalizationSetter.Find("Code");|
|isExecuteInSequence|bool|Whether to execute sequentially||

# Open Source Libraries Used
|Open source library|Open source protocol|
|----|----|
|[roslynpad/roslynpad](https://github.com/roslynpad/roslynpad)|MIT|
|[icsharpcode/AvalonEdit](https://github.com/icsharpcode/AvalonEdit)|MIT|
|[JamesNK/Newtonsoft.Json](https://github.com/JamesNK/Newtonsoft.Json)|MIT|
|[ClosedXML/ClosedXML](https://github.com/ClosedXML/ClosedXML)|MIT|
|[~~rickyah/ini-parser~~ rickyah/ini-parser-netstandard](https://github.com/rickyah/ini-parser)|MIT|
|[amibar/SmartThreadPool](https://github.com/amibar/SmartThreadPool)|MS-PL|
|[punker76/gong-wpf-dragdrop](https://github.com/punker76/gong-wpf-dragdrop)|BSD-3-Clause|
|[Kinnara/ModernWpf](https://github.com/Kinnara/ModernWpf)|MIT|
|[CommunityToolkit/WindowsCommunityToolkit (Microsoft.Toolkit.Mvvm)](https://github.com/CommunityToolkit/WindowsCommunityToolkit)|MIT|
|[microsoft/XamlBehaviorsWpf](https://github.com/microsoft/XamlBehaviorsWpf)|MIT|
|[ZjzMisaka/CustomizableMessageBox](https://github.com/ZjzMisaka/CustomizableMessageBox)|WTFPL|
|[ZjzMisaka/RoslynScriptRunner](https://github.com/ZjzMisaka/RoslynScriptRunner)|[OMSPL](https://github.com/ZjzMisaka/OMSPL/blob/main/LICENSE)|
