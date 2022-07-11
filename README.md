# ExcelTool
可拓展, 批量分析Excel文件, 并将分析结果输出到Excel文件中  
[查看示例](https://github.com/ZjzMisaka/AnalyzersForExcelTool)  
[执行时的gif图](https://www.namanime.com/ZjzMisaka/ExcelTool/ExcelTool.gif?20220603)

### 多语言
- [x] 简体中文
- [x] 日本語
- [x] English

### 主界面
- 设定完成检索信息和分析逻辑后通过下拉框选择, 使一行检索信息和一行分析逻辑一一对应
- 可以设定传入参数, 设定检索根目录和输出目录, 默认输出文件的输出文件名
- 通过规则下拉框选择保存过的规则可以自动填充以上内容
    - 选择规则后可以设定监视对一些文件夹和文件进行监视, 出现变动后自动执行这项规则
- 有顺序执行与同时执行两种执行方式  
<img src="https://www.namanime.com/ZjzMisaka/ExcelTool/ExcelTool01.png?20220603" width="400px" />

### 检索信息分析界面
**用于设定需要查找哪些路径下的哪些Excel文件的哪些Sheet**
- 查找的方式可以有选择全部, 完整匹配, 部分包含和正则表达式
<img src="https://www.namanime.com/ZjzMisaka/ExcelTool/ExcelTool03.png?20220603" width="400px" />

### 逻辑分析界面
**用于设定对某一类Sheet进行怎样的分析, 分析完毕后如何进行输出 (或者如何处理并保留分析的结果供后续使用)**
- 在编辑器中编写代码, 运行中会依次执行.  
<img src="https://www.namanime.com/ZjzMisaka/ExcelTool/ExcelTool02.png?20220603" width="400px" />

#### 编码相关
- 全程自动补全与着色, 可以自行向Dlls文件夹添加dll文件
- 编码内容依赖[ClosedXML](https://github.com/ClosedXML/ClosedXML)开源库
- 可以使用GlobalObjects中提供的函数与属性, 在运行中进行
    - Log的输出
    - 用户输入的读取
    - 额外的Excel文件操作
- 当产生编译错误或者运行错误时, 相关调试信息会出现在主界面最下方的log区域中

##### Logger静态类
```c#
// ---- 输出Log函数 ----
// 根据输出log类型不同, 会有不同的着色区分. 
void Info(string info);
void Warn(string warn);
void Error(string error);
void Print(string str);

// ---- 当找不到函数时是否报出警告属性 ----
bool IsOutputMethodNotFoundWarning { get => isOutputMethodNotFoundWarning; set => isOutputMethodNotFoundWarning = value; }
```

##### Scanner静态类
```c#
// ---- 获取输入函数 ----
// 参数是获取输入的提示语, 执行后会等待直到用户进行输入. 
// 如果有其他线程正在等待输入中, 则会先等待排在前面的线程获取完毕, 再执行此语句的内容.  
string GetInput();
string GetInput(string value);

// ---- 等待输入函数 ----
// 可能是无用函数. 可以在其他线程正在执行输入时等待直到用户输入被获取. 
// 返回最近用户输入的内容. 
string WaitInput();

// ---- 最近输入内容属性 ----
string LastInputValue { get => lastInputValue; set => lastInputValue = value; }
```

##### Output静态类
```c#
// ---- Excel文件操作 ----
// 新建一个excel文件
XLWorkbook CreateWorkbook(string name);
// 获得一个通过CreateWorkbook创建的excel文件
XLWorkbook GetWorkbook(string name);
// 获取一个sheet
IXLWorksheet GetSheet(string workbookName, string sheetName);
IXLWorksheet GetSheet(XLWorkbook workbook, string sheetName);
// 获取创建的所有excel文件
Dictionary<string, XLWorkbook> GetAllWorkbooks();
// 清除创建的所有excel文件
void ClearWorkbooks();

// ---- 是否保存默认输出文件属性 ----
bool IsSaveDefaultWorkBook { get => isSaveDefaultWorkBook; set => isSaveDefaultWorkBook = value; }
```

##### Param类
```c#
// 获取参数
List<string> Get(string key);
string GetOne(string key);
// 获取参数键的集合
IEnumerable<String> GetKeys();
// 判断是否包含参数
bool ContainsKey(string key);
```

##### RunBeforeAnalyze函数
|参数|类型|含义|备注|
|----|----|----|----|
|param|Param|传入的参数||
|globalObjects|Object|全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等||
|allFilePathList|List<string>|将会分析的所有文件路径列表||
|isExecuteInSequence|bool|是否顺序执行||

##### AnalyzeSheet函数
|参数|类型|含义|备注|
|----|----|----|----|
|param|Param|传入的参数||
|sheet|IXLWorksheet|当前被分析的sheet||
|result|ConcurrentDictionary<ResultType, Object>|存储当前文件的信息||
|globalObjects|Object|全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等||
|isExecuteInSequence|bool|是否顺序执行||
|invokeCount|int|此分析函数被调用的次数|第一次调用时值为1|

##### RunBeforeSetResult函数
|参数|类型|含义|备注|
|----|----|----|----|
|param|Param|传入的参数||
|workbook|XLWorkbook|用于输出的Excel文件||
|globalObjects|Object|全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等||
|resultList|ICollection<ConcurrentDictionary<ResultType, Object>>|所有文件的信息||
|allFilePathList|List<string>|分析的所有文件路径列表||
|isExecuteInSequence|bool|是否顺序执行||

##### SetResult函数
|参数|类型|含义|备注|
|----|----|----|----|
|param|Param|传入的参数||
|workbook|XLWorkbook|用于输出的Excel文件||
|result|ConcurrentDictionary<ResultType, Object>|存储当前文件的信息||
|globalObjects|Object|全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等||
|isExecuteInSequence|bool|是否顺序执行||
|invokeCount|int|此输出函数被调用的次数|第一次调用时值为1|
|totalCount|int|总共需要调用的输出函数的次数|当invokeCount与totalCount值相同时即为最后一次调用|

##### RunEnd函数
|参数|类型|含义|备注|
|----|----|----|----|
|param|Param|传入的参数||
|workbook|XLWorkbook|用于输出的Excel文件||
|globalObjects|Object|全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等||
|resultList|ICollection<ConcurrentDictionary<ResultType, Object>>|所有文件的信息||
|allFilePathList|List<string>|分析的所有文件路径列表||
|isExecuteInSequence|bool|是否顺序执行||

# 使用的开源库
|开源库|开源协议|
|----|----|
|[roslynpad/roslynpad](https://github.com/roslynpad/roslynpad)|MIT|
|[icsharpcode/AvalonEdit](https://github.com/icsharpcode/AvalonEdit)|MIT|
|[JamesNK/Newtonsoft.Json](https://github.com/JamesNK/Newtonsoft.Json)|MIT|
|[ClosedXML/ClosedXML](https://github.com/ClosedXML/ClosedXML)|MIT|
|[rickyah/ini-parser](https://github.com/rickyah/ini-parser)|MIT|
|[amibar/SmartThreadPool](https://github.com/amibar/SmartThreadPool)|MS-PL|
|[punker76/gong-wpf-dragdrop](https://github.com/punker76/gong-wpf-dragdrop)|BSD-3-Clause|
|[CommunityToolkit/WindowsCommunityToolkit (Microsoft.Toolkit.Mvvm)](https://github.com/CommunityToolkit/WindowsCommunityToolkit)|MIT|
|[microsoft/XamlBehaviorsWpf](https://github.com/microsoft/XamlBehaviorsWpf)|MIT|
|[ZjzMisaka/CustomizableMessageBox](https://github.com/ZjzMisaka/CustomizableMessageBox)|WTFPL|
