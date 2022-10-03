# ExcelTool
![ICON](https://om.namanime.com/Pictures/ExcelTool/ExcelTool.ico)  
可插件式拓展功能, 可批量处理, 并将结果输出到Excel文件中.  
[English ReadMe](https://github.com/ZjzMisaka/ExcelTool/blob/main/README.md) | [日本語ReadMe](https://github.com/ZjzMisaka/ExcelTool/blob/main/README_ja-JP.md)  
[查看示例](https://github.com/ZjzMisaka/AnalyzersForExcelTool)  
[执行时的gif图](https://www.namanime.com/ZjzMisaka/ExcelTool/ExcelTool.gif?20220603)

### 多语言
- [x] 简体中文
- [x] 日本語
- [x] English

### 主界面
- 通过下拉框选择检索信息和处理逻辑 (插件), 使它们一一对应
- 可以设定传入参数, 设定检索根目录和默认输出目录, 输出文件名
- 有顺序执行与同时执行两种执行方式
- 通过规则下拉框选择保存过的规则可以自动填充以上内容
    - 选择规则后可以设定监视对一些文件夹和文件进行监视, 出现变动后自动执行这项规则  

### 检索信息设定界面
**用于设定需要查找指定路径下指定Excel文件的指定Sheet**
- 查找的方式可以有选择全部, 完整匹配, 部分包含和正则表达式

### 处理逻辑 (插件编码) 界面
**用于设定对某一类Sheet进行的处理逻辑以及处理完毕后的输出逻辑**
- 在编辑器中编写代码, 运行中会依次执行.  
- 可以设定参数, 插件使用者可以在主界面中编辑参数, 并且在运行中传递给代码使用.  

#### 编码相关
- 全程自动补全与着色, 可以自行向Dlls文件夹添加dll文件, 添加后可以直接引用
- 编码内容依赖[ClosedXML](https://github.com/ClosedXML/ClosedXML)开源库
- 可以使用额外提供的函数与属性, 在运行中进行
    - 实时在主界面Log区域输出Log
    - 挂起并等待, 读取用户输入
    - 额外的Excel文件操作
- 当产生编译错误或者运行错误时, 相关调试信息会出现在主界面最下方的log区域中

##### GlobalDic (静态类)
```c#
// ---- 保存数据 ----
// 保存数据
void SetObj(string key, object value);
// 判断存在
bool ContainsKey(string key)
// 获取数据
object GetObj(string key)
// 重置所有
void Reset()
```

##### Logger (静态类)
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

##### Scanner (静态类)
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

##### Output (静态类)
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

##### Param (类)
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
|参数|类型|描述|备注|
|----|----|----|----|
|param|Param|传入的参数||
|globalObjects|Object|全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等||
|allFilePathList|List<string>|将会处理的所有文件路径列表||
|isExecuteInSequence|bool|是否顺序执行||

##### AnalyzeSheet函数
|参数|类型|描述|备注|
|----|----|----|----|
|param|Param|传入的参数||
|sheet|IXLWorksheet|当前被处理的sheet||
|filePath|string|文件路径||
|globalObjects|Object|全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等||
|isExecuteInSequence|bool|是否顺序执行||
|invokeCount|int|此处理函数被调用的次数|第一次调用时值为1|

##### RunBeforeSetResult函数
|参数|类型|描述|备注|
|----|----|----|----|
|param|Param|传入的参数||
|workbook|XLWorkbook|用于输出的Excel文件||
|globalObjects|Object|全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等||
|allFilePathList|List<string>|处理的所有文件路径列表||
|isExecuteInSequence|bool|是否顺序执行||

##### SetResult函数
|参数|类型|描述|备注|
|----|----|----|----|
|param|Param|传入的参数||
|workbook|XLWorkbook|用于输出的Excel文件||
|filePath|string|文件路径||
|globalObjects|Object|全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等||
|isExecuteInSequence|bool|是否顺序执行||
|invokeCount|int|此输出函数被调用的次数|第一次调用时值为1|
|totalCount|int|总共需要调用的输出函数的次数|当invokeCount与totalCount值相同时即为最后一次调用|

##### RunEnd函数
|参数|类型|描述|备注|
|----|----|----|----|
|param|Param|传入的参数||
|workbook|XLWorkbook|用于输出的Excel文件||
|globalObjects|Object|全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等||
|allFilePathList|List<string>|处理的所有文件路径列表||
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
|[Kinnara/ModernWpf](https://github.com/Kinnara/ModernWpf)|MIT|
|[CommunityToolkit/WindowsCommunityToolkit (Microsoft.Toolkit.Mvvm)](https://github.com/CommunityToolkit/WindowsCommunityToolkit)|MIT|
|[microsoft/XamlBehaviorsWpf](https://github.com/microsoft/XamlBehaviorsWpf)|MIT|
|[ZjzMisaka/CustomizableMessageBox](https://github.com/ZjzMisaka/CustomizableMessageBox)|WTFPL|
