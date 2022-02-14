# ExcelTool
批量分析Excel文件, 并将分析结果输出到Excel文件中
### 主界面
- 设定完成检索信息和分析逻辑后通过下拉框选择, 使一行检索信息和一行分析逻辑一一对应
- 可以设定传入参数
- 设定检索根目录和输出目录, 输出文件名
- 通过规则下拉框选择设定好的规则可以自动填充以上内容
    - 选择规则后可以设定监视对一些文件夹和文件进行监视, 出现变动后自动执行这项规则
- 点击开始可以手动执行
<img src="https://www.namanime.com/ZjzMisaka/ExcelTool/ExcelTool01.png?20220215" width="400px" />

### 检索信息分析界面
**用于设定需要查找哪些路径下的哪些Excel文件的哪些Sheet**
- 查找的方式可以有完整匹配, 部分包含和正则表达式
- 设定完成后可保存以便后续使用
<img src="https://www.namanime.com/ZjzMisaka/ExcelTool/ExcelTool03.png?20220215" width="400px" />

### 逻辑分析界面
**用于设定对某一类Sheet进行怎样的分析, 分析完毕后如何进行输出 (或者如何处理并保留分析的结果供后续使用)**
- 在编辑器中编写代码, 在分析阶段程序会调用AnalyzeSheet函数, 在输出阶段会调用SetResult函数. 
- 设定完成后可保存以便后续使用
<img src="https://www.namanime.com/ZjzMisaka/ExcelTool/ExcelTool02.png?20220215" width="400px" />

#### 编码相关
- 编码内容依赖[ClosedXML](https://github.com/ClosedXML/ClosedXML)开源库 **(支持自动补全等功能)**
- 当产生编译错误或者运行错误时, 相关调试信息会出现在主界面最下方的log区域中

##### RunBeforeAnalyze函数
|参数|类型|含义|备注|
|----|----|----|----|
|paramDic|Dictionary<string, string>|传入的参数||
|globalObjects|Object|全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等||

##### AnalyzeSheet函数
|参数|类型|含义|备注|
|----|----|----|----|
|paramDic|Dictionary<string, string>|传入的参数||
|sheet|IXLWorksheet|当前被分析的sheet||
|result|ConcurrentDictionary<ResultType, Object>|存储当前文件的信息||
|globalObjects|Object|全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等||
|invokeCount|int|此分析函数被调用的次数|第一次调用时值为1|

##### RunBeforeSetResult函数
|参数|类型|含义|备注|
|----|----|----|----|
|paramDic|Dictionary<string, string>|传入的参数||
|workbook|XLWorkbook|用于输出的Excel文件||
|globalObjects|Object|全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等||

##### SetResult函数
|参数|类型|含义|备注|
|----|----|----|----|
|paramDic|Dictionary<string, string>|传入的参数||
|workbook|XLWorkbook|用于输出的Excel文件||
|result|ConcurrentDictionary<ResultType, Object>|存储当前文件的信息||
|globalObjects|Object|全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等||
|invokeCount|int|此输出函数被调用的次数|详见上表|
|totalCount|int|总共需要调用的输出函数的次数|当invokeCount与totalCount值相同时即为最后一次调用|

# 使用的开源库
|开源库|开源协议|
|----|----|
|[roslynpad/roslynpad](https://github.com/roslynpad/roslynpad)|MIT|
|[icsharpcode/AvalonEdit](https://github.com/icsharpcode/AvalonEdit)|MIT|
|[JamesNK/Newtonsoft.Json](https://github.com/JamesNK/Newtonsoft.Json)|MIT|
|[ClosedXML/ClosedXML](https://github.com/ClosedXML/ClosedXML)|MIT|
|[rickyah/ini-parser](https://github.com/rickyah/ini-parser)|MIT|
|[amibar/SmartThreadPool](https://github.com/amibar/SmartThreadPool)|MS-PL|
|[ZjzMisaka/CustomizableMessageBox](https://github.com/ZjzMisaka/CustomizableMessageBox)|WTFPL|

### 注意
CustomizableMessageBox不使用Nuget, 需要从库的release中手动下载dll文件. 
