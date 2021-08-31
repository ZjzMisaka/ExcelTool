# ExcelTool
批量分析Excel文件, 并将分析结果输出到Excel文件中
### 主界面
- 设定完成检索信息和分析逻辑后通过下拉框选择, 使一行检索信息和一行分析逻辑一一对应
- 设定检索根目录和输出目录, 输出文件名
- 点击开始
<img src="https://www.iaders.com/wp-content/uploads/2021/08/1.png" width="400px" />
上图: 使用单元测试1的分析和输出代码处理Job, 账单, 界面文件夹的单元测试1测试报告文件中的指定Sheet, 使用单元测试2的分析和输出代码处理Job, 界面, 公共文件夹的单元测试2测试报告文件中的指定Sheet, 输出结果到桌面的输出.xlsx

### 检索信息分析界面
**用于设定需要查找哪些路径下的哪些Excel文件的哪些Sheet**
- 查找的方式可以有完整匹配, 部分包含和正则表达式
- 设定完成后可保存以便后续使用
<img src="https://www.iaders.com/wp-content/uploads/2021/08/2.png" width="400px" />

### 逻辑分析界面
**用于设定对某一类Sheet进行怎样的分析, 分析完毕后如何进行输出 (或者如何处理并保留分析的结果供后续使用)**
- 在编辑器中编写代码, 在分析阶段程序会调用AnalyzeSheet函数, 在输出阶段会调用SetResult函数. 
- 设定完成后可保存以便后续使用
<img src="https://www.iaders.com/wp-content/uploads/2021/08/3.png" width="400px" />

#### 编码相关
- 编码内容依赖[ClosedXML](https://github.com/ClosedXML/ClosedXML)开源库
- 当产生编译错误或者运行错误时, 相关调试信息会出现在主界面最下方的log区域中

##### AnalyzeSheet函数
|参数|类型|含义|备注|
|----|----|----|----|
|sheet|IXLWorksheet|当前被分析的sheet||
|result|ConcurrentDictionary<ResultType, Object>|存储当前分析的结果, 初始存有文件路径 (ResultType.FILEPATH), 文件名 (ResultType.FILENAME), (报错) 消息 (ResultType.MESSAGE), 用户可通过ResultType.RESULTOBJECT键存入自己的东西|当此Sheet对应的SetResult函数调用时, 会传入此变量|
|globalObjects|Object|全局存在的变量, 当有需要和其它Sheet共享的数据或保存当前行号等全局变量的需求时可保存至此|每次调用AnalyzeSheet或SetResult函数时都会传入此变量上次保存的结果. 通过GlobalObjects.GlobalObjects.SetGlobalParam(globalObjects);可进行保存|
|invokeCount|int|此函数被调用的次数|第一次调用时值为1|

##### SetResult函数
|参数|类型|含义|备注|
|----|----|----|----|
|workbook|XLWorkbook|用于输出的Excel文件||
|result|ConcurrentDictionary<ResultType, Object>|存储某个Sheet分析的结果|详见上表|
|globalObjects|Object|全局存在的变量|详见上表|
|invokeCount|int|此函数被调用的次数|详见上表|
|totalCount|int|总共需要调用的输出函数的次数|当invokeCount与totalCount值相同时即为最后一次调用|

# 使用的开源库
|开源库|开源协议|
|----|----|
|[icsharpcode/AvalonEdit](https://github.com/icsharpcode/AvalonEdit)|MIT|
|[JamesNK/Newtonsoft.Json](https://github.com/JamesNK/Newtonsoft.Json)|MIT|
|[ClosedXML/ClosedXML](https://github.com/ClosedXML/ClosedXML)|MIT|
|[rickyah/ini-parser](https://github.com/rickyah/ini-parser)|MIT|
|[amibar/SmartThreadPool](https://github.com/amibar/SmartThreadPool)|MS-PL|
|[ZjzMisaka/CustomizableMessageBox](https://github.com/ZjzMisaka/CustomizableMessageBox)|WTFPL|
### 注意
CustomizableMessageBox不使用Nuget, 需要从库的release中手动下载dll文件. 
