# ExcelTool
Excel批量分析工具
### 主界面
- 设定完成检索信息和分析逻辑后通过下拉框选择, 使一行检索信息和一行分析逻辑一一对应
- 设定检索根目录和输出目录, 输出文件名
- 点击开始
![主界面](https://www.iaders.com/wp-content/uploads/2021/08/1.png)
### 检索信息分析界面
**用于设定需要查找哪些路径下的哪些Excel文件的哪些Sheet**
- 查找的方式可以有完整匹配, 部分包含和正则表达式
![检索信息分析界面](https://www.iaders.com/wp-content/uploads/2021/08/2.png)
### 逻辑分析界面
**用于设定对某一类Sheet进行怎样的分析, 分析完毕后如何进行输出 (或者如何处理并保留分析的结果供后续使用)**
- 当产生编译错误或者运行错误时, 相关调试信息会出现在主界面最下方的log区域中
![逻辑分析界面](https://www.iaders.com/wp-content/uploads/2021/08/3.png)

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
