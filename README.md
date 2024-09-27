# ACAD-VBA-函数库

#### 介绍
适用于AutoCAD二次开发的VBA函数库

#### 开发前的准备工作

1.  下载并安装[AutoCAD VBA模块](https://www.autodesk.com.cn/support/technical/article/caas/tsarticles/tsarticles/CHS/ts/3kxk0RyvfWTfSfAIrcmsLQ.html)
2.  下载并安装具有代码格式化功能的VBE插件，如[VBE2019](https://club.excelhome.net/thread-1461076-1-1.html?_dsign=0fd6df83)
3.  打开AutoCAD，按Alt+F11进入VBE，点击工具-选项，在“编辑器”选项卡中取消勾选“自动语法检测”（否则，当代码中有语法错误时，编辑器不仅会将代码标红，还会弹出警告对话框），在“编辑器格式”选项卡中调整代码颜色、字体、字号。如果对话框内容显示不全，可以在Word中打开VBE进行修改
#### 使用说明

1.  将acaddoc.lsp添加到AutoCAD的支持文件搜索路径里
2.  点击文件-导入文件，导入Universal.bas和ACAD_Only.bas两个文件。Universal.bas中的函数可以在AutoCAD、Word、Excel等各种支持VBA环境的软件中运行，ACAD_Only.bas中的函数只能在AutoCAD中运行
3.  导入example文件夹中的example.bas，可查看其中的示例。其中的例子改编自科学出版社《AutoCAD完全应用指南》（2011年4月第一版）中的示例，书中的示例为AutoLISP程序，文件中改为用VBA实现
4.  点击插入-模块，在新插入的模块中写自己的代码
5.  程序的基本结构：
```basic
Sub c_name() '以“c_”开头的函数会被注册为AutoCAD命令（类似于AutoLISP中以“c:”开头的函数）
    vba_start '初始化vba函数库
    '业务代码放在这里
End Sub
```
6.  按Ctrl+J可以触发代码补全功能
7.  如果要将写好的VBA程序复制给别人，对方电脑上也必须安装[AutoCAD VBA模块](https://www.autodesk.com.cn/support/technical/article/caas/tsarticles/tsarticles/CHS/ts/3kxk0RyvfWTfSfAIrcmsLQ.html)才能运行
#### 函数及全局变量列表
所有返回值是对象的函数均以“o”开头，在将函数的返回值赋值给变量时，需要在变量名前加Set关键字
##### Universal.bas
###### 全局变量
pi，值为3.14159265358979
###### 函数
容器：
|函数|描述|
|-|-|
|oCollection(value1, value2, ...)|初始化一个Collection对象|
|CollToArr(coll As Collection)|Collection转Array|
|oList(value1, value2, ...)|初始化一个[ArrayList对象](https://learn.microsoft.com/zh-cn/dotnet/api/system.collections.arraylist?view=netframework-3.5)（需要在“启用或关闭Windows功能”中勾选“.NET Framework 3.5”）|
|oDict(key1, value1, key2, value2, ...)|初始化一个字典对象|
|Len1(item)|通用的Length函数，可以获取字符串、数组的长度和对象的Count属性|
|Slc(item, start, end可选)|切片(Slice)函数，可以处理字符串、数组、Collection和ArrayList|

字符串：
|函数|描述|
|-|-|
|fmt(str, param1, param2, ...)|字符串格式化，使用“{}”作为占位符，如：fmt("衬衫的价格是：\n{}镑{}便士", 9, 15)，字符串中的"\n"、"\r"、"\r\n"、"\t"会被转义|
|oRegExp(Pattern, Global1可选, IgnoreCase可选, Multiline可选)|创建一个正则表达式对象|

互操作：
|函数|描述|
|-|-|
|oExcelFunc()|创建一个[Excel WorksheetFunction](https://learn.microsoft.com/zh-cn/office/vba/api/excel.worksheetfunction)，从而实现在Excel以外的程序中调用Excel函数|
|ShellOut(command, argv)|通过shell调用其他程序,并获取返回值|

文件系统：
|函数|描述|
|-|-|
|GetFolder(msg)|弹出选择文件夹对话框|
|oFileOpen(f_path, mode)|返回一个[TextStream对象](https://learn.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/textstream-object)，用于读写文件|
##### ACAD_Only.bas
###### 全局变量
dwg，即ThisDrawing
msp，即ThisDrawing.ModelSpace
uty，即ThisDrawing.Utility
###### 函数
开发工具：
|函数|描述|
|-|-|
|vba_start()|初始化vba函数库|
|defun()|立即将已加载的VBA工程中以“c_”开头的函数注册为AutoCAD命令|
|SendCmd(param1, param2, ...)或SendCmd("param1 param2 ...")|在AutoCAD命令行中执行命令|
|TypedArr(value1, value2, ...)|创建数组。用法与VBA自带的[Array函数](https://learn.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/array-function)相同，但[Array函数](https://learn.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/array-function)创建出的数组为Variant()类型，而很多AutoCAD函数对参数的数据类型有明确要求，传入Variant()会报错。而TypedArr创建的数组具有明确的数据类型，如Double()|
|ToTypedArr(object)|将Collection和ArrayList转化为有明确数据类型的数组|

几何：
|函数|描述|
|-|-|
|GetMid(point1, point2)|获取两点连线的中点|
|rc(point, dx, dy, dz可选)|通过“基点”+“沿x、y、z向偏移量”的方式获得下一点|
|Pstr(point)|将以数组形式保存的点转为字符串形式，供command使用|
|GetBlockAttribute(obj, name)|获取动态块的自定义属性|
|SetBlockAttribute(obj, name, value)|设置动态块的自定义属性|
|GetBlockProperty(obj, name)|获取动态块的自定义参数|
|SetBlockProperty(obj, name, value)|设置动态块的自定义参数|
|GetEntity(prompt可选)|由用户在屏幕上选择一个图元，如果没有选中任何图元，会询问用户是否继续选择，而不是直接取得Nothing。返回值为一个Variant()数组，数组的0号元素为选中的图元，1号元素为用户点击的位置|

文件系统：
|函数|描述|
|-|-|
|GetFile(msg, default_path, extension, mode)|弹出选择文件对话框|

绘图：
|函数|描述|
|-|-|
|AddRec2d(point1, point2)|绘制矩形|
|AddDimLinears(distance, point1, point2, point3, ...)|绘制线性标注|

