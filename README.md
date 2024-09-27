# ACAD-VBA-函数库

#### 介绍
适用于AutoCAD二次开发的VBA函数库

#### 开发前的准备工作

1.  下载并安装[AutoCAD VBA模块](https://www.autodesk.com.cn/support/technical/article/caas/tsarticles/tsarticles/CHS/ts/3kxk0RyvfWTfSfAIrcmsLQ.html)
2.  下载并安装具有代码格式化功能的VBE插件，如[VBE2019](https://club.excelhome.net/thread-1461076-1-1.html?_dsign=0fd6df83)
3.  打开AutoCAD，按Alt+F11进入VBE，点击工具-选项-编辑器格式，调整代码颜色、字体、字号。如果对话框内容显示不全，可以在Word中打开VBE进行修改

#### 使用说明

1.  将acaddoc.lsp添加到AutoCAD的支持文件搜索路径里
2.  点击文件-导入文件，导入Universal.bas和ACAD_Only.bas两个文件。Universal.bas中的函数可以在AutoCAD、Word、Excel等各种支持VBA环境的软件中运行，ACAD_Only.bas中的函数只能在AutoCAD中运行
3.  导入example文件夹中的example.bas，可查看其中的示例。其中的例子改编自科学出版社《AutoCAD完全应用指南》（2011年4月第一版）中的示例，书中的示例为AutoLISP程序，文件中改为用VBA实现
4.  点击插入-模块，在新插入的模块中写自己的代码
  



