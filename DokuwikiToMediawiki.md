# 概述 #

1,利用Mediawiki的页面导出功能先批量导出（把需要导出的页面在列表中列出即可）为一个export.xml文件

2，利用工具把PAGE页面一一添加到一个符合格式的XML文件中，使用UltraEDIT稍经编辑(把`</siteinfo>和第一个<page>`之间按回车换行，否则mediawiki-1.18.0会导入异常)就可。

3，利用Mediawiki的页面导入功能批量导入


# 步骤 #

难点:
  * 把DW中的PAGE的TXT文件名（先DEURL再）DeUtf8；难点在于DW中一些特殊字符如♯۰等无法解码，必须把相应的文件名改名后再进行，否则程序会死机。而且不能把文件的路径一起带进来解码，否则会解码无效
  * dw中的URL是不区分英文字母大小写的，而mw区分，所以需要另外转换。


参见
  * [VB6](VB6.md)支持UTF文本文件访问的模块 支持UTF-8无BOM格式编码自动识别 http://www.cnblogs.com/shenhaocn/archive/2011/10/23/2221572.html
  * [VB6](VB6.md)支持UTF文本文件访问的模块  http://blog.csdn.net/zyl910/article/details/762693
  * UTF文本文件访问  http://www.mndsoft.com/blog/VB6/0896.html
  * VB 读取UTF-8编码文件函数 http://www.130bbs.cn/read-htm-tid-3838837.html