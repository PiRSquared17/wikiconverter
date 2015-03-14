openwiki to dokuwiki wikiconverter

1， 可以使用本转换器完成初步转换

2， 建议使用powergrep的正则替换组功能来做修正。

Dokuwiki To Mediawiki  wikiconverter

1,利用Mediawiki的页面导出功能先批量导出（把需要导出的页面在列表中列出即可）为一个export.xml文件

2，利用工具把PAGE页面一一添加到一个符合格式的XML文件中，使用UltraEDIT稍经编辑，否则mediawiki-1.18.0会导入异常)就可。

3，利用Mediawiki的页面导入功能批量导入