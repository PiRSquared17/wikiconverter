##### 从Openwiki迁移到Dokuwiki #####

1，下载Openwiki的旧数据库OpenWikiDist.mdb，先清理openwiki\_revisions表，把所有wrv\_current为0的记录清楚，压缩恢复数据库

2，下载[[to dokuwiki wikiconverter](http://code.google.com/p/wikiconverter/|openwiki)]

3，在当前创建page目录，运行wikiconverter，选择数据库OpenWikiDist.mdb，选择表wrv\_current，点：开始转换。然后点：转换内码＝》添加目录，选择所创建的page目录，把“保留文件备份”前的勾去掉，点：开始处理。

4，把page目录中的所有文件copy到Dokuwiki的\data\pages目录下。完毕。

(需要VB远行环境[http://www.52z.com/soft/14178.html｜COMDLG32.OCX])

##### 从Google Site(原jotspot)迁移到Dokuwiki #####

1， 下载或安装好：[http://code.google.com/p/wikiconverter/|Html2DokuWiki.exe]，Powergrep,office word,notepad++

2， 登录Google Site(使用Safari浏览器比较快)，点编辑页面＝》html，全选，运行Html2DokuWiki.exe进行转换，然后使用office word进行修正。(具体为：\\替换为<sup>p</sup>p，“<sup>p后面加一个空格”反复替换为</sup>p<sup>p，</sup>p<sup>p</sup>p反复替换为<sup>p</sup>p)

3， 对于Google Site的上传附件，

可点显示页面源文件，
选取“
```
<div jotId="goog-attachment-inner" style="">
```
”和“添加附件:
```
<input contentEditable="false" onclick="this.blur()" name="userfile" type="file" onchange="return JOT_ATTACH_handleUploadXfer()" />
<input type="hidden" name="pagePath" value="/Home/projects/sitebuild/pagedesign" /></p></form></div>
```
”之间的内容即附件所在源html，
使用notepad++，把所有“
```
a href="/site/
```
”替换为“
```
a href="https://sites.google.com/site/
```
”，
然后使用Powergrep进行正则式替换(可下载[http://code.google.com/p/wikiconverter/|wiki附件代码清理.pga])： “
```
- 创建时间[\w\W]*?删除]]
```
”删空，“
```
  \* 
```
”替换为一个换行符(相当于word里面的^p)