# Introduction #

利用正则表达式清除HTML代码中的不需要源码

# skydrive代码清理.pga的powergrep源码 #

注意执行步骤：

１，把skydrive网页转存为带图片的html，放于桌面：temp.htm

２，执行＂skydrive代码清理.pga＂中第1-6的替换一次．

３，把temp.htm的内容使用＂Html to Dokuwiki Convertor version 2007-10-22＂ 工具转换后，把所得转换结果再全部COPY并覆盖temp.htm原来全文．

４，执行＂skydrive代码清理.pga＂中第7-1的替换一次．

５，执行＂skydrive代码清理.pga＂中第12的替换不限次数，直到没有匹配结果为止．

```
<?xml version="1.0" encoding="UTF-8"?>
<pgr:powergrep xmlns:pgr="http://www.powergrep.com/powergrep34.xsd" version="3.4">
	<actionfile>
		<fileselection archives="1" globalmasks="1">
			<drive name="E:">
				<folder name="user">
					<folder name="桌面">
						<file name="temp.htm" marked="1"/>
					</folder>
				</folder>
			</drive>
		</fileselection>
		<action actiontype="replace" searchtype="regex list" concurrent="1" targettype="same" backuptype="multi name" backupdest="same folder">
			<searchtext enabled="0">&lt;!DOCTYPE HTML PUBLIC &quot;-//W3C//DTD HTML 4.01 Transitional//EN&quot; &quot;http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd&quot;>&#13;&#10;[\w\W]*?&#13;&#10;&lt;TABLE class=&quot;gvTable IE6gvTable&quot; id=detailsView cellSpacing=0 cellPadding=0>&#13;&#10;  &lt;THEAD>&#13;&#10;  &lt;TR class=gvTableRow>&#13;&#10;    &lt;TD class=&quot;gvTableCell gvHeader&quot; width=&quot;37%&quot;>&#13;&#10;      &lt;DIV class=overflowEllipsis>名称&lt;/DIV>&lt;/TD>&#13;&#10;</searchtext>
			<searchtext enabled="0">&lt;/DIV>&#13;&#10;&lt;DIV class=clear>&lt;/DIV>&#13;&#10;&lt;DIV class=bpViewPermissionsLink sharingLevel=&quot;Shared&quot;>共享者：&lt;A &#13;&#10;href=&quot;https://cid-a1aa74b5f41380a4.skydrive.live.com/viewpermissions.aspx/soft/1Standard&quot;>我选择的人&lt;/A>&lt;/DIV>&lt;INPUT &#13;&#10;id=postVerb type=hidden name=postVerb> &lt;INPUT id=postVerbData type=hidden &#13;&#10;name=postVerbData> &lt;/DIV>&lt;/FORM>&lt;/DIV>&#13;&#10;[\w\W]*?&lt;/BODY>&lt;/HTML></searchtext>
			<searchtext enabled="0">  \* </searchtext>
			<searchtext enabled="0">title=&quot;([\w\W]*?)&quot;</searchtext>
			<searchtext enabled="0">&lt;IMG class=gvCellOverlayFill &#13;&#10;      alt=&quot;</searchtext>
			<searchtext enabled="0">&amp;#10;修改日期([\w\W]*?)src=&quot;temp.files/transparent.gif&quot;></searchtext>
			<searchtext enabled="0">{{temp\.files\:tinyicons_sprites\.png}}{{temp\.files\:tinyicons_sprites\.png}}</searchtext>
			<searchtext enabled="0">\{\{temp\.files\:</searchtext>
			<searchtext enabled="0">\| 名称 \| 修改日期 \| 类型 \| 大小 \|</searchtext>
			<searchtext enabled="0">png}} </searchtext>
			<searchtext enabled="0"> \[\[http</searchtext>
			<searchtext>\|_([\w\W]*?)\|</searchtext>
			<replacetext>&lt;TABLE class=&quot;gvTable IE6gvTable&quot; id=detailsView cellSpacing=0 cellPadding=0>&#13;&#10;  &lt;THEAD>&#13;&#10;  &lt;TR class=gvTableRow>&#13;&#10;    &lt;TD class=&quot;gvTableCell gvHeader&quot; width=&quot;37%&quot;>&#13;&#10;      &lt;DIV class=overflowEllipsis>名称&lt;/DIV>&lt;/TD>&#13;&#10;</replacetext>
			<replacetext/>
			<replacetext>&#13;&#10;</replacetext>
			<replacetext/>
			<replacetext/>
			<replacetext/>
			<replacetext>{{temp.files:tinyicons_sprites.png}}</replacetext>
			<replacetext>{{</replacetext>
			<replacetext>|  | 名称 | 文件名 | 修改日期 | 类型 | 大小 |</replacetext>
			<replacetext>png}}|</replacetext>
			<replacetext>|[[http</replacetext>
			<replacetext>|</replacetext>
			<sectioning sectiontype="whole file count lines"/>
		</action>
	</actionfile>
</pgr:powergrep>
```