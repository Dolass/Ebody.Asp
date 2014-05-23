<!--#include file="../../Ebody/ebody.asp"-->

<%
' ===============================================
' 创建基类
' ===============================================
'Dim Ebody : Set Ebody = New Ebody_Base
Ebody.CharSet = "UTF-8"

' ===============================================
' 组构页面
' ===============================================

Dim lvStyleDir
lvStyleDir = Ebody.VHome & "/Core/Ebody111/Demo/tpl/"

' 载入类
Ebody.Use "tpl"

' 设置未定义标签的处理方式
Ebody.Tpl.UnDefine = "keep"

' 载入模板文件
Ebody.Tpl.File = "model/tpl_home.htm"

' 加载各标签值
ebody.tpl.tag("cssfile") = "<link rel=stylesheet type=text/css href=" & lvStyleDir & "style/base.css /><link rel=stylesheet type=text/css href=" & lvStyleDir & "style/home.css />"
ebody.tpl.tag("jsfile") = "<script language=JavaScript type=text/javascript src=" & lvStyleDir & "style/Script/effect.js></script>"

Ebody.tpl.tag("TopicBoxBGImg") = lvStyleDir & "style/image/bar_topic.gif"
Ebody.tpl.tag("ShadowBGImg") = lvStyleDir & "style/image/bg_pattern002.gif"
Ebody.tpl.tag("InfoBandBGImg") = lvStyleDir & "style/image/bar_title.gif"
Ebody.tpl.tag("FooterBarImg") = lvStyleDir & "style/image/bar_footer.png"

Ebody.tpl.tag("WindowTitle") = "Welcome to taoya.com"
Ebody.tpl.tag("onload") =  "onload=ShadowImg('TopicImg')"
Ebody.tpl.tag("TopBar") =  ""
ebody.tpl.tag("TopTitle") =  ""
'Ebody.tpl.tag("PanelLink") =  oRS
ebody.tpl.tag("PanelSearch") =  "<form name=search><input type=text><input type=submit value=GO></form>"
ebody.tpl.tag("TopSign") =  "<img src=" & lvStyleDir & "style/image/TopSign.png>"
Ebody.tpl.tag("TopicImg") =  "<img src=" & lvStyleDir & "style/Image/Topic_Img003.jpg>"
Ebody.tpl.tag("ArticleTitle") =  "<img src=" & lvStyleDir & "style/image/New_Icons001.gif><a href=taoya.com> 新的一年，新的开始，梦想不断延续，终将有所成就，2010年加油！</a>"
Ebody.tpl.tag("ArticleBody") =  ""
Ebody.tpl.tag("SubTitle") =  "Beta 1.111106"
Ebody.tpl.tag("Title") =  "<img border=0 src=" & lvStyleDir & "style/Image/TaoyaLogo.gif>"
ebody.tpl.tag("TopicTitle") =  "今日一品"
Ebody.tpl.tag("TopicSummary") =  "奔驰汽车，是汽车界历史最优久，品牌知名度最高的品牌，无论是从他的汽车产品还是到品牌营销，都称得上是精典。中文品牌“奔驰”也是寓意深远。让人对这个品牌不禁心感崇拜！"
ebody.tpl.tag("TopicLink") =  "<a href=http://www.Mercedes-benz.com>Mercedes-benz.com</a>"
Ebody.tpl.tag("PanelNavBar") =  "<a href=/Brand>品汇</a> | <a href=/Show>品秀</a> | <a href=/Join>加入</a> | <a href=/Gbook>留言</a>"
Ebody.tpl.tag("PanelTitle") =  "块标题"
Ebody.tpl.tag("PanelInner") =  "欢迎来到互联网鉴品地带!在这里,您可以了解到众多优秀品牌及其历史,同时您也可以将您的品牌故事与朋友们分享,让更多的精彩随网络传播..."
'Ebody.tpl.tag("Copyright") =  vCopyright
Ebody.tpl.tag("FooterBar") =  ""

' 更新块内容
Ebody.tpl.UpdateBlock "PanelLink"

' 显示html内容
Ebody.tpl.Show

' 网页存为文件
Ebody.Use "fso"
If Not Ebody.fso.isFile("Files/Home.html") Then 
	Ebody.fso.CreateFile "Files/Home.html", Ebody.tpl.HTML
End If

' 文件生成地址
Response.write "主机网址：" & Ebody.VHome & "<P>"
Response.write "文件网址：" & Ebody.GetUrlAbs("Files/Home.html") & "<P>"
Response.write "文件路径：" & Ebody.GetPathAbs("Files/Home.html") & "<P>"

' 查看生成的文件
Response.write "<a href=" & Ebody.GetUrlAbs("Files/Home.html") & ">查看生成的文件</a>"

' 关闭类
Ebody.Close "tpl"
Ebody.Close "fso"

' ===============================================
' 销毁基类
' ===============================================
Set Ebody = Nothing
%>