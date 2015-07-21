<!--#include file="../conn.asp"-->
<!--#include file="../ks_cls/kesion.commoncls.asp"-->
<!--#include file="../ks_cls/kesion.label.commoncls.asp"-->
<%
Const CacheTime=360   '缓存更新时间,单位秒
Dim KS:Set KS=New PublicCls
Dim Node,DateStr,Title
If Not IsObject(Application(KS.SiteSN & "NewFreshXML")) or (DateDiff("s",Application(KS.SiteSn &"NewFreshTime"),Now)>=CLng(CacheTime)) Then
	Application(KS.SiteSn &"NewFreshTime")=Now
    LoadNewData
End If
%>
var zzi = 0;
function showzzmb() {
if(zzi== 0) {
$("#zzmb_"+ zzi).hide();
zzi++;
$("#zzmb_"+ zzi).show();

} else {
$("#zzmb_"+ zzi).hide();
zzi = 0;
$("#zzmb_"+ zzi).show();
}
$("#zzmbpage").html(zzi + 1 +"/2");
}
<%
ShowData

Sub ShowData()
If IsObject(Application(KS.SiteSN & "NewFreshXML")) Then
  Dim KSR,UserFace,username,Url,n:n=0
  	
	KS.Echo "document.write('<div class=""zzblog"">');" &vbcrlf
  Set KSR=New Refresh
  For Each Node In Application(KS.SiteSN & "NewFreshXML").DocumentElement.SelectNodes("row")
    UserFace=Node.SelectSingleNode("@userface").text
	If KS.IsNul(UserFace) Then UserFace="images/face/boy.jpg"
	If left(userface,1)<>"/" and lcase(left(userface,4))<>"http" then userface="../" & userface
    DateStr=Node.SelectSingleNode("@adddate").text
	Title=KS.LoseHtml(Node.SelectSingleNode("@content").text)
	If len(Title)>21 Then Title=left(title,21) & "..."
	Title=Replace(Replace(KSR.ReplaceEmot(Title),"'","\'"),chr(10),"<br/>")
	N=n+1
	If N=1 Then 
	 KS.Echo "document.write('  <ul id=""zzmb_0"">');" &vbcrlf
	ElseIf N=6 Then
	 KS.Echo "document.write('  </ul><ul id=""zzmb_1"" style=""display:none"">');" &vbcrlf
	End If
	username=Node.SelectSingleNode("@realname").text
	If KS.IsNul(userName) Then UserName=Node.SelectSingleNode("@username").text
    Url=KS.GetSpaceUrl(Node.SelectSingleNode("@userid").text)
	Dim Url1
	If KS.SSetting(21)="1" Then
	 Url1=KS.GetDomain & "space/list-" &Node.SelectSingleNode("@userid").text& "-" & Node.SelectSingleNode("@id").text &KS.SSetting(22)
	Else
	 Url1=Url & "/log/" & Node.SelectSingleNode("@id").text 
	End If
            KS.Echo "document.write('<li><span><a href="""& url &""" target=""_blank""><img src=""" & UserFace & """ onerror=""this.onerror=null;this.src=\'../images/face/boy.jpg\';"" title=""" & username & """></a><br/><a href=""" & url & """ target=""_blank"" title=""" & username & """>" & Left(UserName,6) & "</a></span><div class=""zzlist""><p>" & Title & "</p></div><div class=""zzjh""><a href=""" & url & """ target=""_blank"">关注他</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=\'" & Url1 & "\' target=\'_blank\'><font style=""color:#999;"">评论</font> (" & Node.SelectSingleNode("@totalput").text & ")</a></div></li>');"&vbcrlf
        
	
    'KS.Echo "document.write('<li><a href=""" & KS.GetClubShowURL(Node.SelectSingleNode("@id").text)&""" title=""" & title & """ target=""_blank"">" & KS.Gottopic(Title,25) & "</a>(<font color=green>" & right("0"&hour(DateStr),2) & ":" &right("0"&minute(DateStr),2) & "</font>)</li>')"&vbcrlf
  Next
  KS.Echo "document.write('</ul>');"&vbcrlf
  KS.Echo "document.write('</div>');"&vbcrlf
  Set KSR=Nothing
  If N>5 Then
   KS.Echo "document.write('<div class=""fypage"" style=""width:110px;clear:both"">');" &vbcrlf
   KS.Echo "document.write('<a href=""javascript:;"" class=""n_link"" title=""后一页"" onclick=""showzzmb();""><em>后一页</em></a> <a href=""javascript:;"" class=""p_link"" title=""前一页"" onclick=""showzzmb();""><em>前一页</em></a><font style=""float:right; margin-right:5px; color:#999; font-size:10px;"" id=""zzmbpage"">1/2</font></div>');" &vbcrlf
  End If
End If
End Sub

Sub LoadNewData()
	Dim SQL,RS
	SQL="select top 10 l.*,userface,u.realname from ks_bloginfo l inner join ks_user u on l.username=u.username where l.istalk=1 and l.status=0 order by id desc"
	Set RS=Server.CreateObject("adodb.recordset")
	RS.Open SQL,Conn,1,1
	If Not RS.Eof Then
	  Set Application(KS.SiteSN & "NewFreshXML")=KS.RsToXml(rs,"",row)
	End If
	RS.Close :Set RS=Nothing
End Sub
%>