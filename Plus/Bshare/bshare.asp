<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../Plus/Session.asp"-->
<!--#include file="../../Plus/md5.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim Chk:Set Chk=New LoginCheckCls1
Chk.Run()
Set Chk=Nothing
Dim KS:Set KS=New PublicCls

Dim Bshare_Open,Bshare_UUID,Bshare_PassWord,Bshare_RePassWord,Bshare_Domain,Bshare_UserName,Bshare_Secret

Dim Action:Action = LCase(Request("action"))
LoadbshareConfig
Select Case Trim(Action)
	Case "save"		Call savebshare
	Case "show"	    Call show
	Case "getstyle" Call GetStyle
	Case "getdata"  Call GetData
	Case Else		Call showmain
End Select

Sub show
 Response.Write "<script>window.open('http://intf.cnzz.com/user/companion/newasp_login.php?site_id=" & Bshare_UUID & "&password=" & Bshare_password & "');history.back();</script>"
End Sub

Sub ShowMain

If Len(Bshare_Domain)<3 Then Bshare_Domain=KS.GetAutoDomain

Response.Write "<html><head><title>��ϵͳ���Ͻӿ�����</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='../wss/Admin_Style.css' rel='stylesheet' type='text/css'></head>" & vbCrLf
Response.Write "</head>"
Response.Write "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"" scroll=no>"
Response.Write "<ul id='menu_top' style='text-align:center;padding-top:10px;font-weight:bold'>bShare������</ul>"
%>
<script src="../../ks_inc/jquery.js"></script>
<table border="0" align="center" cellpadding="3" cellspacing="1" width="100%" class="border">
<%if bshare_open="true" then%>
<tr class="tdbg">
	<td class="clefttitle" colspan="2" height="30"><strong>���Ѿ���ͨbShare����</strong></td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" width="20%" align="right" height="30"><u>��վ����</u>��</td>
	<td width="80%"><%=Bshare_Domain%></td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" width="20%" align="right" height="30"><u>�û���</u>��</td>
	<td width="80%"><%=Bshare_UserName%></td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" width="20%" align="right" height="30"><u>UUID</u>��</td>
	<td width="80%"><%=Bshare_uuid%></td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" width="20%" align="right" height="30"><u>SECRET</u>��</td>
	<td width="80%"><%=Bshare_secret%></td>
</tr>

<%else%>
<form name="myform" method="post" action="?action=save">
<tr class="tdbg">
	<td class="clefttitle" width="20%" align="right"><u>��վ����</u>��</td>
	<td width="80%"><input type="text" name="Bshare_Domain" size="35" value="<%=Bshare_Domain%>"> 
		<font color="red">* </font>
	</td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" align="right"><u>�˺�����</u>��</td>
	<td><input type="radio" name="Bshare_Type" value="1" onclick="$('#rpass').show();" checked="checked"/>��ע��
	    <input type="radio" name="Bshare_Type" value="2" onclick="$('#rpass').hide();"/>�Ѿ����˺�
	</td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" align="right"><u>�� �� ��</u>��</td>
	<td><input type="text" name="Bshare_UserName" size="35" value="<%=Bshare_UserName%>"> 
		<font color="red">* ��дEmail</font>
	</td>
</tr>

<tr class="tdbg">
	<td class="clefttitle" align="right"><u>��¼����</u>��</td>
	<td><input type="password" name="Bshare_PassWord" size="35" value="<%=Bshare_PassWord%>"> 
		<font color="red">* </font>
	</td>
</tr>

<tr class="tdbg" id="rpass">
	<td class="clefttitle" align="right"><u>ȷ������</u>��</td>
	<td class="clefttitle"><input type="password" name="Bshare_RePassWord" size="35" value=""> 
		<font color="red">* </font>
	</td>
</tr>
<tr class="tdbg">
	<td colspan="2" align="center">
	<input type="submit" value="��������" name="B1" class="Button"></td>
</tr>
</form>
<tr>
	<td class="clefttitle" colspan="2"><b>˵��</b><br/>&nbsp;&nbsp;bShare��ֹ��һ������ť��bShare��ȫ�����Ļ�������ǿ����罻�������棡 ֻ��һ����ť������Ϊ������վע���罻�����ܣ�<br/> 
bShare���ܷ��������������û��������ɵؽ���ϲ�������ݷ����罻��վ��΢��������ѷ����û������뿪������վ�����ܿ��ٵؽ��з����������������վ��
</td>
</tr>
<%end if%>


</table>
<%if bshare_open="true" then%>

<br/>
<table border="0" align="center" cellpadding="3" cellspacing="1" width="100%" class="border">
<tr class="tdbg">
	<td class="clefttitle" colspan="2" height="30"><strong>ǰ̨���ô��룺</strong></td>
</tr>
<tr class="tdbg">
	<td colspan="2">
	������´��븴�Ƶ�����ҳģ��������ʾ�ĵط�����<br/>
	<textarea name="bsharecode" style="width:450px;height:90px"><a class="bshareDiv" href="http://www.bshare.cn/share">����ť</a><script language="javascript" type="text/javascript" src="http://static.bshare.cn/b/button.js#uuid=<%=Bshare_uuid%>&amp;style=2&amp;textcolor=#000&amp;bgcolor=none&amp;bp=qqmb,sinaminiblog,sohubai,renren&amp;ssc=false&amp;sn=true&amp;text=����"></script></textarea>
	
	<div style="margin-top:16px;padding-left:10px">
	  <strong>Ч��Ԥ����</strong><br/>
	  <a class="bshareDiv" href="http://www.bshare.cn/share">����ť</a><script language="javascript" type="text/javascript" src="http://static.bshare.cn/b/button.js#uuid=<%=Bshare_uuid%>&style=2&textcolor=#000&bgcolor=none&bp=qqmb,sinaminiblog,sohubai,renren&ssc=false&sn=true&text=����"></script>
	  </div>
	 
	 Tips:�������������ʽ�����⻹����<input type="button" onclick="getStyle();" class="button" value="��˻�ȡ������ʽ"/> 
	  
	</td>
</tr>
</table>
<script src="../../ks_inc/kesion.box.js"></script>
<script type="text/javascript">
function getStyle(){
   var p=new KesionPopup();
    p.PopupImgDir="/";
	p.PopupCenterIframe('ѡ��Bshare��������ʽ','Bshare.asp?action=GetStyle',720,400,'no')
}
</script>
<%
end if

End Sub

Sub GetStyle()
%>
<style type="text/css">
iframe { border-style: none; }
body { margin: 0px;padding: 0px; }
</style>
<iframe src="http://www.bshare.cn/moreStylesEmbed?uuid=<%=Bshare_UUID%>&bp=qqmb%2csinaminiblog%2csohubai%2cbaiduhi%2crenren%2cbgoogle" name="bshare" width="710px" height="400px" scrolling="yes">
<%
End Sub

Sub GetData()
 if cbool(bshare_open)<>true then
   ks.die "<script>alert('����û�п�ͨ����bshare����ȷ��ת������ҳ��!');location.href='bshare.asp';</script>"
 end if
 Dim TS:TS=ToUnixTime(now,8)&"000"
 Dim Sign:Sign=md5("ts=" & ts & "uuid=" & bshare_uuid & bshare_secret,32)
%>
<style type="text/css">
iframe { border-style: none; }
body { margin: 0px;padding: 0px; }
</style>
<iframe src="http://www.bshare.cn/publisherStatisticsEmbed?uuid=<%=bshare_uuid%>&ts=<%=ts%>&sig=<%=sign%>" name="bshare" style="width:100%;height:100%" width="800" height="600" scrolling="yes">
<%
End Sub

Sub savebshare()
	If Len(Request.Form("Bshare_domain")) < 3 Then
		response.write "<script>alert('�����������!');history.back();</script>"
	End If
	Dim XmlDoc,XmlNode,Xml_Files
	Dim Bshare_Type : Bshare_Type = KS.ChkClng(KS.G("Bshare_Type"))
	Xml_Files = "bshare.config"
	Xml_Files = Server.MapPath(Xml_Files)
	Set XmlDoc = KS.InitialObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	If XmlDoc.Load(Xml_Files) Then
		Set XmlNode = XmlDoc.documentElement.selectSingleNode("rs:data/z:row[@id=0]")
		'If Bshare_Type = 2 Then
		'	XmlNode.attributes.getNamedItem("Bshare_UUID").text = KS.S("Bshare_UUID")
		'	XmlNode.attributes.getNamedItem("Bshare_password").text = KS.S("Bshare_password")
		'Else
			If Len(Request.Form("Bshare_domain")) > 3 Then
				Dim strbshareData
				Dim strURL,strDomain,strKey
				Bshare_domain = KS.G("Bshare_domain")
				Bshare_UserName=Request.Form("Bshare_UserName")
				Bshare_PassWord=Request.Form("Bshare_PassWord")
				Bshare_RePassWord=Request.Form("Bshare_RePassWord")
				If Bshare_UserName="" Then KS.Die "<script>alert('�����������û���!');history.back();</script>"
				If Bshare_PassWord="" Then KS.Die "<script>alert('�������¼����!');history.back();</script>"
				If Bshare_Type <>2 and Bshare_PassWord<>Bshare_RePassWord Then KS.Die "<script>alert('������������������벻һ��!');history.back();</script>"
				
				strURL = "http://api.bshare.cn/analytics/reguuid.json?email="  & Bshare_UserName & "&password=" & Bshare_PassWord & "&domain=" & Bshare_domain & "&source=kesion"
				strbshareData = GetbshareData(strURL)
				
				
				If InStr(strbshareData,"{""uuid"":""") > 0 Then
					Dim bshareArray
					bshareArray = Split(strbshareData, ",")
					XmlNode.attributes.getNamedItem("bshare_uuid").text = trim(replace(replace(bshareArray(0),"{""uuid"":""",""),"""",""))
					XmlNode.attributes.getNamedItem("bshare_secret").text = trim(replace(replace(bshareArray(1),"""secret"":""",""),"""}",""))
					XmlNode.attributes.getNamedItem("bshare_domain").text = Bshare_domain
					XmlNode.attributes.getNamedItem("bshare_password").text = Bshare_password
					XmlNode.attributes.getNamedItem("bshare_username").text = Bshare_username
					XmlNode.attributes.getNamedItem("bshare_open").text = "true"
				Else
					Response.Write "<script>alert('����bshareʧ��!������룺" & strbshareData  &"');history.back();</script>"
					Exit Sub
				End If
			End If
		'End If
		XmlDoc.save Xml_Files
		Set XmlNode = Nothing
	End If
	Set XmlDoc = Nothing
	 Response.Write "<script>alert('��ϲ�������뿪ͨbshare�ɹ���');location.href='bshare.asp';</script>"
End Sub
'����ʱ��� 
Function ToUnixTime(strTime, intTimeZone)
If IsEmpty(strTime) or Not IsDate(strTime) Then strTime = Now
If IsEmpty(intTimeZone) or Not isNumeric(intTimeZone) Then intTimeZone = 0
ToUnixTime = DateAdd("h",-intTimeZone,strTime)
ToUnixTime = DateDiff("s","1970-1-1 0:0:0", ToUnixTime)
End Function

Function GetbshareData(ByVal strURL)
	On Error Resume Next
	Dim xmlhttp,TextBody
	Set xmlhttp = KS.InitialObject("msxml2.ServerXMLHTTP")
	xmlhttp.setTimeouts 65000, 65000, 65000, 65000
	xmlhttp.Open "GET",strURL,false
	xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xmlhttp.send()
	'TextBody = strAnsi2Unicode(xmlhttp.responseBody)
	TextBody = xmlhttp.responseText
	Set xmlhttp = Nothing
	GetbshareData = TextBody
End Function
Function strAnsi2Unicode(asContents)
	Dim len1,i,varchar,varasc
	strAnsi2Unicode = ""
	len1=LenB(asContents)
	If len1=0 Then Exit Function
	  For i=1 to len1
	  	varchar=MidB(asContents,i,1)
	  	varasc=AscB(varchar)
	  	If varasc > 127  Then
	  		If MidB(asContents,i+1,1)<>"" Then
	  			strAnsi2Unicode = strAnsi2Unicode & chr(ascw(midb(asContents,i+1,1) & varchar))
	  		End If
	  		i=i+1
	     Else
	     	strAnsi2Unicode = strAnsi2Unicode & Chr(varasc)
	     End If	
	  Next
End Function
Sub LoadbshareConfig()
Dim XmlDoc,XmlNode,Xml_Files
Xml_Files = "bshare.config"
Xml_Files = Server.MapPath(Xml_Files)
Set XmlDoc = KS.InitialObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
If Not XmlDoc.Load(Xml_Files) Then
			Bshare_Open = true
			Bshare_UUID = ""
			Bshare_PassWord = ""
			Bshare_Domain = KS.GetAutoDomain
			Bshare_UserName = ""
Else
			Set XmlNode	= XmlDoc.documentElement.selectSingleNode("rs:data/z:row[@id=0]")
			Bshare_Open = XmlNode.getAttribute("bshare_open")
			Bshare_UUID = XmlNode.getAttribute("bshare_uuid")
			Bshare_SECRET= XmlNode.getAttribute("bshare_secret")
			Bshare_UserName=XmlNode.getAttribute("bshare_username")
			Bshare_PassWord = XmlNode.getAttribute("bshare_password")
			Bshare_Domain = XmlNode.getAttribute("bshare_domain")
			Bshare_UserName = XmlNode.getAttribute("bshare_username")
			Set XmlNode = Nothing
End If
Set XmlDoc = Nothing
End Sub
%>