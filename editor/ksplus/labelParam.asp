<!--#include file="../../conn.asp"-->
<%
Dim Action:Action=Request("Action")
If Action="" Then Response.End()
Select Case Lcase(Action)
  case "ad" ad
  case "artphoto" ArtPhoto
  case "downphoto" DownPhoto
  case "flash" Flash
  case "flashplayer" flashPlayer
  case "getmusiclist" GetMusicList
  case "getspeciallist" GetSpecialList
  case "logo" Logo
  case "tags" tags
  case "moviedown" MovieDown
  case "moviepage" MoviePage
  case "moviephoto" MoviePhoto
  case "movieplay" MoviePlay
  case "productgroupphoto" ProductGroupPhoto
  case "productphoto" ProductPhoto
  case "status1" Status1
  case "status2" Status2
  case "status3" Status3
  case "supplyphoto" SupplyPhoto
  case "topuser" TopUser
  case "userdynamic" UserDynamic
End Select
%>
<%Sub Ad()%>
<html>
<head>
<title>�����������������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
	if (document.myform.leftsrc.value=='')
	  {
	   alert('������������ַ!')
	   document.myform.leftsrc.focus();
	   return false;
	  }
	  if (document.myform.rightsrc.value=='')
	  {
	   alert('������������ַ!')
	   document.myform.rightsrc.focus();
	   return false;
	  }
	 if (document.myform.closesrc.value=='')
	  {
	   alert('������ر�ͼ���ַ!')
	   document.myform.closesrc.focus();
	   return false;
	  }
    Val = '{=JS_Ad("'+document.myform.leftsrc.value+'","'+document.myform.rightsrc.value+'","'+document.myform.closesrc.value+'",'+document.myform.speed.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
 
<link href="editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>�����������</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td align="right"><div align="center">����Flash��ַ</div></td>
    <td ><input name="leftsrc" type="text" id="leftsrc" size="60">
      100*250</td>
  </tr>
  <tr >
    <td align="right"><div align="center">����Flash��ַ</div></td>
    <td ><input name="rightsrc" type="text" id="rightsrc" size="60">
      100*250</td>
  </tr>
  <tr >
    <td align="right"><div align="center">�������ײ��ر�Сͼ��</div></td>
    <td ><input name="closesrc" type="text" id="closesrc" value="/images/close.gif" size="40">
      ����&quot;0&quot;����ʾ</td>
  </tr>
  <tr >
    <td width="24%" align="right"><div align="center">�����ٶ�</div></td>
    <td width="76%" ><input name="speed" type="text" id="speed" value="0.8" size="8" onBlur="CheckNumber(this,'�����ٶ�');">
    ��Χ 0.1~1.0 ֵԽ��,�ٶ�Խ��</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
 
<%End Sub

Sub ArtPhoto()
%>
<html>
<head>
<title>����ͼƬ��������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetPhoto('+document.myform.PhotoWidth.value+','+document.myform.PhotoHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>��ʾͼƬ����</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">ͼƬ��ȣ�</div></td>
    <td width="60%" ><input name="PhotoWidth" type="text" onBlur="CheckNumber(this,'ͼƬ���');" size="6" value="130">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">ͼƬ�߶ȣ�</div></td>
    <td ><input name="PhotoHeight" type="text" size="6" onBlur="CheckNumber(this,'ͼƬ�߶�');" value="90">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
 <%
End Sub
Sub DownPhoto()
%>
<html>
<head>
<title>������������ͼ��������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetDownPhoto('+document.myform.PhotoWidth.value+','+document.myform.PhotoHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>��ʾ��������ͼ����</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">��������ͼ��ȣ�</div></td>
    <td width="60%" ><input name="PhotoWidth" type="text" onBlur="CheckNumber(this,'��������ͼ���');" size="6" value="130">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">��������ͼ�߶ȣ�</div></td>
    <td ><input name="PhotoHeight" type="text" size="6" onBlur="CheckNumber(this,'��������ͼ�߶�');" value="90">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub Flash()
%>
<html>
<head>
<title>����Flash��������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetFlash('+document.myform.FlashWidth.value+','+document.myform.FlashHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>��ʾFlash����</LEGEND>
<table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">Flash��ȣ�</div></td>
    <td width="60%" ><input name="FlashWidth" type="text" onBlur="CheckNumber(this,'Flash���������');" size="6" value="550">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">Flash�߶ȣ�</div></td>
    <td ><input name="FlashHeight" type="text" size="6" onBlur="CheckNumber(this,'Flash�������߶�');" value="380">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%End Sub
Sub FlashPlayer()
%>
<html>
<head>
<title>Flash��������������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetFlashByPlayer('+document.myform.FlashWidth.value+','+document.myform.FlashHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>��ʾFlash����������</LEGEND>
<table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">Flash��������ȣ�</div></td>
    <td width="60%" ><input name="FlashWidth" type="text" onBlur="CheckNumber(this,'Flash���������');" size="6" value="550">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">Flash�������߶ȣ�</div></td>
    <td ><input name="FlashHeight" type="text" size="6" onBlur="CheckNumber(this,'Flash�������߶�');" value="380">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub GetMusicList()
%>
<html>
<head>
<title>���������б��������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var TypeID,Val,ShowSelect,type,ShowMouseTX,ShowDetailTF;
	
	for (var i=0;i<document.myform.ShowSelect.length;i++){
	 var KM = document.myform.ShowSelect[i];
	if (KM.checked==true)	   
		ShowSelect = KM.value
	}
	for (var i=0;i<document.myform.type.length;i++){
	 var KM = document.myform.type[i];
	if (KM.checked==true)	   
		type = KM.value
	}
	for (var i=0;i<document.myform.ShowMouseTX.length;i++){
	 var KM = document.myform.ShowMouseTX[i];
	if (KM.checked==true)	   
		ShowMouseTX = KM.value
	}
	for (var i=0;i<document.myform.ShowDetailTF.length;i++){
	 var KM = document.myform.ShowDetailTF[i];
	if (KM.checked==true)	   
		ShowDetailTF = KM.value
	}

    Val = '{=GetMusicList('+document.myform.TypeID.value+','+ShowSelect+','+type+','+document.myform.Num.value+','+document.myform.RowHeight.value+','+ShowMouseTX+','+ShowDetailTF+','+document.myform.Row.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>

<link href="Editor.css" rel="stylesheet">
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
</head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>���������б��������</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td align="right"><div align="center">ѡ�����</div></td>
    <td >
	<select name="TypeID">
	 <option value='0'>-��ָ���κ����-</option>
	 <option value='-1' style="color:red">-��ǰ���ͨ��-</option>
	 <%
	  dim rs
	  set rs=server.createobject("adodb.recordset")
	  rs.open "select SclassID,Sclass from KS_MSSClass",conn,1,1
	  do while not rs.eof
	    response.write "<option value=""" & rs(0) & """>" & rs(1) & "</option>"
		rs.movenext
	  loop
	  rs.close
	  set rs=nothing
	  conn.close
	  set conn=nothing
	 %>
	</select>
	</td>
  </tr>
  <tr >
    <td align="right"><div align="center">��ʾѡ���</div></td>
    <td ><input name="ShowSelect" type="radio" value="true" checked>
      ��
        <input type="radio" name="ShowSelect" value="false">
        ��</td>
  </tr>
  <tr >
    <td align="right"><div align="center">�б�����</div></td>
    <td ><input name="type" type="radio" value="0" checked>
      ���¸���
        <input type="radio" name="type" value="1">
        �Ƽ�����
        <input type="radio" name="type" value="2">
        �ȵ����</td>
  </tr>
  <tr >
    <td width="24%" align="right"><div align="center">�г������׸���</div></td>
    <td width="76%" ><input name="Num" type="text" id="Num" value="10" size="8" onBlur="CheckNumber(this,'��������');">
      �� ÿ����ʾ: 
        <input name="Row" type="text" id="Row" value="2" size="6" onBlur="CheckNumber(this,'��������');">
        ��</td>
  </tr>
  <tr >
    <td align="right"><div align="center">����֮����о�</div></td>
    <td ><input name="RowHeight" type="text" id="RowHeight" value="25" size="8" onBlur="CheckNumber(this,'��������');">
      px</td>
  </tr>
  <tr >
    <td align="right"><div align="center">��꾭���Ƿ���Ч</div></td>
    <td ><input name="ShowMouseTX" type="radio" value="true" checked>
��
  <input type="radio" name="ShowMouseTX" value="false">
��</td>
  </tr>
  <tr >
    <td align="right"><div align="center">�г��Ƿ���ʾ��ϸ</div></td>
    <td ><input name="ShowDetailTF" type="radio" value="true" checked>
��
  <input type="radio" name="ShowDetailTF" value="false">
�� ����ʾ��������ϸ�������أ��ղص�</td>
  </tr>
</table>
</FIELDSET></td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
<tr>
  <td height="30"><div align="center"><span class="STYLE1">��ע���˱�ǩ����Ƶ��ͨ��</span></div></td>
</tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub GetSpecialList()
%>
<html>
<head>
<title>ר���б��������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val,type,ShowMouseTX,ShowDetailTF;
	
	for (var i=0;i<document.myform.type.length;i++){
	 var KM = document.myform.type[i];
	if (KM.checked==true)	   
		type = KM.value
	}

	for (var i=0;i<document.myform.ShowDetailTF.length;i++){
	 var KM = document.myform.ShowDetailTF[i];
	if (KM.checked==true)	   
		ShowDetailTF = KM.value
	}

    Val = '{=GetMusicSpecialList('+type+','+document.myform.Num.value+','+document.myform.ColNum.value+','+document.myform.PhotoWidth.value+','+document.myform.PhotoHeight.value+','+document.myform.SpecialNameLen.value+','+ShowDetailTF+')}';  
    window.returnValue = Val;
    window.close();
}
</script>

<link href="Editor.css" rel="stylesheet">
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
</head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>ר���б��������</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td align="right"><div align="center">�б�����</div></td>
    <td ><input name="type" type="radio" value="0" checked>
      ����ר��
        <input type="radio" name="type" value="1">
        �Ƽ�ר��
        <input type="radio" name="type" value="2">
        �ȵ�ר��</td>
  </tr>
  <tr >
    <td width="24%" align="right"><div align="center">�г�������ר��</div></td>
    <td width="76%" ><input name="Num" type="text" id="Num" value="10" size="8" onBlur="CheckNumber(this,'�г�������ר��');">
      ��</td>
  </tr>
  <tr >
    <td align="right"><div align="center">ר����������</div></td>
    <td ><input name="ColNum" type="text" id="ColNum" value="1" size="8" onBlur="CheckNumber(this,'ר����������');">
      px</td>
  </tr>
  <tr >
    <td align="right"><div align="center">ר��ͼƬ�Ŀ��</div></td>
    <td ><input name="PhotoWidth" type="text" id="PhotoWidth" value="90" size="8" onBlur="CheckNumber(this,'ר��ͼƬ�Ŀ��');">
px</td>
  </tr>
  <tr >
    <td align="right"><div align="center">ר��ͼƬ�ĸ߶�</div></td>
    <td ><input name="PhotoHeight" type="text" id="PhotoHeight" value="80" size="8" onBlur="CheckNumber(this,'ר��ͼƬ�ĸ߶�');">
px</td>
  </tr>
  <tr >
    <td align="right"><div align="center">ȡר����������</div></td>
    <td ><input name="SpecialNameLen" type="text" id="SpecialNameLen" value="8" size="8" onBlur="CheckNumber(this,'ȡר����������');"> 
      �� һ������=����Ӣ���ַ� </td>
  </tr>
  <tr >
    <td align="right"><div align="center">�Ƿ���ʾ���й�˾����������</div></td>
    <td ><input name="ShowDetailTF" type="radio" value="true" checked>
��
  <input type="radio" name="ShowDetailTF" value="false">
�� ����ʾ��������ϸ�������أ��ղص�</td>
  </tr>
</table>
</FIELDSET></td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
<tr>
  <td height="30"><div align="center"><span class="STYLE1">��ע��Щ��ǩ����Ƶ��ͨ��</span></div></td>
</tr>
</table>
</form>
</body>
</html>
 <%End Sub
Sub Logo
%>
<html>
<head>
<title>������վLogo��������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetLogo('+document.myform.FlashWidth.value+','+document.myform.FlashHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>��ʾLogo����</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">Logo��ȣ�</div></td>
    <td width="60%" ><input name="FlashWidth" type="text" onBlur="CheckNumber(this,'Flash���������');" size="6" value="130">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">Logo�߶ȣ�</div></td>
    <td ><input name="FlashHeight" type="text" size="6" onBlur="CheckNumber(this,'Flash�������߶�');" value="90">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub Tags()
%>
<html>
<head>
<title>����Tags��������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetTags('+document.myform.sorts.value+','+document.myform.num.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>��ʾTags����</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">��ʾTags����</div></td>
    <td width="60%" ><input name="num" type="text" id="num" value="50" size="6">
    ��</td>
  </tr>
  <tr>
    <td align="right"><div align="center">Tags����ʽ��</div></td>
    <td ><select name="sorts">
      <option value="1">���������(����Tags)</option>
      <option value="2">������ʱ�併��</option>
      <option value="3">���ʱ��</option>
    </select>
    </td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%End Sub
Sub MovieDown()
%>
<html>
<head>
<title>����ӰƬ�����б�����</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetMovieDownList('+document.myform.Num.value+',"'+document.myform.Navi.value+'")}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>��ʾӰƬ�����б�����</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">ÿ����ʾ������</div></td>
    <td width="60%" ><input name="Num" type="text" onBlur="CheckNumber(this,'ÿ����ʾ����');" size="15" value="5">
     </td>
  </tr>
  <tr>
    <td align="right"><div align="center">����ͼ�꣺</div></td>
    <td ><input name="Navi" type="text" size="15" onBlur="CheckNumber(this,'����ͼ��');" value="/images/movienavi.gif">
    </td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub MoviePage()
%>
<html>
<head>
<title>��������ҳflv����������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetMoviePagePlay('+document.myform.PhotoWidth.value+','+document.myform.PhotoHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>��ʾ����ҳflv����������</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">ӰƬ��ȣ�</div></td>
    <td width="60%" ><input name="PhotoWidth" type="text" onBlur="CheckNumber(this,'ӰƬ���');" size="6" value="450">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">ӰƬ�߶ȣ�</div></td>
    <td ><input name="PhotoHeight" type="text" size="6" onBlur="CheckNumber(this,'ӰƬ�߶�');" value="450">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub MoviePhoto()
%>
<html>
<head>
<title>����ӰƬͼƬ��������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetMoviePhoto('+document.myform.PhotoWidth.value+','+document.myform.PhotoHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>��ʾӰƬͼƬ����</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">ӰƬͼƬ��ȣ�</div></td>
    <td width="60%" ><input name="PhotoWidth" type="text" onBlur="CheckNumber(this,'ӰƬͼƬ���');" size="6" value="250">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">ӰƬͼƬ�߶ȣ�</div></td>
    <td ><input name="PhotoHeight" type="text" size="6" onBlur="CheckNumber(this,'ӰƬͼƬ�߶�');" value="250">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub MoviePlay()
%>
<html>
<head>
<title>����ӰƬ�����б�����</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetMoviePlayList('+document.myform.Num.value+',"'+document.myform.Navi.value+'")}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>��ʾӰƬ�����б�����</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">ÿ����ʾ������</div></td>
    <td width="60%" ><input name="Num" type="text" onBlur="CheckNumber(this,'ÿ����ʾ����');" size="15" value="5">
     </td>
  </tr>
  <tr>
    <td align="right"><div align="center">����ͼ�꣺</div></td>
    <td ><input name="Navi" type="text" size="15" onBlur="CheckNumber(this,'����ͼ��');" value="/images/movienavi.gif">
    </td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub ProductGroupPhoto()
%>
<html>
<head>
<title>������ƷͼƬ���������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetGroupPhoto('+document.myform.PhotoWidth.value+','+document.myform.PhotoHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>��ʾ��ƷͼƬ������</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">Ԥ��ͼ��ȣ�</div></td>
    <td width="60%" ><input name="PhotoWidth" type="text" onBlur="CheckNumber(this,'��ƷԤ��ͼ���');" size="6" value="200">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">��ƷԤ��ͼ�߶ȣ�</div></td>
    <td ><input name="PhotoHeight" type="text" size="6" onBlur="CheckNumber(this,'��ƷԤ��ͼ�߶�');" value="200">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub ProductPhoto()
%>
<html>
<head>
<title>������Ʒ����ͼ��������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetProductPhoto('+document.myform.PhotoWidth.value+','+document.myform.PhotoHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>��ʾ��Ʒ����ͼ����</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">��Ʒ����ͼ��ȣ�</div></td>
    <td width="60%" ><input name="PhotoWidth" type="text" onBlur="CheckNumber(this,'��Ʒ����ͼ���');" size="6" value="130">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">��Ʒ����ͼ�߶ȣ�</div></td>
    <td ><input name="PhotoHeight" type="text" size="6" onBlur="CheckNumber(this,'��Ʒ����ͼ�߶�');" value="90">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub Status1()
%>
<html>
<head>
<title>����״̬������Ч����������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
	if (document.myform.text.value=='')
	  {
	   alert('����������!')
	   document.myform.text.focus();
	   return false;
	  }
    Val = '{=JS_Status1("'+document.myform.text.value+'",'+document.myform.speed.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<script language="JavaScript">
var msg = "��ӭ��ʹ�ÿ�Ѵ��վ����ϵͳ! " ;
var interval = 120
var spacelen = 120;
var space10=" ";
var seq=0;
function KS_Status1() {
len = msg.length;
window.status = msg.substring(0, seq+1);
seq++;
if ( seq >= len ) {
seq = 0;
window.status = '';
window.setTimeout("KS_Status1();", interval );
}
else
window.setTimeout("KS_Status1();", interval );
}
KS_Status1();
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>״̬������Ч������</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td align="right"><div align="center">����ʾ������</div></td>
    <td ><input name="text" type="text" id="text" size="60"></td>
  </tr>
  <tr >
    <td align="right">&nbsp;</td>
    <td >��: ��ӭ���ٱ�վ!!!</td>
  </tr>
  <tr >
    <td width="24%" align="right"><div align="center">�����ٶ�</div></td>
    <td width="76%" ><input name="speed" type="text" id="speed" value="120" size="8" onBlur="CheckNumber(this,'�����ٶ�');">ֵԽ��,�ٶ�Խ��</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub Status2()
%>
<html>
<head>
<title>����״̬��������״̬���ϴ�������ѭ����ʾ��������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
	if (document.myform.text.value=='')
	  {
	   alert('����������!')
	   document.myform.text.focus();
	   return false;
	  }
    Val = '{=JS_Status2("'+document.myform.text.value+'",'+document.myform.speed.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>

<script>
<!--
function KS_Status2(seed)
{ var m1 = "��ӭ��ʹ�ÿ�Ѵ��վ����ϵͳ!" ;
var m2 = "" ;
var msg=m1+m2;
var out = " ";
var c = 1;
var speed = 120;
if (seed > 100)
{ seed-=2;
var cmd="KS_Status2(" + seed + ")";
timerTwo=window.setTimeout(cmd,speed);}
else if (seed <= 100 && seed > 0)
{ for (c=0 ; c < seed ; c++)
{ out+=" ";}
out+=msg; seed-=2;
var cmd="KS_Status2(" + seed + ")";
window.status=out;
timerTwo=window.setTimeout(cmd,speed); }
else if (seed <= 0)
{ if (-seed < msg.length)
{
out+=msg.substring(-seed,msg.length);
seed-=2;
var cmd="KS_Status2(" + seed + ")";
window.status=out;
timerTwo=window.setTimeout(cmd,speed);}
else { window.status=" ";
timerTwo=window.setTimeout("KS_Status2(100)",speed);
}
}
}
KS_Status2(100);
-->
</script>
      
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>������״̬���ϴ�������ѭ����ʾЧ������</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td align="right"><div align="center">����ʾ������</div></td>
    <td ><input name="text" type="text" id="text" size="60"></td>
  </tr>
  <tr >
    <td align="right">&nbsp;</td>
    <td >��: ��ӭ���ٱ�վ!!!</td>
  </tr>
  <tr >
    <td width="24%" align="right"><div align="center">�����ٶ�</div></td>
    <td width="76%" ><input name="speed" type="text" id="speed" value="120" size="8" onBlur="CheckNumber(this,'�����ٶ�');">
    ֵԽ��,�ٶ�Խ��</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub Status3()
%>
<html>
<head>
<title>����״̬��������״̬���ϴ�������ѭ����ʾ��������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
	if (document.myform.text.value=='')
	  {
	   alert('����������!')
	   document.myform.text.focus();
	   return false;
	  }
    Val = '{=JS_Status3("'+document.myform.text.value+'",'+document.myform.speed.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>

    <SCRIPT LANGUAGE="JavaScript">
<!--
var Message="��ӭ��ʹ�ÿ�Ѵ��վ����ϵͳ! ";
var place=1;
function scrollIn() {
window.status=Message.substring(0, place);
if (place >= Message.length) {
place=1;
window.setTimeout("KS_Status3()",300);
} else {
place++;
window.setTimeout("scrollIn()",200);
}
}
function KS_Status3() {
window.status=Message.substring(place, Message.length);
if (place >= Message.length) {
place=1;
window.setTimeout("scrollIn()", 100);
} else {
place++;
window.setTimeout("KS_Status3()", 200);
}
}
KS_Status3();
-->
</SCRIPT>  
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>������״̬���ϴ�������ѭ����ʾЧ������</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td align="right"><div align="center">����ʾ������</div></td>
    <td ><input name="text" type="text" id="text" size="60"></td>
  </tr>
  <tr >
    <td align="right">&nbsp;</td>
    <td >��: ��ӭ���ٱ�վ!!!</td>
  </tr>
  <tr >
    <td width="24%" align="right"><div align="center">�����ٶ�</div></td>
    <td width="76%" ><input name="speed" type="text" id="speed" value="150" size="8" onBlur="CheckNumber(this,'�����ٶ�');">
    ֵԽ��,�ٶ�Խ��</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%End Sub

Sub SupplyPhoto()
%>
<html>
<head>
<title>���빩����Ϣ����ͼ��������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetSupplyPhoto('+document.myform.PhotoWidth.value+','+document.myform.PhotoHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>��ʾ������Ϣ����ͼ����</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">������Ϣ����ͼ��ȣ�</div></td>
    <td width="60%" ><input name="PhotoWidth" type="text" onBlur="CheckNumber(this,'������Ϣ����ͼ���');" size="6" value="130">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">������Ϣ����ͼ�߶ȣ�</div></td>
    <td ><input name="PhotoHeight" type="text" size="6" onBlur="CheckNumber(this,'������Ϣ����ͼ�߶�');" value="90">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub TopUser()
%>
<html>
<head>
<title>�����û���¼���в�������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetTopUser('+document.myform.num.value+','+document.myform.more.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>��ʾ�û���¼��������</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">��ʾ�û�����</div></td>
    <td width="60%" ><input name="num" type="text" id="num" value="5" size="6">
      λ</td>
  </tr>
  <tr>
    <td align="right"><div align="center">�������ӣ�</div></td>
    <td ><input name="more" type="text" id="more" value="more..." size="20"> ���ղ����</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub UserDynamic()
%>
<html>
<head>
<title>�����û���̬��������</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetUserDynamic('+document.myform.num.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>��ʾ�û���̬��ǩ����</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">��ʾ���£�</div></td>
    <td width="60%" ><input name="num" type="text" id="num" value="10" size="6">
      ��</td>
  </tr>
  
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
%>