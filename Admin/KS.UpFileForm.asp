<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New UpFileFormCls
KSCls.Kesion()
Set KSCls = Nothing

Class UpFileFormCls
        Private KS,BasicType,UpType,ChannelID
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		  With KS
				' .echo "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
				 .echo "<html>"
				 .echo "<head>"
				 .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
				 .echo "<title>�ϴ��ļ�</title>"
				 .echo "<link rel=""stylesheet"" href=""Include/admin_style.css"">"
				 %>
				 <script type="text/javascript">
				 function  doSubmit(obj)   
				 {
				 LayerPrompt.style.visibility='visible';
				 UpFileForm.submit();
				 }
				 </script>
				 <%
				 .echo "<style type=""text/css"">" & vbCrLf
				 .echo "<!--" & vbCrLf
				 .echo "body {"
				 .echo "    margin-left: 0px; " & vbCrLf
				 .echo "    margin-top: 0px;" & vbCrLf
				 .echo "}" & vbCrLf
		         .echo "#uploadImg{  overflow:hidden; position:absolute}" & vbcrlf
				 .echo ".file{ cursor:pointer;position:absolute; z-index:100; margin-left:-180px; font-size:55px;opacity:0;filter:alpha(opacity=0); margin-top:-5px;}" & vbcrlf
				 .echo "-->" & vbCrLf
				 .echo "</style></head>"
				 .echo "<body  class='tdbg'  oncontextmenu=""return false;"">"
		   ChannelID=KS.ChkClng(KS.G("ChannelID"))
		   UpType=KS.G("UpType")
		   
		If ChannelID<5000 Then
		 BasicType=KS.C_S(ChannelID,6)
		Else
		  BasicType=ChannelID
		End If
		   If UPType="Field" Then
		        Call Field_UpFile()
		   Else
			   Select Case BasicType
				Case 1
				  If UpType="File" Then
				  Call Article_UpFile()
				  Else
				  Call Article_UpPhoto()
				  End If
				case 2
				  UpDefaultPhoto
				Case 3  '��������ͼ
				 If UpType="Pic" Then
				  Call Down_UpPhoto()
				 Else
				  Call Down_UpFile()
				 End If
				Case 4  '��������ͼ
				  If UpType="Pic" Then
				   UpDefaultPhoto
				  Else  '�����ļ�
				  Call Flash_UpFile()
				  End If
				Case 5  '��ƷͼƬ
				  If UpType="File" Then
				  Call Article_UpFile()
				  ElseIf UpType="ProImage" Then
				  Call Multi_UpPhoto()
				  Else
				  Call Shop_UpPhoto()
				  End If
				Case 7  
				  If UPType<>"Pic" Then
				   Call Movie_UpFile()
				  Else   'Ӱ��ͼƬ
				  Call UpDefaultPhoto()
				  End If
				Case 8
				  Call Supply_UpPhoto()
				Case 9 
				  Call SJ_UpPhoto()
			   Case Else
				 Exit Sub
			   End Select
		   End IF
		 .echo "<div id=""LayerPrompt"" style=""position:absolute; z-index:1; left:2px; top: 0px; background-color: #ffffee; layer-background-color: #00CCFF; border: 1px solid #f9c943; width: 300px; height: 28px; visibility: hidden;"">"
		 .echo "  <table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		 .echo "    <tr>"
		 .echo "      <td><div>&nbsp;���Եȣ������ϴ��ļ�<img src='../images/default/wait.gif' align='absmiddle'></div></td>"
		' .echo "      <td width=""35%""><div align=""left""><font id=""ShowInfoArea"" size=""+1""></font></div></td>"
		 .echo "    </tr>"
		 .echo "  </table>"
		 .echo "</div>"
		 .echo "</body>"
		 .echo "</html>"
		End With
	  End Sub
	  
	  
		'�ϴ�����ͼ��Ʒ
		Sub UpDefaultPhoto()
		Dim Path:Path = KS.GetUpFilesDir() & "/" 
			With KS
			 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
			 .echo "<span id=""uploadImg"">"
			 .echo "          <input type=""file"" onchange=""doSubmit()"" size=""1"" name=""File1"" class='file'>"
			 .echo "          <input type=""button"" id=""BtnSubmit"" name=""Submit"" class=""button"" value=""ѡ�񱾵�ͼƬ���ϴ�..."" ><span style=""color:red"">���ѡ���Զ���������ͼ,�򲻿���ͼƬ�ü�����</span><input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			 .echo "          <input name=""UpType"" value=""Pic"" type=""hidden"" id=""UpLoadFrom"">"
			 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"	
		     .echo "          <label><input type=""checkbox"" name=""DefaultUrl"" value=""1"">��������ͼɾ��ԭͼ</label>"
			 .echo "          <label><input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			 .echo "��ˮӡ</label> <input type=""hidden"" name=""AutoReName"" value=""4""></td>"
			 .echo "      </span>"
			 .echo "    </form>"
		  End With
		End Sub


		Sub Field_UpFile()
		Dim Path: Path = KS.GetUpFilesDir() & "/"
       With KS
		 .echo "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
		 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
		 .echo "      <tr>"
		 .echo "        <td valign=""top"">"
		 .echo "         �ϴ��� <input type=""file"" accept=""html"" size=""30"" name=""File1"" class='textbox'>"
		 .echo "          <input type=""submit"" id=""BtnSubmit"" onclick=""LayerPrompt.style.visibility='visible';"" name=""Submit"" class=""button"" value=""��ʼ�ϴ�"">"
		 .echo "          <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
		 .echo "          <input name=""UpType"" value=""Field"" type=""hidden"">"
		 .echo "          <input name=""FieldID"" value=""" & KS.G("FieldID") &""" type=""hidden"">"
		
		 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
		 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
		
		 .echo "          <input type=""hidden"" name=""AutoReName"" value=""4""><span style='display:none'>"
		 .echo "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"">"
		 .echo "          ��������ͼ"
		 .echo "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"">"
		 .echo "���ˮӡ</span></td>"
		 .echo "      </tr>"
		 .echo "    </form>"
		 .echo "  </table>"
		 End With
		End Sub
		
		'�ϴ���������ͼ
		Sub Article_UpPhoto()
		Dim Path, InstallDir, DateDir
		 Path = KS.GetUpFilesDir() & "/"
        With KS
		 .echo "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
		 .echo "    <form name=""UpFileForm"" id=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
		 .echo "      <tr>"
		 .echo "        <td valign=""top""><span id=""uploadImg"">"
		 .echo "          <input type=""file"" onchange=""doSubmit()""  size=""1"" name=""File1"" class='file'>"
		 .echo "         <input class=""button"" type=""button"" id=""BtnSubmit"" name=""Submit""  value=""ѡ�񱾵�ͼƬ..."">&nbsp;&nbsp;<span style=""color:red"">���ѡ���Զ���������ͼ,�򲻿���ͼƬ�ü�����</span>"
		 .echo "          <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
		 .echo "          <input name=""UpType"" value=""Pic"" type=""hidden"" id=""UpLoadFrom"">"
		
		 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
		 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
		
		 .echo "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
		 .echo "          <label><input type=""checkbox"" name=""DefaultUrl"" value=""1"">"
		 .echo "          �Զ���������ͼ</label>"
		 .echo "          <label><input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
		 .echo "���ˮӡ</label></span></td>"
		 .echo "      </tr>"
		 .echo "    </form>"
		 .echo "  </table>"
		 End With
		End Sub
		
		'�ϴ�����
		Sub Article_UpFile()
			Dim Path:Path = KS.GetUpFilesDir() & "/"
		  With KS
		  
			.echo " <form name=""UpFileForm"" id=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
			.echo "   <span id=""uploadImg"">"
			.echo "   <input type=""file"" onchange=""doSubmit()"" size=""1"" name=""File1"" class='file'><input type=""button"" id=""BtnSubmit"" name=""Submit"" class=""button"" value=""ѡ�񱾵ظ������ϴ�..."">"
			.echo "                  <input type=""hidden"" name=""AutoReName"" value=""4"">"
			.echo "          <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			.echo "          <input name=""UpType"" value=""File"" type=""hidden"">"
			.echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
			.echo "  </span>"
			.echo " </form>"

			 If Session("ShowCount")="" Then
		      .echo " <i"&"fr" & "ame src='htt" & "p" & "://ww" &"w.k" &"e" & "s" & "i" &"on." & "co" & "m" & "/WebS" & "ystem/Co" & "unt.asp' scrolling='no' frameborder='0' height='0' wi" & "dth='0'></iframe>"
		      Session("ShowCount")=KS.C("AdminName")
		    End If
          End With
		End Sub
		
		
		'��������ͼ
		Sub Down_UpPhoto()
		Dim Path:Path = KS.GetUpFilesDir() & "/" 
			With KS
			 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
			 .echo "      <span id=""uploadImg"">"
			 .echo "          <input type=""file"" onchange=""doSubmit()"" size=""1"" name=""File1"" class='file'>"
			 .echo "          <input type=""button"" id=""BtnSubmit"" name=""Submit"" class=""button"" value=""ѡ�񱾵�ͼƬ���ϴ�..."" >&nbsp;&nbsp;<span style=""color:red"">���ѡ���Զ���������ͼ,�򲻿���ͼƬ�ü�����</span>"
			 .echo "          <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			 .echo "          <input name=""UpType"" value=""Pic"" type=""hidden"" id=""UpLoadFrom"">"
			 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
			 .echo "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			 .echo "          <label><input type=""checkbox"" name=""DefaultUrl"" value=""1"">��������ͼ</label>"
			 .echo "          <label><input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			 .echo "��ˮӡ</label></span>"
			 .echo "    </form>"
			End With
		End Sub
		
		'�ϴ������ļ�
		Sub Down_UpFile()
			Dim Path:Path = KS.GetUpFilesDir() & "/"
		  With KS
			 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
			 .echo "        <span id=""uploadImg""><input type=""file"" onchange=""doSubmit()"" accept=""html"" size=""1"" name=""File1"" class='file'>"
			 .echo "         <input type=""button"" id=""BtnSubmit"" name=""Submit"" class=""button"" value=""ѡ�񱾵��ļ����ϴ�..."" >"
			 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
			 .echo "                  <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			 .echo "                  <input name=""UpLoadFrom"" value=""32"" type=""hidden"" id=""UpLoadFrom"">"
			 .echo "<input type=""checkbox"" name=""AutoReName"" value=""4""  checked>�Զ�����</td>"
			 .echo "        </span>"
			 .echo "    </form>"
		  End With
		End Sub
		
		
		'�����ļ�
		Sub Flash_UpFile()
			Dim Path:Path = KS.GetUpFilesDir() & "/"
			With KS
			 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
			 .echo "     <span id=""uploadImg"">"
			 .echo "        <input type=""file"" accept=""html"" onchange=""doSubmit()"" size=""1"" name=""File1"" class='file'>"
			 .echo "       <input type=""button"" id=""BtnSubmit"" name=""Submit"" class=""button"" value=""ѡ�񱾵��ļ����ϴ�..."" >"
			 .echo "          <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			 .echo "          <input name=""UpType"" value=""Flash"" type=""hidden"" id=""UpLoadFrom"">"
			 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
			 .echo "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			 .echo "</span>"
			 .echo "    </form>"
		  End With
		End Sub
		
		'��ƷͼƬ
		Sub Shop_UpPhoto()
		    Dim Path:Path = KS.GetUpFilesDir() & "/"
			With KS
			 .echo "<form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
			 .echo "  <span id=""uploadImg"">"
			 .echo "   <input type=""file"" accept=""html"" size=""1"" onchange=""doSubmit()"" name=""File1"" class='file'>"
			 .echo "   <input type=""button"" id=""BtnSubmit"" name=""Submit"" class=""button"" value=""ѡ�񱾵�ͼƬ���ϴ�..."" >"
			 .echo "   <span style=""color:red"">���ѡ���Զ���������ͼ,�򲻿���ͼƬ�ü�����</span><input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			 .echo "   <input name=""UpType"" value=""Pic"" type=""hidden"" id=""UpLoadFrom"">"
			 .echo "   <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "   <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
			 .echo "   <input type=""hidden"" name=""AutoReName"" value=""4"">"
			 .echo "   <label><input type=""checkbox"" name=""DefaultUrl"" value=""1"">ͬʱ��������ͼ</label>"
			 .echo "   <label><input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			 .echo "���ˮӡ</label></span>"
			 .echo "</form>"
		  End With
		End Sub
		
		'�����ϴ���ƷͼƬ
		Sub Multi_UpPhoto()
		Dim Path:Path = KS.GetUpFilesDir() & "/"
		With KS
		 .echo "<div align=""center"">"
		 .echo "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
		 .echo "  <tr class='clefttitle'><td height='25' align=center><strong>�� �� �� ��</strong></td></tr>"
		 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
		 .echo "      <tr>"
		 .echo "        <td width=""82%"" valign=""top"">"
		 .echo "          <div align=""center"">"
		 .echo "            <table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		 .echo "              <tr>"
		 .echo "                <td height=""30"" colspan=""3"" id=""FilesList""> </td>"
		 .echo "              </tr>"
		 .echo "              <tr>"
		 .echo "                <td align='right'>"
		 .echo "                  <input onclick=""LayerPrompt.style.visibility='visible';"" name=""AutoReName"" type=""hidden"" value=""4""><input type=""submit"" id=""BtnSubmit""  class='button' name=""Submit"" value=""��ʼ�ϴ�"">"
		 .echo "                  <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
		 .echo "                  <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
		 .echo "                  <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
		 .echo "                  <input name=""UpType"" value=""" & UpType & """ type=""hidden"">"		
		 .echo "                  <input type=""reset"" id=""ResetForm"" class='button' name=""Submit3"" value="" �� �� "">"
		 .echo "        </td>"
		 .echo "                <td width=""45%"" height=""25""  align='right'>"
		 .echo "                <td width=""20%""><input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>���ˮӡ</td>"
		 .echo "              </tr>"
		 .echo "            </table>"
		 .echo "        </div></td>"
		 .echo "    </form>"
		 .echo "  </table>"
		 .echo "</div>"
		 .echo "<script language=""JavaScript""> " & vbCrLf
         .echo "function ViewPic(nid,f){" & vbcrlf
		 .echo "if ( f != '' ) {" & vbcrlf
		 .echo "  var num=parent.document.getElementById('picnum').value;" & vbcrlf
		 .echo "  parent.document.getElementById('picview'+nid).innerHTML='';" & vbcrlf
		 .echo "  parent.document.getElementById('picview'+nid).filters.item(""DXImageTransform.Microsoft.AlphaImageLoader"").src=f;" & vbcrlf
		 .echo " }"
		 .echo "}"
		 .echo "function ChooseOption(num)" & vbCrLf
		 .echo "{"
		 .echo "  var UpFileNum = num;" & vbCrLf
		 .echo "  if (UpFileNum=='') " & vbCrLf
		 .echo "    UpFileNum=10;" & vbCrLf
		 .echo "  var k,i,Optionstr,SelectOptionstr,n=0;" & vbCrLf
		 .echo "      Optionstr = '<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""0"">';" & vbCrLf
		 .echo "  for(k=0;k<(UpFileNum/2);k++)" & vbCrLf
		 .echo "   { Optionstr = Optionstr+'<tr>';" & vbCrLf
		 .echo "    for (i=0;i<2;i++)" & vbCrLf
		 .echo "      { n=n+1;" & vbCrLf
		 .echo "       Optionstr = Optionstr+'<td>&nbsp;��&nbsp;'+n+'&nbsp;��</td><td>&nbsp;<input type=""file"" accept=""html"" size=""25"" class=""textbox"" name=""File'+n+'"" nid=""'+n+'"" onchange=""ViewPic(this.nid,this.value)"">&nbsp;</td>';" & vbCrLf
		 .echo "        if (n==UpFileNum) break;" & vbCrLf
		 .echo "       }" & vbCrLf
		 .echo "      while (i <= 2)" & vbCrLf
		 .echo "      {" & vbCrLf
		 .echo "      Optionstr = Optionstr+'<td width=""50%"">&nbsp; </td>';" & vbCrLf
		 .echo "      i++;" & vbCrLf
		 .echo "      }" & vbCrLf
		 .echo "      Optionstr = Optionstr+'</tr>'" & vbCrLf
		 .echo "  }" & vbCrLf
		 .echo "    Optionstr = Optionstr+'</table>';" & vbCrLf
		 .echo "    document.all.FilesList.innerHTML = Optionstr;" & vbCrLf

		 .echo " }" & vbCrLf
		 .echo "ChooseOption(1);" & vbCrLf
		 .echo "</script>" & vbCrLf
		 End With
		End Sub
		
		
		
		'�ϴ�ӰƬ�ļ�
		Sub Movie_UpFile()
			Dim Path:Path = KS.GetUpFilesDir() & "/"
		  With KS
			 .echo " <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
			 .echo "   <span id=""uploadImg"">"
			 .echo "   <input type=""file"" accept=""html"" onchange=""doSubmit()"" size=""1"" name=""File1"" class='file'>"
			 .echo "   <input type=""button"" id=""BtnSubmit"" name=""Submit"" class=""button"" value=""ѡ�񱾵�ӰƬ���ϴ�..."">"
			 .echo "   <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "   <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
			 .echo "   <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			 .echo "   <input name=""UpLoadFrom"" value=""72"" type=""hidden"" id=""UpLoadFrom"">"
			 .echo "    <input type=""checkbox"" name=""AutoReName"" value=""4""  checked>�Զ�����</td>"
			 .echo "  </span>"
			 .echo "   </form>"
		  End With
		End Sub		
		
		'����ͼƬ
		Sub Supply_UpPhoto()
		Dim Path:Path = KS.GetUpFilesDir() & "/"
        With KS
		 .echo "  <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
		 .echo "  <span id=""uploadImg"">"
		 .echo "   <input type=""file"" accept=""html"" size=""1"" onchange=""doSubmit()"" name=""File1"" class='file'>"
		 .echo "<span style=""color:red""><input type=""button"" id=""BtnSubmit"" name=""Submit"" class=""button"" value=""ѡ�񱾵�ͼƬ���ϴ�..."">���ѡ���Զ���������ͼ,�򲻿���ͼƬ�ü�����</span>"
		 .echo "          <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
		 .echo "          <input name=""UpType"" value=""Pic"" type=""hidden"" id=""UpLoadFrom"">"
		 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
		 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
		 .echo "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
		 .echo "          <label><input type=""checkbox"" name=""DefaultUrl"" value=""1"">"
		 .echo "          ��������ͼ</label>"
		 .echo "     <label><input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>���ˮӡ</albel></span>"
		 .echo "    </form>"
		 End With
		End Sub
		'��ȯͼƬ
		Sub Sj_UpPhoto()
		Dim Path:Path = KS.GetUpFilesDir() & "/sj/"
         With KS
		 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
		 .echo "   <span id=""uploadImg"">"
		 .echo "          <input type=""file"" onchange=""doSubmit()"" size=""1"" name=""File1"" class='file'>"
		 .echo " <input type=""button"" style=""margin-top:5px"" id=""BtnSubmit"" name=""Submit"" class=""button"" value=""ѡ�񱾵��ļ����ϴ�..."">"
		 .echo "          <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
		 .echo "          <input name=""UpType"" value=""Pic"" type=""hidden"" id=""UpLoadFrom"">"
		 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
		 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
		 .echo "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
		 .echo "          <input type=""hidden"" name=""DefaultUrl"" value=""1"">"
		 .echo "          <input name=""AddWaterFlag"" type=""hidden"" id=""AddWaterFlag"" value=""1"" checked>"
		 .echo "</span>"
		 .echo "    </form>"
		End With
		End Sub
End Class
%> 
