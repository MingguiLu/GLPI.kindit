<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************

Dim KSCls
Set KSCls = New Spacemore
KSCls.Kesion()
Set KSCls = Nothing

Class Spacemore
        Private KS, KSRFObj,str,totalput,maxperpage,currpage
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRFObj = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		      Dim FileContent
			  FileContent = KSRFObj.LoadTemplate(KS.SSetting(8))
			  FCls.RefreshType = "MoreGroup" '����ˢ�����ͣ��Ա�ȡ�õ�ǰλ�õ�����
			  Application(KS.SiteSN & "RefreshFolderID") = "0" '���õ�ǰˢ��Ŀ¼ID Ϊ"0" ��ȡ��ͨ�ñ�ǩ
			  If Trim(FileContent) = "" Then FileContent = "�ռ丱ģ�岻����!"
			  grouplist
			 FileContent=Replace(FileContent,"{$ShowMain}",str)
			FileContent=KSRFObj.KSLabelReplaceAll(FileContent)
		   Response.Write FileContent  
		End Sub
		
		'Ȧ���б�
	Sub GroupList()
		 MaxPerPage =KS.ChkClng(KS.SSetting(11))
		 dim classid:classid=KS.ChkClng(KS.S("ClassID"))
		 dim recommend:recommend=KS.ChkClng(KS.S("recommend"))
		   CurrPage = KS.ChkClng(KS.G("page"))
		  If CurrPage<=0 Then CurrPage=1

		 dim rsc:set rsc=conn.execute("select classname,classid from ks_teamClass order by orderid")
		 if not rsc.eof then
		   str="<div class=""categorybox"">" & vbcrlf
		   str=str &"<ul><li>����鿴��</li>"
		   If classid=0 then 
		     str=str &"<li class=""curr""><a href='moregroup.asp'>���з���</a></li>"
		   else
		     str=str &"<li><a href='moregroup.asp'>���з���</a></li>"
		   end if
			do while not rsc.eof
			 if classid=rsc(1) then
			   str=str & "<li class=""curr""><a href='?classid=" & rsc(1) &"'>" & rsc(0) & "</a></li>"
			 else
			   str=str & "<li><a href='?classid=" & rsc(1) &"'>" & rsc(0) & "</a></li>"
			 end if
			 rsc.movenext
			loop
		 end if
		 rsc.close:set rsc=nothing
	   str=str &"</ul>" & vbcrlf
	   str=str &"</div>" &vbcrlf	
	   
		  dim param:param=" where verific=1"
          if classid<>0 then param=param & " and a.classid=" & classid
			 if recommend<>0 then param=param & " and  recommend=1"
		 if ks.s("key")<>"" then param=param & " and teamname like '%" & ks.r(ks.s("key")) &"%'"
		  str=str & "  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "select a.*,b.classname from KS_team a inner join ks_teamclass b on a.classid=b.classid " & Param & " order by id desc",Conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				 str=str & "<tr><td style=""border: #efefef 1px dotted;text-align:center"">û���û�����Ȧ�ӣ�</td></tr>"
				 Else
							totalPut = RSObj.RecordCount
							If currpage >1 and (currpage - 1) * MaxPerPage < totalPut Then
								RSObj.Move (currpage - 1) * MaxPerPage
							End If
							call ShowGroup(RSObj)

				 End If
		 
		 str=str &"            </table>" & vbcrlf
		 RSObj.Close:Set RSObj=Nothing
		 str=str & KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
		 str=str & "<div class=""clear""></div>"
		 str=str &"<table border=""0"" cellpadding=""1"" cellspacing=""1"" align=""center"" width=""98%"">" &vbcrlf
		  str=str & "<form name=""myform"" action=""moregroup.asp"" method=""get""/> <tr height=""22"">"
	   str=str & "<td style=""text-align:left"" colspan=2><strong>��Ȧ������������</strong><input style=""border:1px #000 solid;height:18px;"" type=""text"" size=""12"" name=""key"">&nbsp;&nbsp;<input type=""submit"" value= "" �� �� "" class=""btn""></td>"
	   str=str & "</form></tr>"
	   str=str & "</table><br/><br/>"
	 End Sub
			 
	 Sub ShowGroup(RS)		 
		 Dim I
		 Do While Not RS.Eof 
		     str=str &"<tr>"
		     str=str &"<td style=""border-bottom:#9999CC dotted 1px;"">"& vbcrlf
			 str=str & " <a href=""group.asp?id=" & rs("id") & """ title=""" & rs("teamname") & """ target=""_blank""><img align=""left"" src=""" & rs("photourl") & """ width=""110"" height=""80"" border=""0"" style=""border:1px solid #f1f1f1;padding:2px;margin-right:6px;""></a>"
			str=str & "<a class=""teamname"" href=""group.asp?id=" & rs("id") & """ title=""" & rs("teamname") & """ target=""_blank""> " & rs("TeamName") & "</a><br>�����ߣ�" & rs("username") & "<br>����ʱ��:" &rs("adddate") & "<br>Ȧ�ӷ��ࣺ" & rs("classname") & "<br>����/�ظ���" & conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id") & "  and parentid=0")(0) & "/" & conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id"))(0) & "&nbsp;&nbsp;&nbsp;��Ա:" & conn.execute("select count(username)  from ks_teamusers where status=3 and teamid=" & rs("id"))(0) & "��  </td>"
			str=str & "</tr>"
			rs.movenext
			I = I + 1
		  If I >= MaxPerPage Then Exit Do
		 Loop
	 End Sub
		
		
	
End Class
%>
