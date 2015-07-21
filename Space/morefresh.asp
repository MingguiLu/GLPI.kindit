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
        Private KS, KSR,MaxPerPage,CurrPage,TotalPut
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		      Dim FileContent
				   FileContent = KSR.LoadTemplate(KS.SSetting(31))
				   FCls.RefreshType = "Morefresh" '设置刷新类型，以便取得当前位置导航等
				   Application(KS.SiteSN & "RefreshFolderID") = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
				   If Trim(FileContent) = "" Then FileContent = "空间副模板不存在!"
				   FileContent=Replace(FileContent,"{$ShowMain}",GetLogList())
				   FileContent=KSR.KSLabelReplaceAll(FileContent)
		   Response.Write FileContent  
		End Sub
		Function GetLogList()
		  MaxPerPage=20
		  CurrPage=KS.ChkClng(Request("page"))
		  If CurrPage<=0 Then CurrPage=1
		  Dim RSObj,i,Str:Set RSObj=Server.CreateObject("adodb.recordset")
		  Dim SQLStr:SQLStr="select top 500 b.*,u.userface,u.realname from ks_bloginfo b inner join ks_user u on u.username=b.username where b.istalk=1 and b.status=0 order by b.id desc"
		  RSObj.Open SQLStr,conn,1,1
		  If RSObj.Eof And RSObj.Bof Then
		     str="没有新鲜事!"
		  ELSE
		     totalPut = RSObj.recordcount
			 If CurrPage > 1  and (CurrPage - 1) * MaxPerPage < totalPut Then
					RSObj.Move (CurrPage - 1) * MaxPerPage
			 Else
					CurrPage = 1
			 End If
			i=0
			str="<script src=""js/ks.space.js""></script>"
			str=str & "<table width='100%' border='0'>"
					Do While Not RSObj.Eof
								     Dim Uid:Uid=RSObj("UserID")
									 dim userfacesrc:userfacesrc=RSObj("userface")
									 if KS.IsNul(userfacesrc) then userfacesrc="../Images/Face/boy.jpg"
									 if left(userfacesrc,1)<>"/" and lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
									 dim username:username=rsobj("realname"): If KS.IsNul(username) Then UserName=rsobj("username")
								
									 str=str & "<tr class=""loglist"">"
									 str= str & "<td style='text-align:center;width:54px' class='splittd' valign='top'><div style='margin:5px 2px 5px 0px'><a href=""" &KS.GetSpaceUrl(uid) &""" target=""_blank""><img class='faceborder'  src='" & userfacesrc & "' width='40' height='40' border='0'/></a></div></td><td class='splittd'><a target=""_blank"" href=""" &KS.GetSpaceUrl(rsobj("userid")) &""" style=""font-size:16px;color:0F5FBB"">" & username & "</a> <span style='font-size:14px'>" &  KSR.ReplaceEmot(rsobj("content")) & "</span> " & KS.GetTimeFormat(rsobj("adddate"))& " - "
									 
									 
									 
									 Dim CmtNum:CmtNum=KS.ChkClng(rsobj("totalput"))
									 Dim CmtNumStr:CmtNumStr="(" & CmtNum & ")"
									 If CmtNum>0 Then CmtNumStr="(<span style='color:red'>" & CmtNum & "</span>)"
									 str=str & " <a href=""javascript:void(0)"" onclick=""showcmt(" & rsobj("id") & ")""> 评论" & CmtNumStr & "</a>"
									 

									 str=str & "<div id=""sc" & rsobj("id") & """ style="""
									 If i>0 Then str=str & "display:none;"
									 str=str &"padding:5px;margin-bottom:6px;margin-left:2px;width:400px;border:1px solid #C1DEFB;background:#E8EFF9;"">"
									 
									 If CmtNum>0 Then
									   Dim RSC:Set RSC=Conn.Execute("Select Top 3 C.AnounName,C.UserName,C.Content,C.Replay,C.replaydate,C.AddDate,U.UserFace,U.UserID,U.RealName From KS_BlogComment C Left Join KS_User U On C.AnounName=u.UserName Where C.LogID=" & KS.ChkClng(rsobj("id")))
									   If Not RSC.Eof Then 
									       Dim UserStr,Urls,Facestr
										   str=str & "<table width='100%' cellspacing='0' cellpadding='0'>"
										   str=str & "<tr><td class='splittd' colspan='2'>此条新鲜事共有 <span style='color:red'>" & CmtNum & "</span> 条评论，<a href='../space/?" & uid & "/log/" & KS.ChkClng(rsobj("id")) & "' target='_blank'>查看全部...</a></td></tr>"
										   Do While Not RSC.Eof
										    UserStr=RSC("AnounName")
											If KS.IsNul(UserStr) Then UserStr=RSC("UserName")
											UID=KS.ChkClng(RSC("UserID"))
											Dim RealName:RealName=RSC("RealName") : If KS.IsNul(RealName) Then RealName=RSC("AnounName")
											Facestr=RSC("UserFace") : If KS.IsNul(Facestr) Then Facestr="images/face/boy.jpg"
											 if left(Facestr,1)<>"/" and lcase(left(Facestr,4))<>"http" then Facestr="../" & Facestr
											If UID=0 Then Urls="#" Else Urls=KS.GetSpaceUrl(UID)
										    str=str & "<tr><td valign='top' class='splittd' style='width:50px;text-align:center;margin:5px 2px 2px 0px;'><img class='faceborder' src='" &facestr & "' width='40' height='40'/></td><td class='splittd' style=""width:300px""><a href='" & Urls & "'>" & RealName & "</a> " & KS.LoseHtml(RSC("Content")) & " " & KS.GetTimeFormat(RSC("Adddate")) 
											 If Not KS.IsNul(RSC("Replay")) Then
											 str=str & "<div style=""margin : 5px 20px; border : 1px solid #efefef; padding : 5px;background : #ffffee; line-height : normal;""><b>主人 <a href='" & KS.GetSpaceUrl(rsobj("userid")) & "' target='_blank'>" & username & "</a> 回复:</b><br>" & RSc("Replay") & "<br><div align=right>时间:" & rsc("replaydate") &"</div></div>"
											 End If

											str=str & "</td></tr>"
										   RSC.MoveNext
										   Loop
										   str=str & "</table>"
									   End If
										   RSC.Close : Set RSC=Nothing
									 End If
									 
									 
									 str=str &"<form  name=""form" & rsobj("id") & """ action=""../space/writecomment.asp"" method=""post""><input type=""hidden"" name=""action"" value=""CommentSave""/><input type=""hidden"" name=""id"" value=""" & rsobj("id") & """/><input type=""hidden"" name=""AnounName"" value=""" & KS.C("UserName") & """/><input type=""hidden"" name=""from"" value=""1""/>"
									 str=str &"<textarea name=""Content"" id=""c" &rsobj("id") & """ class=""cmttextarea"" onblur=""ThisBlur(" & rsobj("id") & ")"" onfocus=""ThisFocus(" & rsobj("id") & ")"" cols=""50"" rows=""2"">我也说一句...</textarea><br/><div style=""margin:4px 0px 4px 0px""><input type=""submit"" class=""btn"" onclick=""return(postcmt(" & rsobj("id") & "))"" value=""发表""/></div></form></div>"
									 
									 
									 
									 
									 

								 str=str &"</td></tr>"
								 I = I + 1
								  If I >= MaxPerPage Then Exit Do
								RSObj.MoveNext
								Loop
								str=str & "</table>"
			
		  End If
		  RSObj.Close:Set RSObj=Nothing
		  GetLogList=Str & KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
		  
		

		End Function
End Class
%>
