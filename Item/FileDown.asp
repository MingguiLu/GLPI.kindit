<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KS,KSUser
Set KS=New PublicCls
Dim ID,Node,Action,BSetting,LoginTF,Confirm,Score,LimitScore,FileName
ID = KS.ChkClng(KS.S("ID"))
Action=KS.G("Action")
Confirm=KS.G("Confirm")
If Action="hits" Then
   Set RS=Conn.Execute("Select top 1 hits From KS_UploadFiles Where ID=" &ID)
   If RS.Eof Then
     response.Write "document.write('0');"
   ELSE
     Response.Write "document.write('" & RS(0) & "');"
   End If
   RS.Close : Set RS=Nothing
Else
   Set KSUser=New UserCls
   LoginTF=KSUser.UserLoginChecked
   Set RS=Server.CreateObject("adodb.recordset")
   RS.Open "Select top 1 * From KS_UploadFiles Where ID=" & ID,conn,1,1
   If RS.Eof Then
     RS.Close : Set RS=Nothing
     KS.Die "<script>alert('附件已不存在!');history.back();</script>"
   Else
	   FileName=RS("FileName")
	   Dim ChannelID:ChannelID=KS.ChkClng(RS("ChannelID"))
	   Dim InfoID:InfoID=KS.ChkClng(RS("InfoID"))
	   Dim ClassID:ClassID=RS("ClassID")
	   Dim UserName:UserName=RS("UserName")
	   RS.Close : Set RS=Nothing
	   If ChannelID<5000 Then      '模型附件
	     Dim AnnexPoint:AnnexPoint=KS.ChkClng(KS.C_S(ChannelID,50))
		 If AnnexPoint<=0 Then
		   Call DownLoad()
		 Else
		   Dim ModelChargeType:ModelChargeType=KS.ChkClng(KS.C_S(ChannelID,34))
		   Call CheckConfirm(AnnexPoint,ModelChargeType)
		 End If
	   ElseIf ChannelID=9994 and ClassID<>0 Then  '论坛附件
	     KS.LoadClubBoard
		 Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & ClassID &"]")
		 If Node Is Nothing Then KS.Die "非法参数!"
		 BSetting=Node.SelectSingleNode("@settings").text
		 BSetting=BSetting & "$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
		 BSetting=Split(BSetting,"$")
		 LimitScore=KS.ChkClng(BSetting(15))
		 Score=KS.ChkClng(BSetting(16))
		 If (LimitScore>0 or Score>0) And LoginTF=false Then
		  KS.Die "<script>alert('本附件设置需要积分验证,请先登录!');location.href='" &KS.GetDomain & "user/login/';</script>"
		 End If
		 If LimitScore>0 and KS.ChkClng(KSUser.GetUserInfo("Score"))<LimitScore Then
		  KS.Die "<script>alert('对不起,本附件设置用户积分达到" & LimitScore & "分才可以下载,您当前积分"+KSUser.GetUserInfo("Score")+"!');window.close();</script>"
		 End If
		  Call CheckConfirm(Score,2)
	   End If
	   
	   DownLoad()
   End If 
 
End If

'权限下载附件并扣费处理
Sub CheckConfirm(Point,ModelChargeType)
  If Point<=0 Then DownLoad() : Exit Sub
	Dim ChargeStr,TableName,DateField,CurrPoint
	Select Case ModelChargeType
			case 0 ChargeStr=KS.Setting(46)&KS.Setting(45) : TableName="KS_LogPoint" : DateField="AddDate" : CurrPoint=KSUser.GetUserInfo("Point")
			case 1 ChargeStr="元人民币": TableName="KS_LogMoney" : DateField="PayTime": CurrPoint=KSUser.GetUserInfo("Money")
			case 2 ChargeStr="分积分": TableName="KS_LogScore": DateField="AddDate": CurrPoint=KSUser.GetUserInfo("Score")
			case else exit sub
	End Select
			
If Point>0 and KS.ChkClng(CurrPoint)<Point and ksuser.getedays<0 Then
		  KS.Die "<script>alert('对不起,下载本附件需要消费" & Point & ChargeStr & ",您当前剩余" & CurrPoint & ChargeStr&",不足支付!');window.close();</script>"
Else			
  If Conn.Execute("Select top 1 * From " & TableName & " Where UserName='" & KSUser.UserName & "' and datediff(" & DataPart_H &"," & DateField & "," & SqlNowString & ")<24 and ChannelID=9994 and InfoID=" & ID).Eof And KSUser.UserName<>UserName Then
		       If Confirm<>"true" Then
		    	KS.Die "<script>if(confirm('下载本附件需要消费" & Point & ChargeStr & ",确定下载吗?')){location.href='" & KS.GetDomain & "item/filedown.asp?confirm=true&id=" & id&"';}else{window.close();}</script>"
			   Else
			     Select Case ModelChargeType
				  case 0
					  IF Cbool(KS.PointInOrOut(9994,ID,KSUser.UserName,2,Point,"系统","下载附件[附件ID号:" & ID & "]!",0))=True Then 
					   DownLoad()
					  Else
					   KS.Die "<script>alert('扣费处理出错,请联系管理人员!');window.close();</script>"
					  End If
					  
				  case 1
					  IF Cbool(KS.MoneyInOrOut(KSUser.UserName,KSUser.UserName,Point,4,2,now,0,"系统","下载附件[附件ID号:" & ID & "]!",9994,ID,1))=True Then 
					   DownLoad()
					  Else
					   KS.Die "<script>alert('扣费处理出错,请联系管理人员!');window.close();</script>"
					  End If
				  case 2
					If Cbool(KS.ScoreInOrOut(KSUser.UserName,2,Point,"系统","下载附件[附件ID号:" & ID & "]!",9994,id)) Then
					  DownLoad()
					Else
					  KS.Die "<script>alert('扣费处理出错,请联系管理人员!');window.close();</script>"
					End If
				 end select
			   End If
  Else
		      DownLoad()
  End If
 End If
End Sub
Sub DownLoad()
       Conn.Execute("Update KS_UploadFiles Set Hits=Hits+1 Where ID=" & ID)
	   Response.Redirect FileName
End Sub
Call CloseConn()
Set KS=Nothing
Set KSUser=Nothing
%> 
