<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="payfunction.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New User_PayOnline
KSCls.Kesion()
Set KSCls = Nothing

Class User_PayOnline
        Private KS,KSUser
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		%>
		<!--#include file="../KS_Cls/UserFunction.asp"-->
		<%
		Public Sub LoadMain()
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		Call KSUser.Head()
		Call KSUser.InnerLocation("����֧��")
		Response.Write "<div class=""tabs"">"
		Response.Write " <ul class="""">"
		Response.Write " <li class='select'><a href=""User_PayOnline.asp"">����֧����ֵ</a></li>"
		Response.Write " <li><a href=""user_recharge.asp"">��ֵ����ֵ</a></li>"
		If Mid(KS.Setting(170),1,1)="1" or Mid(KS.Setting(170),2,1)="1" Then
		Response.Write " <li><a href=""user_exchange.asp?Action=Point"">�һ�" & KS.Setting(45) & "</a></li>"
		End If
		If Mid(KS.Setting(170),3,1)="1" or Mid(KS.Setting(170),4,1)="1" Then
		Response.Write " <li><a href=""user_exchange.asp?Action=Edays"">�һ���Ч��</a></li>"
		End If
		If Mid(KS.Setting(170),5,1)="1" Then
		Response.Write " <li><a href=""user_exchange.asp?Action=Money"">" & KS.Setting(45) & "�һ��˻��ʽ�</a></li>"
		End If
		Response.Write "</ul>"
		Response.Write "</div>"
		Select Case KS.S("Action")
		 Case "PayStep2"
		    Call PayStep2()
		 Case "PayStep3"
		    Call PayStep3()
		 Case "Payonline"
		    Call PayShopOrder()
	     Case Else
		    Call PayOnline()
		End Select
       End Sub
	  
	   
	   Sub PayOnline()
	    %>
	   <script type="text/javascript">
	     function Confirm(v)
		 {
		  $("#paytype").val(v);
		  if (v==1){
		    return(confirm('�˲��������棬ȷ��ʹ�����֧��������'));
		  }
		  if (document.myform.Money.value=="")
		  {
		   alert('��������Ҫ��ֵ�Ľ��!')
		   document.myform.Money.focus();
		   return false;
		  }
		  return true;
		  }
	   </script>
		<FORM name=myform action="User_PayOnline.asp" method="post">
		  <table class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr>
			  <td class=chargetitle colSpan=2 height=22>���߳�ֵ</td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=213>�û�����</td>
			  <td width="754"><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td width="213" align=right>�Ʒѷ�ʽ��</td>
			  <td><%if KSUser.ChargeType=1 Then 
		  Response.Write "�۵���</font>�Ʒ��û�"
		  ElseIf KSUser.ChargeType=2 Then
		   Response.Write "��Ч��</font>�Ʒ��û�,����ʱ�䣺" & cdate(KSUser.GetUserInfo("BeginDate"))+KSUser.GetUserInfo("Edays") & ","
		  ElseIf KSUser.ChargeType=3 Then
		   Response.Write "������</font>�Ʒ��û�"
		  End If
		  %>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=213>�ʽ���</td>
			  <td><input type='hidden' value='<%=KSUser.GetUserInfo("Money")%>' name='Premoney'><%=formatnumber(KSUser.GetUserInfo("Money"),2,-1)%> Ԫ</td>
			</tr>
			<%If KSUser.ChargeType=1 then%>
			<tr class=tdbg>
			  <td align=right width=213>����<%=KS.Setting(45)%>��</td>
			  <td><%=KSUser.GetUserInfo("Point")%>&nbsp;<%=KS.Setting(46)%></td>
			</tr>
			<%end if%>
			<%If KSUser.ChargeType=2 then%>
			<tr class=tdbg>
			  <td align=right width=213>ʣ��������</td>
			  <td>
			  <%if KSUser.ChargeType=3 Then%>
			  ������
			  <%else%>
			  <%=KSUser.GetEdays%>&nbsp;��
			  <%end if%></td>
			</tr>
		   <%end if%>
			<tr class=tdbg>
			  <td align=right>��ǰ����</td>
			  <td><%=KS.U_G(KSUser.GroupID,"groupname")%></td>
		    </tr>
			<tr>
			  <td class=chargetitle colSpan=2 height=22>ѡ�����߳�ֵ��ʽ</td>
			</tr>

			<tr class=tdbg>
			  <td colspan="2">
			  <%
			   Dim HasCard:HasCard=false
			   Dim RSC,AllowGroupID:Set RSC=Conn.Execute("Select ID,GroupName,Money,AllowGroupID From KS_UserCard Where CardType=1 and DateDiff(" & DataPart_S & ",EndDate," & SqlNowString& ")<0")
			   Do While NOt RSC.Eof 
			      HasCard=true
			      AllowGroupID=RSC("AllowGroupID") : If IsNull(AllowGroupID) Then AllowGroupID=" "
			     If KS.IsNul(AllowGroupID) Or KS.FoundInArr(AllowGroupID,KSUser.GroupID,",")=true Then
			    response.write "&nbsp;&nbsp; <label><input checked name=""UserCardID"" onclick=""$('#m').hide();$('#paybutton').attr('disabled',false);"" type=""radio"" value=""" & rsc("ID") & """/>" & rsc(1) & " (��Ҫ���� <span style='color:red'>" & formatnumber(RSC(2),2,-1) & "</span> Ԫ)</label><br/>"
				End If
			    RSC.MoveNext
			   Loop
			   RSC.Close
			   Set RSC=Nothing
			  %>
			  <%If Mid(KS.Setting(170),6,1)="1" Then%>
			  &nbsp;&nbsp; <label><input onClick="$('#m').show();$('#paybutton').attr('disabled',true);" type="radio" value="0" name="UserCardID">���ɳ�(��������������Ҫ��ֵ�Ľ��)</label><br/>
			  <%end if%>
			  <span id='m' style="display:none"> &nbsp;&nbsp;&nbsp;&nbsp;��������Ҫ��ֵ�Ľ�&nbsp;<input style="text-align:center;line-height:22px" name="Money" type="text" class="textbox" value="100" size="10" maxlength="10"> Ԫ</span>
			  </td>
		    </tr>
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
		        <Input id="Action" type="hidden" value="PayStep2" name="Action"> 
				<Input class="button" id=Submit type=submit value=" ��������֧�� " onClick="return(Confirm(0))" name=Submit>
				<%if HasCard then%>
				<input type='hidden' name='paytype' id='paytype' value='1'/>
				<Input class="button" id="paybutton" type=submit value=" ʹ�����֧�� " onClick="return(Confirm(1))"  name=Submit>
				<%end if%>
				 </td>
			</tr>
		  </table>
		</FORM>
		<br/><br/>
	   <%
	   End Sub
	   
	   Sub PayStep2()
	    Dim UserCardID:UserCardID=KS.ChkClng(KS.G("UserCardID"))
	   	Dim Money:Money=KS.S("Money")
		Dim Title,PayType
		PayType=KS.ChkClng(KS.S("PayType"))
		
		If UserCardID<>0 Then
		   Dim RS:Set RS=Conn.Execute("Select Top 1 Money,GroupName From KS_UserCard Where ID=" & UserCardID)
		   If Not RS.Eof Then
		    Title=RS(1)
		    Money=RS(0)
			RS.Close : Set RS=Nothing
		   Else
		    RS.Close : Set RS=Nothing
		    Call KS.AlertHistory("��������",-1)
			Exit Sub 
		   End If
		   
		   '�ж��û���û��ѡ������
		   If PayType=1 Then
		     If round(KSUser.GetUserInfo("money"))<round(Money) Then
		      Call KS.AlertHistory("�Բ��������ý��㣬����ֵ����Ҫ����" & Money & "Ԫ������ǰ�Ŀ������Ϊ" & Formatnumber(KSUser.GetUserInfo("money"),2,-1,-1) & "Ԫ����ѡ�����߹���֧����",-1)
			  Exit Sub
			 End If
			 Call UpdateByCard(1,UserCardID,KSUser.UserName,KSUser.GetUserInfo("RealName"),KSUser.GetUserInfo("Edays"),KSUser.GetUserInfo("BeginDate"),UserCardID,"")
			 Session(KS.SiteSN&"UserInfo")=empty
			 Response.Write("<script>alert('��ϲ��[" & title & "]����ɹ���');location.href='user_logmoney.asp';</script>")
			 response.End()
		   End If 
		   
		   
		ElseIf Mid(KS.Setting(170),6,1)="0" Then
		  KS.AlertHintScript "�Բ��𣬱�վ�������Ա���ɳ�ֵ��"
		  Exit Sub
		Else
		   Title="Ϊ�Լ����˻���ֵ"
		End If

		If Not IsNumeric(Money) Then
		  Call KS.AlertHistory("�Բ���������ĳ�ֵ����ȷ��",-1)
		  exit sub
		End If
		
		If Money=0 Then
		  Call KS.AlertHistory("�Բ��𣬳�ֵ������Ϊ0.01Ԫ��",-1)
		  exit sub
		End If
		Dim OrderID:OrderID=KS.Setting(72) & Year(Now)&right("0"&Month(Now),2)&right("0"&Day(Now),2)&hour(Now)&minute(Now)&second(Now)
		
		%>
	   <FORM name=myform action="User_PayOnline.asp" method="post">
		  <table id="c1" class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> ȷ �� �� ��</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=167>�û�����</td>
			  <td width="505"><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td width="167" align=right>֧����ţ�</td>
			  <td><input type='hidden' value='<%=OrderID%>' name='OrderID'><%=OrderID%>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=167>֧����</td>
			  <td><input type='hidden' value='<%=Money%>' name='Money'><%=FormatNumber(Money,2,-1)%> Ԫ</td>
			</tr>
			<%If title<>"" then%>
			<tr class=tdbg>
			  <td align=right width=167>֧����;��</td>
			  <td style="color:red">��<%=title%>��</td>
			</tr>
			<%end if%>

			<tr class=tdbg>
			  <td align=right width=167>ѡ������֧��ƽ̨��</td>
			  <td>
			  <%
			   Dim SQL,K,Param
			   If UserCardID<>0 Then
			    Param=" and id in(1,10,7,12,13,6)"
			   End IF
			   Set RS=Server.CreateOBject("ADODB.RECORDSET")
			   RS.Open "Select ID,PlatName,Note,IsDefault From KS_PaymentPlat Where IsDisabled=1 " & Param & " Order By OrderID",conn,1,1
			   If Not RS.Eof Then SQL=RS.GetRows(-1)
			   RS.Close:Set RS=Nothing
			   If Not IsArray(SQL) Then
			    Response.Write "<font color='red'>�Բ��𣬱�վ�ݲ���ͨ����֧�����ܣ�</font>"
			   Else
			     For K=0 To Ubound(SQL,2)
				   Response.Write "<input type='radio' value='" & SQL(0,K) & "' name='PaymentPlat'"
				   If SQL(3,K)="1" Then Response.Write " checked"
				   Response.Write ">"& SQL(1,K) & "(" & SQL(2,K) &")<br>"
				 Next
			   End If
			  %>
			  </td>
			</tr>
			
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
		        <Input id=Action type=hidden value="PayStep3" name="Action"> 
		        <Input id=Action type=hidden value="<%=UserCardID%>" name="UserCardID"> 
		        <Input type=hidden value="user" name="PayFrom"> 
				<input class="button" type="button" value=" ��һ�� " onClick="javascript:history.back();"> 
				<Input class="button" id=Submit type=submit value=" ��һ�� " name=Submit>
				</td>
			</tr>
		  </table>
		</FORM>
		<%
	   End Sub
	   
	   
	   '֧���̳Ƕ���
	   Sub PayShopOrder()
	  	 Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 OrderID,MoneyTotal,DeliverType From KS_Order Where ID="& ID,Conn,1,1
		 If RS.Eof Then
		  rs.close:set rs=nothing
		  KS.Die "<script>alert('������!');history.back();</script>"
		 End If 
		Dim OrderID:OrderID=RS("OrderID")
	   	Dim Money:Money=RS("MoneyTotal")
		Dim DeliverType:DeliverType=RS("DeliverType")
		RS.Close
		Dim DeliverName,ProductName
		RS.Open "Select Top 1 TypeName From KS_Delivery Where Typeid=" & DeliverType,conn,1,1
		If Not RS.Eof Then
		 DeliverName=RS(0)
		End IF
		RS.Close
		
		RS.Open "Select top 10 Title From KS_Product Where ID in(Select proid From KS_OrderItem Where OrderID='" & OrderID& "')",conn,1,1
		If RS.Eof And RS.Bof Then
		 ProductName=OrderID
		Else
			Do While Not RS.Eof
			 if ProductName="" Then
			   ProductName=rs(0)
			 Else
			   ProductName=ProductName&","&rs(0)
			 End If
			 RS.MoveNext
			Loop
		End If
		RS.Close
		
		If Not IsNumeric(Money) Then
		  Call KS.AlertHistory("�Բ��𣬶�������ȷ��",-1)
		  exit sub
		End If
		If Money=0 Then
		  Call KS.AlertHistory("�Բ��𣬶���������Ϊ0.01Ԫ��",-1)
		  exit sub
		End If
		%>
	   <FORM name=myform action="User_PayOnline.asp" method="post">
		  <table id="c1" class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> ȷ �� �� ��</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=167>�û�����</td>
			  <td width="505"><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td width="167" align=right>��Ʒ���ƣ�</td>
			  <td><input type='hidden' value='<%=ProductName%>' name='ProductName'><%=ProductName%>&nbsp;
			  <input type='hidden' value='<%=DeliverName%>' name='DeliverName'>
			  </td>
		    </tr>
			<tr class=tdbg>
			  <td width="167" align=right>֧����ţ�</td>
			  <td><input type='hidden' value='<%=OrderID%>' name='OrderID'><%=OrderID%>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=167>֧����</td>
			  <td>
			  <%If KS.Setting(82)<>"0" And IsNumeric(KS.Setting(82)) Then%>
			   <input type="hidden" name="zfdj" value="1">
			   <strong>�����ܽ�� <span style='color:red'><%=Money%></span> Ԫ ,������������Ҫ��֧��<span style='color:red'> <%=KS.Setting(82)%></span> Ԫ����</strong>
			   <br/> ��������Ҫ֧���Ľ�<input type='text' size=6 value='<%=Money%>' name='Money'> Ԫ
			  <%else%>
			  <input type='hidden' value='<%=Money%>' name='Money'><%=Money%> Ԫ
			  <%end if%>
			  </td>
			</tr>
			
			<tr class=tdbg>
			  <td align=right width=167>ѡ������֧��ƽ̨��</td>
			  <td>
			  <%
			   Dim SQL,K
			   RS.Open "Select ID,PlatName,Note,IsDefault From KS_PaymentPlat Where IsDisabled=1 Order By OrderID",conn,1,1
			   If Not RS.Eof Then SQL=RS.GetRows(-1)
			   RS.Close:Set RS=Nothing
			   If Not IsArray(SQL) Then
			    Response.Write "<font color='red'>�Բ��𣬱�վ�ݲ���ͨ����֧�����ܣ�</font>"
			   Else
			     For K=0 To Ubound(SQL,2)
				   Response.Write "<input type='radio' value='" & SQL(0,K) & "' name='PaymentPlat'"
				   If SQL(3,K)="1" And KS.ChkClng(KS.S("PaymentPlat"))=0 Then Response.Write " checked"
				   iF KS.ChkClng(SQL(0,K))=KS.ChkClng(KS.S("PaymentPlat")) Then Response.Write " checked"
				   Response.Write ">"& SQL(1,K) & "(" & SQL(2,K) &")<br>"
				 Next
			   End If
			  %>
			  </td>
			</tr>
			
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
		        <Input id=Action type=hidden value="PayStep3" name="Action"> 
		        <Input type=hidden value="shop" name="PayFrom"> 
				<Input class="button" id=Submit type=submit value=" ��һ�� " name=Submit>
				<input class="button" type="button" value=" ��һ�� " onClick="javascript:history.back();"> </td>
			</tr>
		  </table>
		</FORM>
		<%
	   End Sub
	   
	   Sub PayStep3()
	    Dim UserCardID,Title
		UserCardID=KS.ChkClng(KS.S("UserCardID"))
	    Dim Money:Money=KS.S("Money")
		If KS.S("zfdj")="1" Then
		  If Not IsNumeric(Money) Then
		    KS.AlertHintScript "����ȷ!"
		  ElseIf  Money<KS.Setting(82) Then
		    KS.AlertHIntScript "������������ҪԤ�� " & KS.Setting(82) & " Ԫ�Ķ���!"
		  End If
		End If
		Dim OrderID:OrderID=KS.S("OrderID")
		Dim ProductName:ProductName=KS.CheckXSS(KS.S("ProductName"))
		Dim PaymentPlat:PaymentPlat=KS.ChkClng(KS.S("PaymentPlat"))
		Dim PayUrl,PayMentField,ReturnUrl,RealPayMoney,RealPayUSDMoney,RateByUser,PayOnlineRate
        Call GetPayMentField(OrderID,PaymentPlat,Money,UserCardID,ProductName,KS.S("PayFrom"),KSUser,PayMentField,PayUrl,ReturnUrl,Title,RealPayMoney,RealPayUSDMoney,RateByUser,PayOnlineRate)
		
		 %>
	   	  <FORM name="myform"  id="myform" action="<%=PayUrl%>" <%if PaymentPlat=11 or PaymentPlat=9 then response.write "method=""get""" else response.write "method=""post"""%>  target="_blank">
		  <table id="c1" class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> ȷ �� �� ��</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=167>�û�����</td>
			  <td width="505"><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td width="167" align=right>֧����ţ�</td>
			  <td><%=OrderID%>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=167>֧����</td>
			  <td><%=formatnumber(Money,2,-1)%> Ԫ</td>
			</tr>
			<%if title<>"" then%>
			<tr class=tdbg>
			  <td align=right width=167>֧����;��</td>
			  <td style="color:red">��<%=title%>��</td>
			</tr>
			<%end if%>
			<%
			if RateByUser=1 then
			%>
			<tr class=tdbg>
			  <td align=right width=167>�����ѣ�</td>
			  <td><%=PayOnlineRate%>%</td>
			</tr>
			<%end if%>
			<tr class=tdbg>
			  <td align=right width=167>ʵ��֧����</td>
			  <td>
			  <%=formatnumber(RealPayMoney,2,-1)%></td>
			</tr>
			<%If PaymentPlat=12 Then%>
			<tr class=tdbg>
			  <td align=right width=167>ʵ��֧������</td>
			  <td style="color:#FF6600;font-weight:bold">
			  $<%=formatnumber(RealPayUSDMoney,2,-1)%> USD</td>
			</tr>
			<%End If%>
			<tr class=tdbg>
			  <td colspan=2>�����ȷ��֧������ť�󣬽���������֧�����棬�ڴ�ҳ��ѡ���������п���</td>
		    </tr>
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
			    <%=PayMentField%>
				<%if PaymentPlat=9 then%>
				<Input class="button" id=Submit type=button onClick="document.all.c1.style.display='none';document.all.c2.style.display='';$('#myform').submit()" value=" ȷ��֧�� ">
				<%else%>
				<Input class="button" id=Submit type=submit value=" ȷ��֧�� " onClick="document.all.c1.style.display='none';document.all.c2.style.display='';">
				<%end if%>
				<input class="button" type="button" value=" ��һ�� " onClick="javascript:history.back();"> </td>
			</tr>
		  </table>
		</FORM>
		  <table id="c2" style="display:none" class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle height=22><B> ȷ �� �� ��</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=center height="150">�밴ҳ����ʾ�������ֵ��</td>
			</tr>
          </table>
	   <%
	   End Sub
		
End Class
%> 
