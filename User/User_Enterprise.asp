<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New User_EditInfo
KSCls.Kesion()
Set KSCls = Nothing

Class User_EditInfo
        Private KS,KSUser
		Private FieldsXml,Action
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
       Public Sub loadMain()
		if KS.SSetting(0)<>1 then
		  Response.Write "<script>alert('ϵͳû�п�ͨ�ռ书��!');history.back();</script>"
		  Response.end
		End If
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		Action=Request("action")
		Call KSUser.Head()
		%>
		<script src="../ks_inc/kesion.box.js" language="JavaScript"></script>

		<div class="tabs">	
			<ul>
			<li><a href="User_EditInfo.asp">������Ϣ</a></li>
			<li><a href="User_EditInfo.asp?Action=face">����ͷ��</a></li>
			<li><a href="User_EditInfo.asp?Action=ContactInfo">������ϸ����</a></li>
			<li><a href="User_EditInfo.asp?Action=PassInfo">��������</a></li>
	        <li<%if action="" then response.write " class='select'"%>><a href="user_enterprise.asp">��ҵ�������� </a></li>
	        <li<%if action="intro" then response.write " class='select'"%>><a href="?action=intro">��ҵ���</a></li>
			<%if action="job" then
			 if KS.C_S(10,21)="0" then response.write "<li class='select'><a href='?action=job'>��ҵ��Ƹ</a></li>"
			end if%>
			</ul>
			
		</div>

		<%
		Dim HasEnterprise:HasEnterprise=Not Conn.execute("select top 1 id from KS_Enterprise where username='" & KSUser.UserName & "'").eof
		Set FieldsXml=LFCls.GetXMLFromFile("SpaceFields")
		Select Case KS.S("Action")
		  Case "BasicInfoSave"
		   Call BasicInfoSave()
		  Case "intro"
		   If (HasEnterprise) then
	        Call KSUser.InnerLocation("��ҵ���")
		    Call Intro()
		   Else
		    Response.Write "<script>alert('�Բ����㻹û����д��ҵ������Ϣ!')</script>"
	       Call KSUser.InnerLocation("��ҵ������Ϣ")
		   Call EditBasicInfo()
		   End If
		  case "IntroSave"
		   Call IntroSave()
		  Case "job"
		   If (HasEnterprise) then
	        Call KSUser.InnerLocation("��ҵ��Ƹ")
			If KS.C_S(10,21)="1" Then
			 Response.Redirect("User_JobCompanyZW.asp")
			Else
		    Call Job()
			End If
		   Else
		    Response.Write "<script>alert('�Բ����㻹û����д��ҵ������Ϣ!')</script>"
	       Call KSUser.InnerLocation("��ҵ������Ϣ")
		   Call EditBasicInfo()
		   End If
		  Case "JobSave"
		   Call JobSave()
		  Case Else
	       Call KSUser.InnerLocation("��ҵ������Ϣ")
		   Call EditBasicInfo()
		End Select
	   End Sub
	   
	   '������Ϣ
	   Sub EditBasicInfo()
		   %>
      <script>
       function CheckForm() 
		{ 
			
			if (document.myform.CompanyName.value =="")
			{
			alert("����д��˾���ƣ�");
			document.myform.CompanyName.focus();
			return false;
			}
			if (document.myform.LegalPeople.value =="")
			{
			alert("����д��ҵ���ˣ�");
			document.myform.LegalPeople.focus();
			return false;
			}
			if (document.myform.TelPhone.value =="")
			{
			alert("��������ϵ�绰��");
			document.myform.TelPhone.focus();
			return false;
			}
		  return true;	
		}
		
    </script>
	<%	   

	 Dim CompanyName,Province,City,Address,ZipCode,ContactMan,Telphone,Fax,WebUrl,Profession,CompanyScale,RegisteredCapital,LegalPeople,BankAccount,AccountNumber,BusinessLicense,Intro,flag,ClassID,SmallClassID,qq,mobile,Email,MapMarker
	 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	 RS.Open "Select top 1 * From KS_Enterprise where username='" & KSUser.UserName & "'",conn,1,1
	 IF Not RS.Eof Then
	   CompanyName=RS("CompanyName")
	   Province=RS("Province")
	   City=RS("City")
	   Address=RS("Address")
	   ZipCode=RS("ZipCode")
	   ContactMan=RS("ContactMan")
	   Telphone=RS("TelPhone")
	   Fax=RS("Fax")
	   WebUrl=RS("WebUrl")
	   Profession=RS("Profession")
	   CompanyScale=RS("CompanyScale")
	   RegisteredCapital=RS("RegisteredCapital")
	   LegalPeople=RS("LegalPeople")
	   BankAccount=RS("BankAccount")
	   AccountNumber=RS("AccountNumber")
	   BusinessLicense=RS("BusinessLicense")
	   ClassID=RS("ClassID")
	   SmallClassID=RS("SmallClassID")
	   qq=rs("qq")
	   MapMarker=rs("MapMarker")
	   Email=rs("Email")
	   mobile=rs("mobile")
	   flag=true
	 Else
	   flag=false
	    if KS.SSetting(17)<>"" then
	    if KS.FoundInArr(KS.SSetting(17),KSUser.groupid,",")=false then  Set KSUser=Nothing:call KS.AlertHistory("�Բ��������ڵ��û���û��Ȩ������Ϊ��ҵ�ռ䣡",-1):exit sub
	   end if
	   If IsObject(FieldsXml) Then
	     on error resume next
	     Dim objNode,i,j,objAtr
	     Set objNode=FieldsXml.documentElement 
		 For i=0 to objNode.ChildNodes.length-1 
				set objAtr=objNode.ChildNodes.item(i) 
				' response.write objAtr.Attributes.item(0).Text&"=" &objAtr.Attributes.item(1).Text & " <br>" 
				 Execute(objAtr.Attributes.item(0).Text&"=""" & LFCls.GetSingleFieldValue("select " & objAtr.Attributes.item(1).Text & " From KS_User Where UserName='" & KSUser.UserName & "'") & """") 
		 Next

	   End If
	   
	 End If
	 If ClassID="" or isnull(ClassID) Then  ClassID=0
	 If SmallClassID="" or isnull(ClassID) Then SmallClassID=0

    RS.Close:Set RS=Nothing	
	%>
          
          <table  cellspacing="1" cellpadding="3"  width="98%" align="center" border="0">
					  <form action="?Action=BasicInfoSave" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
					      <input type="hidden" value="<%=KS.S("ComeUrl")%>" name="ComeUrl">
                          <tr class="tdbg">
                            <td class="clefttitle">��˾���ƣ�</td>
                            <td><input name="CompanyName" type="text" class="textbox" id="CompanyName" value="<%=CompanyName%>" size="30" maxlength="200" />
                                <span style="color: red">* </span> <span class="msgtips">����д���ڹ��̾�ע��Ǽǵ����ơ�</span></td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle">Ӫҵ���գ�</td>
                            <td><input name="BusinessLicense" class="textbox" type="text" id="BusinessLicense" value="<%=BusinessLicense%>" size="30" maxlength="50" /> <span class="msgtips">��д���Ӫҵִ��ͼƬ���ڵ�ַ��Ӫҵִ�պ��롣</span></td>
                          </tr>
                         <tr class="tdbg">
                            <td class="clefttitle">��˾��ҵ��</td>
                            <td><%
		dim rss,sqls,count
		set rss=server.createobject("adodb.recordset")
		sqls = "select * from KS_enterpriseClass Where parentid<>0 order by orderid"
		rss.open sqls,conn,1,1
		%>
          <script language = "JavaScript">
		var onecount;
		subcat = new Array();
				<%
				count = 0
				do while not rss.eof 
				%>
		subcat[<%=count%>] = new Array("<%= trim(rss("id"))%>","<%=trim(rss("parentid"))%>","<%= trim(rss("classname"))%>");
				<%
				count = count + 1
				rss.movenext
				loop
				rss.close
				%>
		onecount=<%=count%>;
		function changelocation(locationid)
			{
			document.myform.SmallClassID.length = 0; 
			for (var i=0;i < onecount; i++)
				{ 
					if (parseInt(subcat[i][1]) == parseInt(locationid))
					{ 			
						document.myform.SmallClassID.options[document.myform.SmallClassID.length] = new Option(subcat[i][2], subcat[i][0]);
					}        
				}
			}    
		
		</script>
		  <select class="face" name="ClassID" onChange="changelocation(document.myform.ClassID.options[document.myform.ClassID.selectedIndex].value)" size="1">
		<% 
		dim rsb,sqlb
		set rsb=server.createobject("adodb.recordset")
        sqlb = "select * from ks_enterpriseClass where parentid=0 order by orderid"
        rsb.open sqlb,conn,1,1
		if rsb.eof and rsb.bof then
		else
		    Dim N
		    do while not rsb.eof
			          N=N+1
					  If N=1 and flag=false Then ClassID=rsb("id")
					  If ClassID=rsb("id") then
					  %>
                    <option value="<%=trim(rsb("id"))%>" selected><%=trim(rsb("ClassName"))%></option>
                    <%else%>
                    <option value="<%=trim(rsb("id"))%>"><%=trim(rsb("ClassName"))%></option>
                    <%end if
		        rsb.movenext
    	    loop
		end if
        rsb.close
			%>
                  </select>
                  <font color=#ff6600>&nbsp;*</font>
                  <select class="face" name="SmallClassID">
                    <%dim rsss,sqlss
						set rsss=server.createobject("adodb.recordset")
						sqlss="select * from ks_enterpriseclass where parentid="&ClassID&" order by orderid"
						rsss.open sqlss,conn,1,1
						if not(rsss.eof and rsss.bof) then
						do while not rsss.eof
							  if SmallClassID=rsss("id") then%>
							<option value="<%=rsss("id")%>" selected><%=rsss("ClassName")%></option>
							<%else%>
							<option value="<%=rsss("id")%>"><%=rsss("ClassName")%></option>
							<%end if
							rsss.movenext
						loop
					end if
					rsss.close
					%>
                </select>
							 <span class="msgtips">��д��˾��������ҵ��</span> 
							  </td>
                          </tr>
						  
                          <tr class="tdbg">
                            <td class="clefttitle">��ҵ���ˣ�</td>
                            <td><input name="LegalPeople" class="textbox" type="text" id="LegalPeople" value="<%=LegalPeople%>" size="30" maxlength="50" />
                            <span style="color: red">* </span> <span class="msgtips">��д��˾�ķ��˻�����Ҫ�����ˡ�</span></td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle">��˾��ģ��</td>
                            <td><select name="CompanyScale" id="CompanyScale">
							  <option value="1-20��"<%if CompanyScale="1-20��" then response.write " selected"%>>1-20��</option>
                      <option value="21-50��"<%if CompanyScale="21-50��" then response.write " selected"%>>21-50��</option>
                      <option value="51-100��"<%if CompanyScale="51-100��" then response.write " selected"%>>51-100��</option>
                      <option value="101-200��"<%if CompanyScale="101-200��" then response.write " selected"%>>101-200��</option>
                      <option value="201-500��"<%if CompanyScale="201-500��" then response.write " selected"%>>201-500��</option>
                      <option value="501-1000��"<%if CompanyScale="501-1000��" then response.write " selected"%>>501-1000��</option>
                      <option value="1000������"<%if CompanyScale="1000������" then response.write " selected"%>>1000������</option>
						    </select>
							<span class="msgtips">��ѡ��˾��Ա������</span>
							</td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle">ע���ʽ�</td>
                            <td><select name="RegisteredCapital" id="RegisteredCapital">
							<option value="10������"<%if RegisteredCapital="10������" then response.write " selected"%>>10������</option>
                      <option value="10��-19��"<%if RegisteredCapital="10��-19��" then response.write " selected"%>>10��-19��</option>
                      <option value="20��-49��"<%if RegisteredCapital="20��-49��" then response.write " selected"%>>20��-49��</option>
                      <option value="50��-99��"<%if RegisteredCapital="50��-99��" then response.write " selected"%>>50��-99��</option>
                      <option value="100��-199��"<%if RegisteredCapital="100��-199��" then response.write " selected"%>>100��-199��</option>
                      <option value="200��-499��"<%if RegisteredCapital="200��-499��" then response.write " selected"%>>200��-499��</option>
                      <option value="500��-999��"<%if RegisteredCapital="500��-999��" then response.write " selected"%>>500��-999��</option>
                      <option value="1000������"<%if RegisteredCapital="1000������" then response.write " selected"%>>1000������</option>
					   </select> <span class="msgtips">��ѡ���˾��ע���ʽ�</span></td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle">���ڵ�����</td>
                            <td><script src="../plus/area.asp" language="javascript"></script>
							<script language="javascript">
							  <%if Province<>"" then%>
							  $('#Province').val('<%=province%>');
								  <%end if%>
							  <%if City<>"" Then%>
							  $('#City')[0].options[1]=new Option('<%=City%>','<%=City%>');
							  $('#City')[0].options(1).selected=true;
							  <%end if%>
							</script>
							  <span class="msgtips">ѡ����ҵ���ڵ�ʡ�ݺͳ��С�</span>
							  </td>
                          </tr>
						  <tr class="tdbg">
                            <td class="clefttitle">���ӵ�ͼ��</td>
                            <td>��γ�ȣ�<input value="<%=MapMarker%>" class="textbox" maxlength="255" type='text' name='MapMark' id='MapMark' /> <a href='javascript:void(0)' onclick='addMap()'> <img src='images/edit_add.gif' align='absmiddle' border='0'>��ӵ��ӵ�ͼ��־</a>
	 <script type="text/javascript">
		  function addMap(){
		  new KesionPopup().PopupCenterIframe('���ӵ�ͼ��ע','../plus/baidumap.asp?MapMark='+escape($("#MapMark").val()),760,430,'auto');
		  }
		  </script><span class="msgtips">��ѡ���˾���ڵ�λ�á�</span>
							  </td>
                          </tr>

                          <tr class="tdbg">
                            <td class="clefttitle">�� ϵ �ˣ�</td>
                            <td><input name="ContactMan" class="textbox" type="text" id="ContactMan" value="<%=ContactMan%>" size="30" maxlength="50" /> <span class="msgtips">����ȷ��дҵ����ϵ�˵�������</span></td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle">��˾��ַ��</td>
                            <td><input name="Address" class="textbox" type="text" id="Adress" value="<%=Address%>" size="30" maxlength="50" /> <span class="msgtips">��д��˾����ϵ��ַ</span></td>
                          </tr>
       
                          <tr class="tdbg">
                            <td class="clefttitle">�������룺</td>
                            <td><input name="ZipCode" class="textbox" type="text" id="ZipCode" value="<%=ZipCode%>" size="10" maxlength="6" /> <span class="msgtips">����д�������롣</span> </td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle"> QQ���룺</td>
                            <td><input name="qq" class="textbox" type="text" id="qq" value="<%=qq%>" size="10" maxlength="50" />
                            <span style="color: red">* </span> <span class="msgtips">����ҵ����ϵ����д��ϵQQ��</span></td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle"> �ֻ����룺</td>
                            <td><input name="Mobile" class="textbox" type="text" id="Mobile" value="<%=Mobile%>" size="30" maxlength="50" />
                           </td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle"> ��ϵ�绰��</td>
                            <td><input name="TelPhone" class="textbox" type="text" id="TelPhone" value="<%=Telphone%>" size="30" maxlength="50" />
                            <span style="color: red">* </span> <span class="msgtips">������,��˾�칫�绰������ҵ����ϵ��</span></td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle"> ������룺</td>
                            <td><input name="Fax" class="textbox" type="text" id="Fax" value="<%=Fax%>" size="30" maxlength="50" /> <span class="msgtips">��˾�Ĵ�����롣</span></td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle"> �������䣺</td>
                            <td><input name="Email" class="textbox" type="text" id="Email" value="<%=Email%>" size="30" maxlength="50" /></td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle">��˾��վ��</td>
                            <td><input name="WebUrl" class="textbox" type="text" id="WebUrl" value="<%=WebUrl%>" size="30" maxlength="50" /> <span class="msgtips">��д�㹫˾����ַ��</span></td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle">�������У�</td>
                            <td><input name="BankAccount" class="textbox" type="text" id="BankAccount" value="<%=BankAccount%>" size="30" maxlength="50" /></td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle">�����˺ţ�</td>
                            <td><input name="AccountNumber" class="textbox" type="text" id="AccountNumber" value="<%=AccountNumber%>" size="30" maxlength="50" /> <span class="msgtips">��˾�����ʻ����Է������������ϵ�����С�</span></td>
                          </tr>
                          <tr class="tdbg">
                            <td class="clefttitle">&nbsp;</td>
                            <td><button class="pn" name="Submit" type="submit"><strong>OK,ȷ �� �� ��</strong></button></td>
                          </tr>
		    </form>
            </table>
          <%
  End Sub
  
  Sub Intro()
  %>
   <table  cellspacing="1" cellpadding="3" width="98%" align="center" border="0">
			<form action="?Action=IntroSave" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
               <tr class="tdbg">
                  <td class="msgtips">
				  ������������ϸ˵����˾�ĳ�����ʷ����Ӫ��Ʒ��Ʒ�ơ���������ƣ�<br>
 ��������ݹ��ڼ򵥻����д�����Ĳ�Ʒ���ܣ����п����޷�ͨ����ˡ�<br>
����ϵ��ʽ���绰�����桢�ֻ�����������ȣ����ڻ�����������д�� �˴������ظ���д��<br>
                    <%
					Dim Intro:Intro=Conn.Execute("Select Intro From ks_Enterprise where username='" & KSUser.UserName & "'")(0)
					If trim(Intro)="" Or IsNull(Intro) Then
						If IsObject(FieldsXml) Then
						 'on error resume next
						 Dim objNode,i,j,objAtr
						 Set objNode=FieldsXml.documentElement 
						 For i=0 to objNode.ChildNodes.length-1 
								set objAtr=objNode.ChildNodes.item(i)
								If lcase(objAtr.Attributes.item(0).Text)="intro" Then 
								 Intro=LFCls.GetSingleFieldValue("select " & objAtr.Attributes.item(1).Text & " From KS_User Where UserName='" & KSUser.UserName & "'") 
								End If
						 Next
				
					   End If
					End If
					
			        Response.Write "<textarea ID='Intro' name='Intro' style='display:none'>" & KS.HTMLCode(Intro) & "</textarea>"
					%> 
					<script type="text/javascript" src="../editor/ckeditor.js" mce_src="../editor/ckeditor.js"></script>
					<script type="text/javascript">
                CKEDITOR.replace('Intro', {width:"98%",height:"350px",toolbar:"Simple",filebrowserBrowseUrl :"../editor/ksplus/SelectUpFiles.asp",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
			    </script>  
					</td>
                          </tr>
						  <tr class="tdbg">
                            <td align="center"><button class="pn" name="Submit" type="submit"><strong>OK,ȷ �� �� ��</strong></button></td>
                          </tr>
				</form>
	</table>
  <%
  End Sub
 
 Sub IntroSave()
  Dim Intro
  Intro = Request.Form("Intro")
  Intro=KS.CheckScript(KS.HtmlCode(Intro))
  Intro=KS.HtmlEncode(Intro)
  IF Intro="" Then
  	 Response.Write "<script>alert('�Բ�����û�����빫˾���');history.back();</script>"
	 Response.end
  End If
  If IsObject(FieldsXml) Then
	on error resume next
	Dim objNode,i,j,objAtr
	 Set objNode=FieldsXml.documentElement 
	 For i=0 to objNode.ChildNodes.length-1 
		set objAtr=objNode.ChildNodes.item(i)
		If lcase(objAtr.Attributes.item(0).Text)="intro" Then 
		 Conn.Execute("UPDATE KS_User Set " & objAtr.Attributes.item(1).Text & "='" & Intro & "' Where UserName='" & KSUser.UserName & "'")
		End If
	 Next
				
  End If
  Conn.Execute("Update KS_EnterPrise Set Intro='" & Intro &"' WHERE UserName='" & KSUser.UserName & "'")
  Dim EID:EID=Conn.Execute("Select top 1 ID From KS_Enterprise Where UserName='" & KSUser.UserName & "'")(0)
  Call KS.FileAssociation(1033,EID,Intro,1)
  Call KSUser.AddLog(KSUser.UserName,"�޸�����ҵ������!",200)
  Response.Write "<script>alert('��ҵ����޸ĳɹ�!');history.back();</script>"
 End Sub
 
 
  Sub Job()
  %>
   <table  cellspacing="1" cellpadding="3" width="98%" align="center" border="0">
			<form action="?Action=JobSave" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
               <tr class="title">
                   <td height="22" colspan="2" align="center"> �� ҵ �� Ƹ</td>
               </tr>
               <tr class="tdbg">
                  <td>
                    <%
					Response.Write "<textarea ID='Job' name='Job' style='display:none'>" & KS.HTMLCode(Conn.Execute("Select top 1 Job From ks_Enterprise where username='" & KSUser.UserName & "'")(0)) & "</textarea>"

					%>  
<script type="text/javascript" src="../editor/ckeditor.js" mce_src="../editor/ckeditor.js"></script>
					<script type="text/javascript">
                CKEDITOR.replace('Job', {width:"98%",height:"350px",toolbar:"Simple",filebrowserBrowseUrl :"../editor/ksplus/SelectUpFiles.asp",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
			    </script> 					 
					</td>
                          </tr>
						  <tr class="tdbg">
                            <td align="center"><input  class="button" name="Submit" type="submit"  value=" OK,�� �� " />
                              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input  class="button" name="Submit2" type="reset" value=" �� �� " />                            </td>
                          </tr>
				</form>
	</table>
  <%
  End Sub
 
 Sub JobSave()
  Dim Job
  Job= Request.Form("Job")
  Job=KS.CheckScript(KS.HtmlCode(Job))
  Job=KS.HtmlEncode(Job)
  IF Job="" Then
  	 Response.Write "<script>alert('�Բ�����û����Ƹ��Ϣ');history.back();</script>"
	 Response.end
  End If
  Conn.Execute("Update KS_EnterPrise Set Job='" & Job &"' WHERE UserName='" & KSUser.UserName & "'")
  Response.Write "<script>alert('��Ƹ��Ϣ�޸ĳɹ�!');history.back();</script>"
 End Sub
 
  
  Sub BasicInfoSave() 
	   Dim CompanyName:CompanyName=KS.LoseHtml(KS.S("CompanyName"))
	   Dim Province:Province=KS.S("Province")
	   Dim City:City=KS.S("City")
	   Dim Address:Address=KS.LoseHtml(KS.S("Address"))
	   Dim ZipCode:ZipCode=KS.LoseHtml(KS.S("ZipCode"))
	   Dim ContactMan:ContactMan=KS.LoseHtml(KS.S("ContactMan"))
	   Dim QQ:QQ=KS.S("QQ")
	   Dim Mobile:mobile=KS.S("Mobile")
	   Dim Email:Email=KS.S("Email")
	   Dim Telphone:TelPhone=KS.LoseHtml(KS.S("TelPhone"))
	   Dim Fax:Fax=KS.LoseHtml(KS.S("Fax"))
	   Dim WebUrl:WebUrl=KS.LoseHtml(KS.S("WebUrl"))
	   Dim Profession:Profession=KS.LoseHtml(KS.S("Profession"))
	   Dim CompanyScale:CompanyScale=KS.LoseHtml(KS.S("CompanyScale"))
	   Dim RegisteredCapital:RegisteredCapital=KS.LoseHtml(KS.S("RegisteredCapital"))
	   Dim LegalPeople:LegalPeople=KS.LoseHtml(KS.S("LegalPeople"))
	   Dim BankAccount:BankAccount=KS.LoseHtml(KS.S("BankAccount"))
	   Dim AccountNumber:AccountNumber=KS.LoseHtml(KS.S("AccountNumber"))
	   Dim BusinessLicense:BusinessLicense=KS.LoseHtml(KS.S("BusinessLicense"))
	   Dim ClassID:ClassID=KS.ChkClng(KS.G("ClassID"))
	   Dim SmallClassID:SmallClassID=KS.ChkClng(KS.G("SmallClassID"))
	   Dim MapMarker:MapMarker=KS.G("MapMark")
	   Dim NewReg:NewReg=false
		
	   If CompanyName="" Then Response.Write "<script>alert('��˾���Ʊ�������');history.back();</script>":response.end

            Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select top 1 * From KS_Enterprise Where UserName='" & KSUser.UserName & "'",Conn,1,3
			  IF RS.Eof And RS.Bof Then
				 RS.AddNew
				 RS("UserName")=KSUser.UserName
				 RS("AddDate")=Now
				 RS("Recommend")=0
				 If KS.SSetting(2)=1 then
				 RS("status")=0
				 Else
				 RS("status")=1
				 End If
				 Dim RSS:Set RSS=Server.CreateObject("ADODB.RECORDSET")
				 RSS.Open "select top 1 * from ks_blog where username='" & KSUser.UserName & "'",conn,1,3
				 if RSS.Eof Then
				      RSS.AddNew
					  RSS("UserName")=KSUser.UserName
					  RSS("ClassID") = KS.ChkClng(Conn.Execute("Select Top 1 ClassID From KS_BlogClass")(0))
					  RSS("Announce")="���޹���!"
					  RSS("ContentLen")=500
					  RSS("Recommend")=0
				 End If
					  if KS.SSetting(2)=1 then
					  RSS("Status")=0
					  else
					  RSS("Status")=1
					  end if
				  RSS("TemplateID")=KS.ChkClng(Conn.Execute("Select Top 1 ID From KS_BlogTemplate Where flag=4 and IsDefault='true'")(0))
     			  RSS("BlogName")=CompanyName
				  RSS.Update
				  RSS.Close
				  Set RSS=Nothing
				  NewReg=true
				 
			  End If
			     RS("CompanyName")=CompanyName
				 RS("Province")=Province
				 RS("City")=City
				 RS("Address")=Address
				 RS("ZipCode")=ZipCode
				 RS("ContactMan")=ContactMan
				 RS("QQ")=QQ
				 RS("Mobile")=Mobile
				 RS("Email")=Email
				 RS("Telphone")=Telphone
				 RS("Fax")=Fax
				 RS("WebUrl")=WebUrl
				 RS("Profession")=Profession
				 RS("CompanyScale")=CompanyScale
				 RS("RegisteredCapital")=RegisteredCapital
				 RS("LegalPeople")=LegalPeople
				 RS("BankAccount")=BankAccount
				 RS("AccountNumber")=AccountNumber
				 RS("BusinessLicense")=BusinessLicense
				 RS("ClassID")=ClassID
				 RS("SmallClassID")=SmallClassID
				 RS("MapMarker")=MapMarker
				 'RS("Intro")=KS.HtmlEncode(Request.Form("Intro"))
		 		 RS.Update
				 Conn.Execute("Update KS_User Set UserType=1 where UserName='" & KSUser.UserName & "'")
				 If KS.C_S(8,21)="1" Then
				 Conn.Execute("Update KS_GQ Set ContactMan='" & ContactMan &"',Tel='" & Telphone & "',CompanyName='" & CompanyName & "',Address='" & Address & "',Province='" & Province & "',City='" & City & "',Zip='" & ZipCode & "',Fax='" & Fax & "',Homepage='" & WebUrl & "' where inputer='" & KSUser.UserName & "'")
				 End If
				 
				 
				 Set RSS=Conn.Execute("Select top 1 BlogName From KS_Blog Where UserName='" & KSUser.UserName & "'")
				 If Not RSS.Eof Then
				   If Instr(RSS(0),"���˿ռ�")<>0 Then
				    Conn.Execute("Update KS_Blog Set BlogName='" & CompanyName & "' where username='" & KSUser.UserName &"'")
				   End If
				 End If
				 RSS.Close
				 Set RSS=Nothing
				 
				 If IsObject(FieldsXml) Then
					 Dim objNode,i,j,objAtr
					 Set objNode=FieldsXml.documentElement 
					 If objNode.Attributes.item(0).Text="2" Then
						 For i=0 to objNode.ChildNodes.length-1 
								set objAtr=objNode.ChildNodes.item(i) 
								on error resume next
								If lcase(objAtr.Attributes.item(0).Text)<>"intro" Then 
								Conn.Execute("UPDATE KS_User Set " & objAtr.Attributes.item(1).Text & "='" & RS(objAtr.Attributes.item(0).Text) & "' Where UserName='" & KSUser.UserName & "'")
								End If
						 Next
					 End If
			
				   End If
				 
				 RS.Close:Set RS=Nothing
				 Call KSUser.AddLog(KSUser.UserName,"�޸�����ҵ������Ϣ����!",200)
				 If KS.S("ComeUrl")<>"" then
				 Response.Write "<script>alert('��ҵ������Ϣ�����޸ĳɹ���');location.href='" & KS.S("ComeUrl") & "';</script>"
				 Else
				  if NewReg=true Then
				 Response.Write "<script>alert('��ҵ������Ϣ�����޸ĳɹ�,��ȷ����д��ҵ���ܣ�');top.location.href='user_Enterprise.asp?action=intro';</script>"
				  Else
				 Response.Write "<script>alert('��ҵ������Ϣ�����޸ĳɹ���');location.href='user_Enterprise.asp';</script>"
				  End If
				End If
  End Sub
 

End Class
%> 
