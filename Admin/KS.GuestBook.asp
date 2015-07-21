<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Guest_Manage
KSCls.Kesion()
Set KSCls = Nothing

Class Guest_Manage
        Private KS,Action,KSCls
	    Private MaxPerPage, TotalPut , CurrPage, TotalPage, i, j, Loopno
	    Private KeyWord, SearchType,SqlStr,RS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
	Public Sub Kesion()
	KeyWord = KS.R(Trim(Request("keyword")))
	SearchType = KS.R(Trim(Request("SearchType")))
	Action=KS.G("Action")
	Select Case Action
	 Case "Main"  Call GuestMain()
	 Case "Del"  Call GuestDel()
	 Case "Reply" Call Reply()
	 Case "Revert" Call Revert()
	 Case "DelRecycle" Call DelRecycle()
	 Case "DelRecycleAll" Call DelRecycleAll()
	 Case Else  Call GuestMain()
	 End Select
	End Sub
	Sub GuestMain()
			If Not KS.ReturnPowerResult(0, "KSMS20004") Then
			  Call KS.ReturnErr(1, "")
			  exit sub
			End If

%>
<html>
<head>
<title>��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Include/admin_Style.css" type="text/css">
<script src="../ks_inc/jquery.js"></script>
<script language="JavaScript">
<!--
function CheckSelect()
{
	var count=0;
	for(i=0;i<document.KS_GuestBook.elements.length;i++)
	{
		if(document.KS_GuestBook.elements[i].name=="GuestID")
		{		
			if(document.KS_GuestBook.elements[i].checked==true)
			{
				count++;					
			}				
		}			
	}
		
	if(count<=0)
	{
		alert("��ѡ��һ��Ҫ��������Ϣ��");
		return false;
	}

	return true;
}

function cdel()
{
	if(CheckSelect()==false)
	{
		return false;
	}
	
	if (confirm("�����Ҫɾ���������Լ�¼�𣿲��ɻָ���")){
		document.KS_GuestBook.Flag.value = "del";
		document.KS_GuestBook.submit();
	}
}

function ccheck()
{
	if(CheckSelect()==false)
	{
		return false;
	}
	
	if (confirm("��ȷ��Ҫ�����Щ��Ϣ��")){
		document.KS_GuestBook.Flag.value = "check";
		document.KS_GuestBook.submit();
	}
}

function cuncheck()
{
	if(CheckSelect()==false)
	{
		return false;
	}
	
	if (confirm("��ȷ��Ҫ������Щ��Ϣ������߽���������Щ��Ϣ��")){
		document.KS_GuestBook.Flag.value = "uncheck";
		document.KS_GuestBook.submit();
	}
}

function SelectCheckBox()
{
	for(i=0;i<document.KS_GuestBook.elements.length;i++)
	{
		if(document.all("selectCheck").checked == true)
		{
			document.KS_GuestBook.elements[i].checked = true;					
		}
		else
		{
			document.KS_GuestBook.elements[i].checked = false;
		}
	}
}
//-->
</script>

<div class='topdashed sort' style="text-align:left;padding-left:10px"> <a href="KS.GuestBook.asp">���ӹ���</a>  <a href="KS.GuestBook.asp?Action=Recycle">����վ</a></div>
<%if request("action")="Recycle" Then
    Call Recycle() : Exit Sub
  end If
%>

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="TableBar">
      <form action="KS.GuestBook.asp" method="post" name="search" id="search">
        <tr>
          <td height="25">�������� --&gt;&gt;&gt; �ؼ��ʣ�
            <input type="text" name="keyword" class="inputtext" size="35" value="<%=KeyWord%>" onMouseOver="this.focus()" onFocus="this.select()">
                <select name="SearchType" size="1" class="inputlist">
                  <option value="content" <%If SearchType = "content" Then Response.Write "selected"%>>��������</option>
                  <option value="author" <%If SearchType = "author" Then Response.Write "selected"%>>�� �� ��</option>
                </select>
                <input type="submit" name="imageField" value="����"></td>
        </tr>
      </form>
    </table>
<table border="0" width="100%" align="center" style='border-top:1px solid #cccccc' cellpadding="0" cellspacing="0">
	<form name="KS_GuestBook" action="KS.GuestBook.asp?Action=Del" method=post>
	<input name="Flag" type="hidden" value="" id="Flag">
		<tr class="sort">
					<td>&nbsp;</td>
					<td>����</td>
					<td>������</td>
					<td>�ظ�/�鿴</td>
					<td>��󷢱�</td>
					<td>״̬</td>
					<td>�������</td>
		</tr>
	<%
	Dim Param:Param=" Deltf=0"
	If Not KS.IsNul( KeyWord) Then
		If SearchType = "content" Then
			Param=param & " and Subject LIKE '%"& KeyWord &"%'"  
		Else
			Param=param & " and UserName LIKE '%"& KeyWord &"%'" 
		End If
	ENd If
	MaxPerPage=20
	CurrPage = KS.ChkClng(Request("Page")) : If CurrPage<=0 Then CurrPage=1
	SQLStr=KS.GetPageSQL("KS_GuestBook","id",MaxPerPage,CurrPage,1,Param,"*")
	Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open SqlStr,Conn,1,1 
	If RS.Eof or RS.Bof Then 
		Response.Write "<tr><td colspan='10' align='center' height='30'><font color=#FF0000>��ʱ��û���κμ�¼��</font></td></tr>"
	Else
	    If Param<>"" Then Param=" Where " & Param
		totalPut = Conn.Execute("Select count(id) from [KS_GuestBook] " & Param)(0)

		i = 0
		Do While Not RS.Eof 
%>
        <tr>
          <td  height="30" class='splittd' align="center" valign="middle"><input type="checkbox" name="GuestID" value="<%=Trim(RS("ID"))%>"></td>
		 <td class='splittd'><img src="../club/images/common.gif" align="absmiddle">
		  
		 <% on error resume next
		   response.write "[<a href='" & KS.GetClubListUrl(rs("boardid")) & "' target='_blank'>" & conn.execute("select boardname from ks_guestboard where id=" & rs("boardid"))(0) & "</a>]"
		 if KS.Setting(59)="1" Then
		  response.write "<a href='?action=Reply&guestid=" & rs("id") & "'>"
		  else
		  %>
		 <a href="<%=KS.GetClubShowUrl(rs("id"))%>" target="_blank">
		 <%end if%><%=rs("subject")%></a>
		 <%if not ks.isnul(rs("annexext")) then%>
		 <img src="../editor/ksplus/fileicon/<%=rs("annexext")%>.gif" alt="<%=rs("annexext")%>����" align="absmiddle">
		 <%end if%>
		 <%if rs("ispic")="1" then%>
		 <img src="../editor/ksplus/fileicon/gif.gif" alt="gifͼƬ����" align="absmiddle">
		 <%elseif rs("ispic")="2" then%>
		 <img src="../editor/ksplus/fileicon/jpg.gif" alt="jpgͼƬ����" align="absmiddle">
		 <%end if%>
		 <%if rs("isslide")="1" then%>
		  <font color=red>��</font>
		 <%end if%>
		 </td>
		 <td class='splittd'>
		 <%
		 if ks.isnul(rs("username")) then 
		  response.write "�ο�"
		 else
		  response.write rs("username")
		 end if
		 %>
		 </td>
		 <td class='splittd' align="center">
		 <%
		 if KS.Setting(59)="1" Then
			  if conn.execute("select top 1 id from " & rs("posttable") &" where parentid<>0 and topicid=" & rs("id")).eof then
			   response.write "<font color=red>δ�ظ�</font>"
			  else
			   response.write "<font color=green>�ѻظ�</font>"
			  end if
		 else
		  response.write RS("TotalReplay") & "/" & rs("hits")
		 end if
		 %>
		 </td>
		 <td class='splittd'>
		 <%
		 if ks.isnul(RS("LastReplayUser")) then 
		  response.write "�ο�"
		 else
		  response.write RS("LastReplayUser")
		 end if
		 %>
		 </td>
		 <td class='splittd' align='center'>
		 <%
		  If rs("verific")=1 then
		   response.write "<a href='?Action=Del&flag=uncheck&guestid=" & rs("id") & "'><font color=blue>����</font></a>"
		  else
		   response.write "<a href='?Action=Del&flag=check&guestid=" & rs("id") & "'><font color=red>δ��</font></a>"
		  end if
		 %>
		 </td>

		 <td class='splittd' align="center">
		 <%
		  If rs("isslide")="1" then
		   response.write "<a href='?Action=Del&flag=unslide&guestid=" & rs("id") & "'><font color=red>ȡ���õ�</font></a>"
		  else
		   if rs("ispic")<>"0" then
		   response.write "<a href='?Action=Del&flag=slide&guestid=" & rs("id") & "'>���ûõ�</a>"
		   end if
		  end if
		 %>
		 
		  <%
		 if KS.Setting(59)="1" Then
		   response.write "<a href='?action=Reply&guestid="& rs("id") & "'>�ظ�/�޸�</a>  | "
		 end if
		   %>
		 

		 <a href="?Action=Del&flag=del&guestid=<%=rs("id")%>" onClick="return(confirm('���и������µĻظ�Ҳ����ɾ����ȷ����'))">ɾ��</a> | <a href="<%=KS.GetClubShowUrl(rs("id"))%>" target="_blank">�鿴</a> 
		 
		 </td>
		</tr>
        <%
		i=i+1
		if i>=maxperpage then exit do
	RS.MoveNext
	Loop
	%>
</form>
	</table>
	<%
End if
RS.Close
Set RS=Nothing
%>
        <table border="0" width="100%" cellspacing="0" cellpadding="2"  align="center" >
          <tr>
		    <td ><label><input type="checkbox"  name='selectCheck' onClick="SelectCheckBox()">ȫ��ѡ��</label>
              <input name="delbtn" value="ɾ��"  class="button" type="button" onClick="cdel();">
			  <input name="delbtn" value="���" class="button" type="button" onClick="ccheck();">
	          <input name="delbtn" value="ȡ�����" class="button" type="button" onClick="cuncheck();">
			</td>

          </tr>
      </table>
 <%
 Call KS.ShowPage(totalput, MaxPerPage, "", CurrPage,true,true)
%>
<br style="clear:both">

<div class="attention">
<strong>�ر����ѣ�</strong>
ֻ���ϴ�ͼƬ���������Ӳſ������ûõ�����,����ֻ����jpg��ʽ����������Ϊ�õ�,������ܵ��ò�������
</div>
<br>
<br>
<%
 End Sub
 
 Sub Recycle()
    Dim Table:Table=KS.G("Table")
    If KS.IsNul(Table) Then Table="KS_GuestBook"
   %>
   <strong>ѡ�����ݱ�</strong><select name="table" onChange="location.href='?action=Recycle&table='+this.value">
   <option value="KS_GuestBook">�����(KS_GuestBook ��<%=conn.execute("select count(1) from KS_GuestBook where deltf=1")(0)%>��)</option>
   <%
 
    Dim Node,TableXML:set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	TableXML.async = false
	TableXML.setProperty "ServerHTTPRequest", true 
	TableXML.load(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))
	Dim Url,N:N=0
    For Each Node In TableXML.DocumentElement.SelectNodes("item")
	  If KS.S("Table")=Node.SelectSingleNode("tablename").text Then
	  Response.Write "<option value='" & Node.SelectSingleNode("tablename").text &"' selected>�ظ���(" & Node.SelectSingleNode("tablename").text&" ��" & conn.execute("select count(1) from " &Node.SelectSingleNode("tablename").text &" where deltf=1")(0) &"��)</option>"
	  Else
	  Response.Write "<option value='" & Node.SelectSingleNode("tablename").text &"'>�ظ���(" & Node.SelectSingleNode("tablename").text&" ��" & conn.execute("select count(1) from " &Node.SelectSingleNode("tablename").text &" where deltf=1")(0) &"��)</option>"
	  End If
	Next
	
	Dim param:Param=" DelTF=1"
	MaxPerPage=20
	CurrPage = KS.ChkClng(Request("Page")) : If CurrPage<=0 Then CurrPage=1
	SQLStr=KS.GetPageSQL(Table,"id",MaxPerPage,CurrPage,1,Param,"*")
	If Param<>"" Then Param=" Where " & Param
	totalPut = Conn.Execute("Select count(id) from [" & Table & "] " & Param)(0)
 %>
   </select>
   
   ��ǰ���ڹ�������ݱ�<font color=blue><%=Table%></font>,���� <font color=red><%=totalput%></font> ��
 	<form name="KS_GuestBook" action="KS.GuestBook.asp" method="post">

 <table border="0" width="100%" cellspacing="0" cellpadding="2"  align="center" >
          <tr>
		    <td ><label><input type="checkbox"  name='selectCheck' onClick="SelectCheckBox()">ȫ��ѡ��</label>
              <input name="delbtn" value="����ɾ��"  class="button" type="submit" onClick="if (confirm('�˲��������棬ȷ������ɾ��ѡ�еļ�¼��')){$('#action').val('DelRecycle');}else{return false;}">
              <input name="delbtn" value="һ�����"  class="button" type="submit" onClick="if (confirm('�˲��������棬ȷ������һ����ռ�¼��')){$('#action').val('DelRecycleAll');}else{return false;}">
	          <input name="delbtn" value="������ԭ" class="button" type="submit" onClick="$('#action').val('Revert');">
			</td>

          </tr>
      </table>
 <table border="0" width="100%" align="center" style='border-top:1px solid #cccccc' cellpadding="0" cellspacing="0">
	<input type="hidden" name="action" id="action" value=""/>
	<input type="hidden" name="table" id="table" value="<%=table%>"/>
		<tr class="sort">
					<td>&nbsp;</td>
					<%if lcase(table)<>"ks_guestbook" Then%>
					<td>�ظ�����</td>
					<td>����</td>
					<td>����ʱ��</td>
					<%else%>
					<td>����</td>
					<td>����</td>
					<td>����</td>
					<td>��󷢱�</td>
				    <%end if%>
					<td>�������</td>
		</tr>
		<%
        on error resume next
		Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		RS.Open SqlStr,conn,1,1
		If RS.Eof or RS.Bof Then 
		Response.Write "<tr><td colspan='10' align='center' height='30'><font color=#FF0000>����վ��û�м�¼��</font></td></tr>"
	    Else
			i = 0
			Do While Not RS.Eof 
			%>
			<tr>
             <td  height="30" class='splittd' align="center" valign="middle"><input type="checkbox" name="ID" value="<%=Trim(RS("ID"))%>"></td>
			 <%if lcase(table)<>"ks_guestbook" Then%>
		     <td class='splittd'><img src="../club/images/common.gif" align="absmiddle">
		      <a href="<%=KS.GetClubShowUrl(rs("topicid"))%>" target="_blank"><%=ks.gottopic(rs("content"),80)%></a>
			 </td>
		     <td class='splittd'><%=rs("username")%></td>
		     <td class='splittd'><%=formatdatetime(rs("ReplayTime"),2)%></td>
			 <%else%>
		     <td class='splittd'><img src="../club/images/common.gif" align="absmiddle">
		      <a href="<%=KS.GetClubShowUrl(rs("id"))%>" target="_blank"><%=KS.Gottopic(rs("subject"),38)%></a> (����<font color=red> <%=rs("TotalReplay")%></font> ��)
			 </td>
			 <td class="splittd"><%response.write "<a href='" & KS.GetClubListUrl(rs("boardid")) & "' target='_blank'>" & conn.execute("select top 1 boardname from ks_guestboard where id=" & rs("boardid"))(0) & "</a>"
             %></td>
			 <td class="splittd"><a href="<%=KS.GetSpaceUrl(rs("userid"))%>" target="_blank"><%=rs("username")%></a></td>
			 <td class="splittd" style="text-align:center"><%=Formatdatetime(rs("LastReplayTime"),2)%></td>
			 <%end if%>
			 <td class="splittd" nowrap style="text-align:center"><a href="?table=<%=table%>&action=Revert&id=<%=rs("id")%>">��ԭ</a> <a href="?action=DelRecycle&table=<%=table%>&id=<%=rs("id")%>" onClick="return(confirm('�˲��������棬ȷ��ִ��ɾ����'));">ɾ��</a></td>
		    </tr>
			<%
			RS.MoveNext
			Loop
	    End If
		RS.Close:Set RS=nothing
		%>
  </table>
<table border="0" width="100%" cellspacing="0" cellpadding="2"  align="center" >
          <tr>
		    <td >
              <input name="delbtn" value="����ɾ��"  class="button" type="submit" onClick="if (confirm('�˲��������棬ȷ������ɾ��ѡ�еļ�¼��')){$('#action').val('DelRecycle');}else{return false;}">
              <input name="delbtn" value="һ�����"  class="button" type="submit" onClick="if (confirm('�˲��������棬ȷ������һ����ռ�¼��')){$('#action').val('DelRecycleAll');}else{return false;}">
	          <input name="delbtn" value="������ԭ" class="button" type="submit" onClick="$('#action').val('Revert');">
			</td>

          </tr>
      </table>
	 </form>
  <%
 Call KS.ShowPage(totalput, MaxPerPage, "", CurrPage,true,true)
%>
<div style="clear:both"></div>
<div class="attention">
<strong>�ر����ѣ�</strong>
����ɾ���󣬽����ָܻ������ز�����
</div>
<br>
<br>

  <%
 End Sub
 
 '��ԭ
 Sub Revert()
  Dim ID:ID=KS.FilterIds(KS.S("ID"))
  Dim Table:Table=KS.G("Table")
  If KS.IsNul(ID) Or Table="" Then KS.AlertHintScript "û��ѡ��Ҫ��ԭ�ļ�¼!"
  if Lcase(table)<>"ks_guestbook" Then
    Dim RS:Set RS=Conn.Execute("Select TopicID From " & Table &" Where id In ( "& ID & ")")
	Do While Not RS.Eof
	  Conn.Execute("Update KS_GuestBook Set TotalReplay=TotalReplay+1 Where id=" & rs(0))
	 RS.MoveNext
	Loop
	RS.Close
	Set RS=Nothing
  End If
  Conn.Execute("Update " & Table & " Set DelTF=0 Where ID In(" & ID &")")
  KS.AlertHintScript "��ϲ����ԭ�ɹ�!"
 End Sub
 
 'һ�����
 Sub DelRecycleAll()
 Dim RS,Table:Table=KS.G("Table")
  if Lcase(table)<>"ks_guestbook" Then  'ɾ���ظ�
	   Set RS=Server.CreateObject("ADODB.RECORDSET")
	   RS.Open "Select ID,TopicID From " & Table & " Where DelTF=1",conn,1,1
	   Do While Not RS.Eof 
		 Conn.Execute("Delete From KS_GuestComment Where Tid=" & rs(1) & " and pid=" & rs(0))
	   RS.MoveNext
	   Loop
	   RS.CLOSE:Set RS=Nothing
    Conn.Execute("Delete From " &Table & " Where DelTF=1")
	KS.AlertHintScript "��ϲ��һ��������ݱ�" & Table & "����վ�����ݳɹ�!"
  Else
	  Dim TopicIds
	  Set RS=Conn.Execute("Select Id From KS_GuestBook Where DelTF=1")
	  Do While Not RS.Eof 
		   If TopicIDs="" Then
			 TopicIDs=RS(0)
			Else
			TopicIDs=TopicIDs & "," & RS(0)
			End If
		  RS.MoveNext
		  Loop
	   RS.Close : Set RS=Nothing
	   If TopicIds<>"" Then
		Call DoDelete(TopicIds)
	   Else
		KS.AlertHintScript "���ݱ�" & Table & "����վ��û�м�¼!"
	   End If
  End If
 End Sub
 
 '����ɾ��
 Sub DelRecycle()
  Dim TopicIds:TopicIds=KS.FilterIds(KS.S("ID"))
  Dim Table:Table=KS.G("Table")
  If KS.IsNul(TopicIds) Or Table="" Then KS.AlertHintScript "û��ѡ��Ҫɾ���ļ�¼!"
  if Lcase(table)<>"ks_guestbook" Then  'ɾ���ظ�
   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "Select ID,TopicID From " & Table & " Where ID in("& TopicIds&")",conn,1,1
   Do While Not RS.Eof 
     Conn.Execute("Delete From KS_GuestComment Where Tid=" & rs(1) & " and pid=" & rs(0))
   RS.MoveNext
   Loop
   RS.CLOSE:Set RS=Nothing
   Conn.Execute("Delete From " &Table & " Where ID in("& TopicIds&")")
	KS.AlertHintScript "��ϲ��������ݱ�" & Table & "����վ��ѡ�е����ݳɹ�!"
  Else
   Call DoDelete(TopicIds)
  End If
 End Sub
 Sub doDelete(TopicIds)
  Dim TodayNum:TodayNum=0
  dim boardid,postTable,userName,id,BSetting
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select UserName,boardid,subject,AddTime,PostTable,ID From KS_GuestBook Where ID in(" & TopicIds &")",conn,1,1
			If Not RS.Eof Then
			 Do While Not RS.Eof
				  id=RS("ID")
				  boardid=rs(1)
				  postTable=rs(4)
				  userName=rs(0)
				  If DateDiff("d",rs(3),Now)=0 Then
				   TodayNum=TodayNum+1
				  End If
				  If boardid<>0 then 
					 KS.LoadClubBoard()
					 On Error Resume Next
					 Dim Node:Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
					 Dim LastPost,LastPostArr:LastPostArr=Split(Node.SelectSingleNode("@lastpost").text,"$")
					 
					 '������ҳ����������
					 If KS.ChkClng(LastPostArr(0))=ID Then
					   Dim RSNew:Set RSNew=Conn.Execute("Select top 1 ID,BoardID,Subject,AddTime From KS_GuestBook Where BoardID=" & boardid & " and verific=1 and id<>" & id & " order by id desc")
					   If Not RSNew.Eof Then
						 LastPost=RSNew(0) & "$" & RSNew(3) & "$" & Replace(left(RSNew(2),200),"$","") & "$$$$$$$$"
					   Else
						 LastPost="��$��$��$$$$$$$$"
					   End If
					   Conn.Execute("Update KS_GuestBoard Set LastPost='" & LastPost & "' Where ID=" & BoardID)
					   Node.SelectSingleNode("@lastpost").text=LastPost
					 End If
				  end if
				  
				  if not KS.ISNul(rs(0)) then
				     On Error Resume Next
					 BSetting=Node.SelectSingleNode("@settings").text
					 If Not KS.IsNul(BSetting) Then
						 If KS.ChkClng(Split(BSetting,"$")(34))<>0 Then
						  Conn.Execute("Update KS_User Set Prestige=Prestige-" & KS.ChkClng(Split(BSetting,"$")(34)) & " Where UserName='" & rs(0) &"' and Prestige>0")
						 End If
					 
					   If KS.ChkClng(Split(BSetting,"$")(7))>0 Then
						Call KS.ScoreInOrOut(rs(0),2,KS.ChkClng(Split(BSetting,"$")(7)),"ϵͳ","����̳�����������[" & rs(2) & "]��ɾ��!",0,0)
					   End If
					 End If
				  end if
				  
				  Dim Num,replyNum:replyNum=Conn.Execute("Select count(id) from " & PostTable & " where topicid=" & id)(0)
				  TodayNum=TodayNum+Conn.Execute("Select count(id) from " & PostTable & " where topicid=" & id &" and datediff(" & DataPart_D & ",ReplayTime," & SqlNowString&")=0")(0)
				  Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				  Doc.async = false
				  Doc.setProperty "ServerHTTPRequest", true 
				  Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
				  Dim XMLDate:XMLDate=doc.documentElement.attributes.getNamedItem("date").text
				  Num=KS.ChkClng(doc.documentElement.attributes.getNamedItem("todaynum").text)-TodayNum
				  If Num<0 Then Num=0
				  doc.documentElement.attributes.getNamedItem("todaynum").text=Num
				  Num=KS.ChkClng(doc.documentElement.attributes.getNamedItem("postnum").text)-replyNum
				  If Num<0 Then Num=0
				  doc.documentElement.attributes.getNamedItem("postnum").text=Num
				  Num=KS.ChkClng(doc.documentElement.attributes.getNamedItem("topicnum").text)-1
				  If Num<0 Then Num=0
				  doc.documentElement.attributes.getNamedItem("topicnum").text=Num
				  
				  Conn.Execute("Update KS_GuestBoard Set TodayNum=TodayNum-" & TodayNum & " where id=" &boardid &" and todaynum>=" & TodayNum)
				  Conn.Execute("Update KS_GuestBoard Set PostNum=PostNum-" & replyNum -1& " where id=" &boardid &" and PostNum>=" & replyNum-1)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.selectSingleNode("row[@id=" & boardid & "]/@postnum").text=Conn.Execute("Select PostNum From KS_GuestBoard Where id=" & boardid)(0)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.selectSingleNode("row[@id=" & boardid & "]/@todaynum").text=Conn.Execute("Select TodayNum From KS_GuestBoard Where id=" & boardid)(0)
		
				  doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
					
					Conn.Execute("update KS_User set postNum=postNum-1 where userName='" & UserName & "' and postNum>0")
					Conn.Execute("delete from KS_Guestbook where id=" & ID)
					Conn.Execute("Delete From KS_GuestComment Where tid=" & ID)
					Conn.Execute("delete from " & PostTable & " where TopicID=" & ID)
					Conn.Execute("delete from KS_UploadFiles where ID=" & ID & " and channelid=9994")
			  RS.MoveNext
			Loop 
			End If
			rs.close:set rs=nothing
			
    KS.AlertHintScript "��ϲ��ɾ���ɹ�!"

 End Sub
 
 
 'ɾ������
 Sub GuestDel()
			Dim strIdList,arrIdList,iId,i,Flag,SqlStr
			strIdList = Trim(KS.G("GuestID"))
			Flag = Trim(KS.G("Flag"))
			Select Case Flag
			Case "del"
				If Not IsEmpty(strIdList) Then
					arrIdList = Split(strIdList,",")
					For i = 0 To UBound(arrIdList)
						iId = Clng(arrIdList(i))			
						'dim PostTable,rst :set rst=conn.execute("select top 1 PostTable From KS_GuestBook Where ID=" & iId)
					   'If RSt.Eof Then
						'  RSt.Close :Set RSt=Nothing
						'  KS.Die "error"
					   'End If
					   'PostTable=RSt(0)
					   'RSt.Close : Set RSt=Nothing
					   
						SqlStr = "Update KS_GuestBook Set DelTF=1 WHERE ID=" & iId
						Conn.Execute SqlStr	
                        'on error resume next
						'Conn.Execute("Delete FROM " & PostTable & " Where TopicID=" & iId)		
						'Conn.Execute("Delete FROM KS_UploadFiles Where ChannelID=9994 and infoID=" & iId)		
						'if err then err.clear
					Next
					Call KS.Alert("��Ϣɾ���ɹ���ȷ�Ϸ��أ�",Request.ServerVariables("HTTP_REFERER"))
				Else
					Call KS.AlertHistory("������ѡ��һ����Ϣ��¼��",-1)
				End If
			Case "check"
				If Not IsEmpty(KS.FilterIds(strIdList)) Then
				    Dim RS
					Set RS=Conn.Execute("Select * From KS_GuestBook Where ID in(" & KS.FilterIds(strIdList) & ")")
					Do While Not RS.Eof
						Conn.Execute("update " & RS("PostTable") &" set verific=1 where TopicID=" & RS("ID"))
					RS.MoveNext
					Loop
					RS.Close :Set RS=Nothing
					Conn.Execute("UPDATE KS_GuestBook SET Verific = 1 WHERE ID in(" & KS.FilterIds(strIdList) & ")")
					Call KS.Alert("��Ϣ��˳ɹ���ȷ�Ϸ��أ�",Request.ServerVariables("HTTP_REFERER"))
				Else
					Call KS.AlertHistory("������ѡ��һ����Ϣ��¼��",-1)
				End If
			Case "uncheck"
					If Not IsEmpty(KS.FilterIds(strIdList)) Then
						Set RS=Conn.Execute("Select * From KS_GuestBook Where ID in(" & KS.FilterIds(strIdList) & ")")
						Do While Not RS.Eof
							Conn.Execute("update " & RS("PostTable") &" set verific=0 where TopicID=" & RS("ID"))
						RS.MoveNext
						Loop
						RS.Close :Set RS=Nothing
						Conn.Execute("UPDATE KS_GuestBook SET Verific = 0 WHERE ID in(" & KS.FilterIds(strIdList) & ")")
						Call KS.Alert("��Ϣȡ����˳ɹ���ȷ�Ϸ��أ�",Request.ServerVariables("HTTP_REFERER"))
					Else
						Call KS.AlertHistory("������ѡ��һ����Ϣ��¼��",-1)
					End If
		  case "slide"
				If Not IsEmpty(strIdList) Then
					arrIdList = Split(strIdList,",")
					For i = 0 To UBound(arrIdList)
						iId = Clng(arrIdList(i))			
						Conn.Execute("UPDATE KS_GuestBook SET isslide = 1 WHERE ID="&iId&"")			
					Next
					Call KS.Alert("���ûõ����Գɹ���ȷ�Ϸ��أ�",Request.ServerVariables("HTTP_REFERER"))
				Else
					Call KS.AlertHistory("������ѡ��һ����Ϣ��¼��",-1)
				End If
		  case "unslide"
				If Not IsEmpty(strIdList) Then
					arrIdList = Split(strIdList,",")
					For i = 0 To UBound(arrIdList)
						iId = Clng(arrIdList(i))			
						Conn.Execute("UPDATE KS_GuestBook SET isslide = 0 WHERE ID="&iId&"")			
					Next
					Call KS.Alert("ȡ���õ����Գɹ���ȷ�Ϸ��أ�",Request.ServerVariables("HTTP_REFERER"))
				Else
					Call KS.AlertHistory("������ѡ��һ����Ϣ��¼��",-1)
				End If
		End Select
	End Sub
	
	Sub Reply()
	Dim Flag, pagetxt, guestid, ssubject, sanser, sadminhead, scheckbox, sansertime,SqlStr,RSObj,postTable
			Dim DomainStr:DomainStr= KS.GetDomain
			Flag =KS.G("Flag")
			pagetxt = Request("cpage")
			guestid = KS.ChkClng(Request("guestid"))
			if Flag="ok" then
			   ssubject =KS.G("txtcontop")   
			   sadminhead = KS.G("adminhead")
			   scheckbox = KS.G("htmlok")
			   sansertime = Now()
			   set rsobj=server.createobject("adodb.recordset")
			   rsobj.open "select top 1 postTable from ks_guestbook where id=" & guestid,conn,1,1
			   if rsobj.eof and rsobj.bof then
			    response.write "error!"
				response.End()
			   end if
			    postTable=rsobj("postTable")
				rsobj.close
				rsobj.open "select top 1 content from " & postTable & " where parentid=0 and topicid=" & guestid,conn,1,3
				if rsobj.eof and rsobj.bof then
				rsobj.addnew
				end if
				rsobj(0)=request.Form("content")
				rsobj.update
                rsobj.close
				
			   rsobj.open "select top 1 * from " & postTable &" where parentid<>0 and topicid=" & guestid,conn,1,3
			   if rsobj.eof and rsobj.bof then
			    rsobj.addnew
			   end if
			    rsobj("username")=KS.C("AdminName")
				rsobj("userip")=KS.GetIP()
				rsobj("TopicID")=guestid
				rsobj("parentid")=guestid
				rsobj("content")=request.Form("txtanser")
				rsobj("ReplayTime")=now()
				rsobj("txthead")=sadminhead
				rsobj("Verific")=1
				rsobj.update
			    rsobj.close:set rsobj=nothing
			   Response.write "<script>alert('��ϲ�����Իظ��ɹ���');location.href='KS.Guestbook.asp?page=" &pagetxt& "';</script>"
			End If
                Set RSObj=Server.CreateObject("Adodb.Recordset")
				SqlStr="SELECT top 1 * FROM KS_GuestBook WHERE ID="&guestid
				RSObj.Open SqlStr,Conn,1,1
			%>
			<html>
			<head>
			<title>�������</title>
			<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
			<link rel="stylesheet" href="Include/admin_Style.css" type="text/css">
			<br>
			<table width="540" border="0" cellspacing="0" cellpadding="0" align="center">
			  <form method="POST" action="KS.GuestBook.asp?Action=Reply&guestid=<%Response.Write guestid%>&amp;cpage=<%Response.Write pagetxt%>" name="repleBook">
				<tr>
				  <td valign="top"> <br>
					  <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
						<tr> 
						  <td colspan="2" align="center" height="14">:::::::::::::::::::::::::::::::::::: �� �� �� �� ::::::::::::::::::::::::::::::::::::</td>
						</tr>
						<tr> 
						  <td width="18%" align="center" height="32"><img src="<%=DomainStr%>Images/face/<%=RSObj("Face")%>"><br><%=RSObj("UserName")%></td>
						  <td>
						  <%
						  dim content,rs:set rs=server.createobject("adodb.recordset")
						  rs.open "select top 1 Content,txthead from " & rsobj("postTable") &" where parentid=0 and TopicID=" & guestid,conn,1,1
						  if not rs.eof then
						    content=rs(0)
						  else
						    content=" "
						  end if
						  rs.close
						  %>
						  <textarea  id="content" name="content"  style="display:none"><%=Server.HTMLEncode(content)%></textarea>
		   <script type="text/javascript" src="../editor/ckeditor.js" mce_src="../editor/ckeditor.js"></script>
		   <script type="text/javascript">
                CKEDITOR.replace('content', {width:"620",height:"150px",toolbar:"Basic",filebrowserBrowseUrl :"Include/SelectPic.asp?from=ckeditor&Currpath=<%=KS.GetUpFilesDir()%>",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
			</script>	
						  </td>
						</tr>
					  </table>
					<table width="100%" border="0" cellspacing="0" cellpadding="0" height="150" class="font" align="center">
					  <tr> 
						<td > </td>
					  </tr>
					  <tr> 
						<td nowrap align="center">:::::::::::::::::::::::::::::::::::: վ �� �� �� ::::::::::::::::::::::::::::::::::::</td>
					  </tr>
					  <tr> 
						<td nowrap align="center"  height="135" valign="middle" style="padding-left:80px"> 
						  <p> 
						  <%
						  dim replycontent,TxtHead
						  rs.open "select top 1 Content,txthead from " & rsobj("postTable") &" where parentid<>0 and TopicID=" & guestid,conn,1,1
						  if rs.eof then
						   replycontent=" "
						   TxtHead=1
						  else
						   replycontent=rs(0)
						   TxtHead=rs(1)
						  end if
						  rs.close:set rs=nothing%>
							<textarea rows="8" name="txtanser" cols="70" class="inputmultiline"><%=Server.HTMLEncode(replycontent)%></textarea>
							<script type="text/javascript">
                CKEDITOR.replace('txtanser', {width:"620",height:"150px",toolbar:"Basic",filebrowserBrowseUrl :"Include/SelectPic.asp?from=ckeditor&Currpath=<%=KS.GetUpFilesDir()%>",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
			</script>
						</td>
					  </tr>
					</table>
					  
					<div align="center">
					  <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
						<tr valign="bottom">
						  <td nowrap="nowrap" colspan="16" class="font"><div align="center">::::::::::::::::::::::::::::::::::::: ѡ �� 
							�� �� :::::::::::::::::::::::::::::::::::::</div></td>
						</tr>
						<tr height="25" align="center">
						  <td colspan="16"><%
						    Dim I,istr
							For I=1 To 24
							   if istr<9 then istr="0"&i else istr=i
							   Response.Write "<input type=""radio"" name=""Adminhead"" value=""" & istr & """"
							   IF I =TxtHead or i=1 Then Response.Write(" Checked")
							  Response.Write" ><img src=""../editor/ubb/images/smilies/default/" & istr & ".gif"" border=""0"">"
							  IF I Mod 12=0 Then Response.Write("<BR>")
							  
							 Next
					
					
%></td>
					    </tr>
					  </table>
					  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="font">
						<tr>
						  <td align="center"><font color="#400040">......................................................................................</font></td>
						</tr>
					  </table>
					  <table width="530" border="0" cellspacing="0" cellpadding="0" class="font">
						<tr>
						  <td height="35" align="center"> 
							  <input type="submit" value=" ȷ �� "  name="cmdOk" class="button">
							  &nbsp; 
							  <input type="reset" value=" �� �� " name="cmdReset" class="button">
							  &nbsp; 
							  <input type="button" value=" �� �� " name="cmdExit" class="button" onClick=" history.back()">
						  <input type="hidden" name="Flag" value="ok"></td>
						</tr>
					  </table>
					</div>
					</td>
				</tr>
			  </form>
			</table>
			<p>&nbsp;</p>
			<%
	End Sub
End Class
%>
 
