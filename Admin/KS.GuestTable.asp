<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<%
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim KS:Set KS=New PublicCls
If Not KS.ReturnPowerResult(0, "KSMS20004") Then
	Call KS.ReturnErr(1, "")
	response.end
End If
Dim TableXML,Node,N,TaskUrl,Taskid,Action
'Set TableXML=LFCls.GetXMLFromFile("task")
set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
TableXML.async = false
TableXML.setProperty "ServerHTTPRequest", true 
TableXML.load(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))

Action=Request.QueryString("Action")
    Manage


Sub manage()
%>
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<title>��̳���ݱ����</title>
<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>
<script language="JavaScript" src="../KS_Inc/Jquery.js"></script>
</head>
<body>
<ul id='mt'> <div id='mtl'>��̳���ݱ����</div></ul>
	  <table width="100%" align='center' border="0" cellpadding="0" cellspacing="0">
	  <form name="myform" action="KS.GuestTable.asp?action=ModifySave" method="post">
      <tr class="sort">
	    <td>���</td>
	    <td>������</td>
	    <td>����</td>
		<td>��ǰĬ��</td>
		<td>��¼��</td>
		<td>˵��</td>
		<td>�������</td>
	  </tr>
<%
  If TableXML.DocumentElement.SelectNodes("item").length=0 Then
      Response.Write "<tr class='list'><td colspan=7 height='25' class='splittd' align='center'>��û�����С��̳���ݱ�!</td></tr>"
  Else
	  N=0
	  For Each Node In TableXML.DocumentElement.SelectNodes("item")
	  %>
			  <tr  onmouseout="this.className='list'" onMouseOver="this.className='listmouseover'">               
			   <td class='splittd' height="30" align="center"><%=Node.SelectSingleNode("@id").text%></td>
			   <td class='splittd' height="30"><%=Node.SelectSingleNode("tablename").text%></td>
			   <td class='splittd' style='text-align:center' height="30"><%
			   if Node.SelectSingleNode("@issys").text="1" then
			    response.write "<span style='color:red'>ϵͳ</span>"
			   else
			    response.write "<span style='color:green'>�Զ���</span>"
			   end if
			   %></td>
			   <td class='splittd' align="center">
			   <%
				 if node.selectSingleNode("@isdefault").text="1" then
				  response.write "<input type='radio' name='isdefault' value='" & Node.SelectSingleNode("@id").text & "' checked>"
				 else
				  response.write "<input type='radio' name='isdefault' value='" & Node.SelectSingleNode("@id").text & "'>"
				 end if
				%>
			   </td>
			   <td class='splittd' align="center">
			   <%
			     dim num
				 num=conn.execute("select count(1) from " & Node.SelectSingleNode("tablename").text)(0)
				 response.write "<font color='#ff6600'>" & num & "</font>"
			   %>
			   </td>
			   <td class='splittd' align="center">
			   <%=Node.SelectSingleNode("descript").text%>
			   </td>
			   
			   <td class='splittd' align="center">
			    <%if node.selectSingleNode("@isdefault").text="1" or num>0 or Node.SelectSingleNode("@issys").text="1" then%>
				 <span style="color:#999999">ɾ��</span>
				<%else%>
				 <a href="?action=del&itemid=<%=Node.SelectSingleNode("@id").text%>" onClick="return(confirm('ȷ��ɾ����������?'))">ɾ��</a>
				<%end if%>
			   </td>
			  </tr>
	  <%
		n=n+1
	  Next
  End If
  %>
		
	  </table>
       <br/>
	   <div style="text-align:center">
	    <input name="Submit" type="submit"  class="button" disabled value="��������">
		
	   </div>
	 </form>
  
		  
		  
	   <div class="attention">
<strong>�ر����ѣ�</strong><br/>
1�������ܽ�����ҵ�û����š�<br/>
2�����ݱ���ѡ�е�Ϊ��ǰ��̳��ʹ��������ظ��������ݵı�һ�������ÿ�����е�����Խ����̳������ʾ�ٶ�Խ�죬�������е����������ݱ��е������г������������ʱ��������һ�����ݱ���������������,���ᷢ����̳�ٶȿ�ܶ�ܶࡣ<br/>
3����Ҳ���Խ���ǰ��ʹ�õ����ݱ����������ݱ����л�����ǰ��ʹ�õ��������ݱ���ǰ��̳�û�����ʱĬ�ϵı����������ݱ�<br/>
4�����������ǰ����ʹ�õ����ݱ����м�¼�����ݱ����ϵͳ�Դ����ݱ�����ɾ����
</div>
</body>
</html>
<%
End Sub




Set KS=Nothing
CloseConn
%>