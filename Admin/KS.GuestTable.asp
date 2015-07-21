<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<%
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
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
<title>论坛数据表管理</title>
<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>
<script language="JavaScript" src="../KS_Inc/Jquery.js"></script>
</head>
<body>
<ul id='mt'> <div id='mtl'>论坛数据表管理</div></ul>
	  <table width="100%" align='center' border="0" cellpadding="0" cellspacing="0">
	  <form name="myform" action="KS.GuestTable.asp?action=ModifySave" method="post">
      <tr class="sort">
	    <td>序号</td>
	    <td>表名称</td>
	    <td>类型</td>
		<td>当前默认</td>
		<td>记录数</td>
		<td>说明</td>
		<td>管理操作</td>
	  </tr>
<%
  If TableXML.DocumentElement.SelectNodes("item").length=0 Then
      Response.Write "<tr class='list'><td colspan=7 height='25' class='splittd' align='center'>您没有添加小论坛数据表!</td></tr>"
  Else
	  N=0
	  For Each Node In TableXML.DocumentElement.SelectNodes("item")
	  %>
			  <tr  onmouseout="this.className='list'" onMouseOver="this.className='listmouseover'">               
			   <td class='splittd' height="30" align="center"><%=Node.SelectSingleNode("@id").text%></td>
			   <td class='splittd' height="30"><%=Node.SelectSingleNode("tablename").text%></td>
			   <td class='splittd' style='text-align:center' height="30"><%
			   if Node.SelectSingleNode("@issys").text="1" then
			    response.write "<span style='color:red'>系统</span>"
			   else
			    response.write "<span style='color:green'>自定义</span>"
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
				 <span style="color:#999999">删除</span>
				<%else%>
				 <a href="?action=del&itemid=<%=Node.SelectSingleNode("@id").text%>" onClick="return(confirm('确定删除该任务吗?'))">删除</a>
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
	    <input name="Submit" type="submit"  class="button" disabled value="批量设置">
		
	   </div>
	 </form>
  
		  
		  
	   <div class="attention">
<strong>特别提醒：</strong><br/>
1、本功能仅对商业用户开放。<br/>
2、数据表中选中的为当前论坛所使用来保存回复帖子数据的表，一般情况下每个表中的数据越少论坛帖子显示速度越快，当您上列单个帖子数据表中的数据有超过几万的帖子时不妨新添一个数据表来保存帖子数据,您会发现论坛速度快很多很多。<br/>
3、您也可以将当前所使用的数据表在上列数据表中切换，当前所使用的帖子数据表即当前论坛用户发贴时默认的保存帖子数据表。<br/>
4、以免出错，当前正在使用的数据表、已有记录的数据表或是系统自带数据表不允许删除。
</div>
</body>
</html>
<%
End Sub




Set KS=Nothing
CloseConn
%>