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
Set KSCls = New Convention_Main
KSCls.Kesion()
Set KSCls = Nothing

Class Convention_Main
        Private KS,DomainStr,GaoAndKuan
		Private TotalPage,MaxPerPage,adssql,RSObj,totalPut,CurrentPage,TotalPages,i,advlistact,px,adsrs
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
       Sub Kesion()
	   	    If Not KS.ReturnPowerResult(0, "KSMS20006") Then
			  Call KS.ReturnErr(1, "")
			  exit sub
			End If

	    Select Case KS.G("Action")
		 Case "Adw"
		   Call AdsAdw()
		 Case "Addads"
		   Call AdsAddads()
		 Case "Help"
		   Call AdsHelp()
		 Case "Adslist"
		   Call Adslist()
		 Case "Listip"
		   Call AdsListip()
		 Case "IPDel"
		   Call IPDel()
		 Case "Manage"
		   Call AdsManage()
		 case "AddPlace"
		   Call AddPlace()
		 case "AdwSave"
		   Call AdwSave()
		 case "CreateJs"
		   Call CreateJs(ks.chkclng(request("id")))
		   Response.Write("<script>alert('��ϲ,���ɹ��JS�ɹ�!');location.href='KS.Ads.asp?Action=Adw';</script>")
		 Case Else
		  Call AdsMain()
		End Select
	   End Sub
	   Sub AdsMain()
         With Response
		 
		   .Write "<html>"
			.Write"<head>"
			.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<script language=""JavaScript"" src=""../KS_Inc/Jquery.js""></script>"
			.Write "<script language=""JavaScript"" src=""../KS_Inc/Kesion.Box.js""></script>"
			.Write "<script type=""text/javascript"">"
			.Write "function AddPlace(id){new KesionPopup().PopupCenterIframe('���/�༭���λ','KS.Ads.asp?Action=AddPlace&id='+id,480,200,'no')}"
			.Write "</script>"

			.Write"</head>"
			.Write"<body scroll=no leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
            .Write "<table width='100%' height='100%' border='0' cellspacing='0' cellpadding='0'><tr><td height='25'>"
		    .Write "<ul id='menu_top'>"
			.Write "<li class='parent' onclick=""AddPlace(0)""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/unverify.gif' border='0' align='absmiddle'>���ӹ��λ</span></li>"
			.Write "<li class='parent' onclick=""Ads.location.href='KS.Ads.asp?Action=Addads'""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>���ӹ��</span></li>"
			.Write "<li class='parent' onclick=""Ads.location.href='KS.Ads.asp?Action=Help'""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/s.gif' border='0' align='absmiddle'>�鿴˵��</span></li><li></li>"
			.Write "<div>&nbsp;�鿴ѡ�"
			.Write "<input onclick=""Ads.location.href='?Action=Adslist'"" name=""Option1"" title=""�鿴�������"" type=""radio"">�������"
			.Write "<input onclick=""Ads.location.href='?type=img&Action=Adslist'"" name=""Option1"" title=""�鿴����ͼƬ���"" type=""radio"">ͼƬ���"
            .Write "<input onclick=""Ads.location.href='?type=txt&Action=Adslist'"" name=""Option1"" title=""�鿴�����ı����"" type=""radio"">�ı����"	
            .Write "<input onclick=""Ads.location.href='?type=click&Action=Adslist'"" name=""Option1"" title=""��������в鿴���й��"" type=""radio"">�������"	
            .Write "<input onclick=""Ads.location.href='?type=close&Action=Adslist'"" name=""Option1"" title=""�鿴������ͣ�Ĺ��"" type=""radio"">��ͣ���"	
            .Write "<input onclick=""Ads.location.href='?type=lose&Action=Adslist'"" name=""Option1"" title=""������ʧЧ�Ĺ��"" type=""radio"">���ڹ��"	
			.write "</ul>"
			.write "</tr><tr><td>"
			.Write " <iframe name=""Ads"" scrolling=""auto"" frameborder=""0"" src=""KS.Ads.asp?Action=Adw"" width=""100%"" height=""100%""></iframe>"
            .Write " </td></tr></table>"
		End With
  End Sub
  
  Sub CreateJs(id)
    
	dim param:param=" where show_flag=1"
	If ID<>0 Then param=param & " and place=" & id
	dim rs:set rs=server.createobject("adodb.recordset")
	rs.open "select * from KS_ADPlace" & param,conn,1,1
	if not rs.eof then
	    do while not rs.eof
				 dim rst:set rst=server.createobject("adodb.recordset")
				 dim str,i,placeId,SaveFilePath,placelei,placewid,placehei
				 i=0 : placeId=rs("place") :placelei=rs("placelei") : str=""
				placehei=rs("placehei")
				placewid=rs("placewid")
				
				GaoAndKuan=""
				
				if Not KS.IsNUL(placehei) then GaoAndKuan=" height="&placehei&" "
				if Not KS.IsNul(placewid) then GaoAndKuan=GaoAndKuan&" width="&placewid&" "		 
				 rst.open "select * from KS_Advertise where act=1 and place="& placeId & " order by AdOrderID,id",conn,1,1
				 select case placelei
				   case 1
					 str="var GetRandomn = 1;" & vbcrlf
					 str=str & "function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}" & vbcrlf
					 str=str &" var a" & placeId & "=new Array();" & vbcrlf
					 str=str & "var t"&placeId & "=new Array();" &vbcrlf
					 str=str & "var ts" & placeId &"=new Array();" & vbcrlf
					 do while not rst.eof
					   if rst("xslei")="swf" then
					    str=str & "a" & placeId & "[" & i & "]=""" & DggtXs(rst) & """;" & vbcrlf
					   else
					    str=str & "a" & placeId & "[" & i & "]=""<span onclick=\""addHits" & placeId&"(" & rst("clicks") &"," & rst("id") & ")\"">" & DggtXs(rst) & "</span>"";" & vbcrlf
					   end if
					   str=str & "t" & placeId&"["&i&"]=" & rst("class") &";" &vbcrlf
					   str=str & "ts" & placeId&"["&i&"]=""" & formatdatetime(rst("lasttime"),2) & """;" & vbcrlf
					   i=i+1
					  rst.movenext
					 loop
					 str=str & "var temp" & placeId & "=new Array();" &vbcrlf
					 str=str & "var k=0;" & vbcrlf
					 str=str & "for(var i=0;i<a" & placeId &".length;i++){" &vbcrlf
					 str=str & "if (t" & placeId &"[i]==1){" & vbcrlf
					 str=str & "if (checkDate"&placeId&"(ts" & placeId&"[i])){" &vbcrlf
					 str=str &"	temp"& placeId&"[k++]=a" &placeId&"[i];" & vbcrlf & "}"&vbcrlf
					 str=str &"	}else{"&vbcrlf
					 str=str &" temp" & placeID&"[k++]=a" & placeID&"[i];" & vbcrlf &"}"&vbcrlf
					 str=str & "}" & vbcrlf
					 
					 str=str & "if (temp"&placeId & ".length>0){"&vbcrlf
					 str=str & "GetRandom(temp" & placeId & ".length);" & vbcrlf
					 str=str & "document.write(a" &placeId &"[GetRandomn-1]);" & vbcrlf
					 str=str & "}"&vbcrlf
					 str=str & getClicks(placeId)
					 
				  case 2
				   do while not rst.eof
				     str=str &"if(" & rst("class")&"==0 || (" & rst("class") & "==1 && checkDate"&placeId &"('" & rst("lasttime") &"'))){" &vbcrlf
					 if rst("xslei")="swf" then
					  str=str &"document.writeln(""" & DggtXs(rst) & "<br/>"");" & vbcrlf
					 else
					  str=str &"document.writeln(""<span onclick=\""addHits" & placeId&"(" & rst("clicks") &"," & rst("id") & ")\"">" & DggtXs(rst) & "</span><br/>"");" & vbcrlf
					 end if
					 str=str & "}" &vbcrlf
					rst.movenext
				   loop
					 str=str & getClicks(placeId)
				  case 3
				   do while not rst.eof
				     str=str &"if(" & rst("class")&"==0 || (" & rst("class") & "==1 && checkDate"&placeId &"('" & rst("lasttime") &"'))){" &vbcrlf
					 if rst("xslei")="swf" then
					  str=str &"document.write(""" & DggtXs(rst) & "&nbsp;"");" & vbcrlf
					 else
					  str=str &"document.write(""<span onclick=\""addHits" & placeId&"(" & rst("clicks") &"," & rst("id") & ")\"">" & DggtXs(rst) & "</span>&nbsp;"");" & vbcrlf
					 end if 
					 str=str &"}" &vbcrlf
					rst.movenext
				   loop
					 str=str & getClicks(placeId)
				  case 4
				   str="document.write('<marquee  direction=""up"""&GaoAndKuan&">');" & vbcrlf
				   do while not rst.eof
				     str=str &"if(" & rst("class")&"==0 || (" & rst("class") & "==1 && checkDate"&placeId &"('" & rst("lasttime") &"'))){" &vbcrlf
					 if rst("xslei")="swf" then
					 str=str &"document.write(""" & DggtXs(rst) & "<br/><br/>"");" & vbcrlf
					 else
					 str=str &"document.write(""<span onclick=\""addHits" & placeId&"(" & rst("clicks") &"," & rst("id") & ")\"">" & DggtXs(rst) & "</span><br/><br/>"");" & vbcrlf
					 end if
					 str=str &"}" &vbcrlf
					rst.movenext
				   loop
				   str=str &"document.write(""</marquee>"");" & vbcrlf
				   str=str & getClicks(placeId)
				  case 5
				   str="document.write('<marquee"&GaoAndKuan&">');" & vbcrlf
				   do while not rst.eof
				     str=str &"if(" & rst("class")&"==0 || (" & rst("class") & "==1 && checkDate"&placeId &"('" & rst("lasttime") &"'))){" &vbcrlf
					 if rst("xslei")="swf" then
					 str=str &"document.write(""" & DggtXs(rst) & "&nbsp;"");" & vbcrlf
					 else
					 str=str &"document.write(""<span onclick=\""addHits" & placeId&"(" & rst("clicks") &"," & rst("id") & ")\"">" & DggtXs(rst) & "</span>&nbsp;"");" & vbcrlf
					 end if
					 str=str &"}" &vbcrlf
					rst.movenext
				   loop
				   str=str &"document.write(""</marquee>"");" & vbcrlf
				   str=str & getClicks(placeId)
				  case 6
				   do while not rst.eof
					 call gaokuan(rs,rst)
					 str=str &"if(" & rst("class")&"==0 || (" & rst("class") & "==1 && checkDate"&placeId &"('" & rst("lasttime") &"'))){" &vbcrlf
					str=str & "window.open('"&DomainStr&"plus/ShowA.asp?Action=AdOpen&i="&rst("id")&"','" & KS.Setting(0) & "������"&rst("id")&"','"&GaoAndKuan&"');" &vbcrlf
					str=str &"}" &vbcrlf
		
					rst.movenext
				   loop
				   str=str & getClicks(placeId)
				  case 7
					 str="var GetRandomn = 1;" & vbcrlf
					 str=str & "function GetRandom(n){GetRandomn=Math.floor(Math.random()*n+1)}" & vbcrlf
					 str=str &" var a" & placeId & "=new Array();" & vbcrlf
					 str=str &" var gk" & placeId & "=new Array();" & vbcrlf
					 str=str & "var t"&placeId & "=new Array();" &vbcrlf
					 str=str & "var ts" & placeId &"=new Array();" & vbcrlf
					 do while not rst.eof
					   str=str & "t" & placeId&"["&i&"]=" & rst("class") &";" &vbcrlf
					   str=str & "ts" & placeId&"["&i&"]=""" & formatdatetime(rst("lasttime"),2) & """;" & vbcrlf
					   str=str & "a" & placeId & "[" & i & "]="""&DomainStr&"plus/ShowA.asp?Action=AdOpen&i="&rst("id")&""";" & vbcrlf
					   call gaokuan(rs,rst)
					   str=str & "gk" & placeId & "[" & i & "]="""&GaoAndKuan&""";" & vbcrlf
					   i=i+1
					  rst.movenext
					 loop
					 str=str & "var temp" & placeId & "=new Array();" &vbcrlf
					 str=str & "var k=0;" & vbcrlf
					 str=str & "for(var i=0;i<a" & placeId &".length;i++){" &vbcrlf
					 str=str & "if (t" & placeId &"[i]==1){" & vbcrlf
					 str=str & "if (checkDate"&placeId&"(ts" & placeId&"[i])){" &vbcrlf
					 str=str &"	temp"& placeId&"[k++]=a" &placeId&"[i];" & vbcrlf & "}"&vbcrlf
					 str=str &"	}else{"&vbcrlf
					 str=str &" temp" & placeID&"[k++]=a" & placeID&"[i];" & vbcrlf &"}"&vbcrlf
					 str=str & "}" & vbcrlf
					 str=str & "if (temp"&placeId & ".length>0){"&vbcrlf
					 str=str & "GetRandom(temp" & placeId & ".length);" & vbcrlf
					 str=str & "window.open(temp" &placeId &"[GetRandomn-1],'"&KS.Setting(0)&"������',gk"&PlaceId&"[GetRandomn-1]);" & vbcrlf
		             str=str & "}"&vbcrlf
					 str=str & getClicks(placeId)
				 end select	
				   rst.close : set rst=nothing
				 SaveFilePath = KS.Setting(3) & KS.Setting(93) 
				 Call KS.CreateListFolder(SaveFilePath)
				 Call KS.WriteTOFile(SaveFilePath& placeId & ".js", str)
		  RS.MoveNext
	   Loop
	end if
	rs.close
	set rs=nothing
    
  End Sub
  
  function getClicks(placeId)
   Dim str
   str="function addHits" & placeId&"(c,id){if(c==1){try{jQuery.getScript('" & domainStr &"plus/ajaxs.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}" & vbcrlf
   str=str & "function checkDate" & placeId&"(date_arr){" &vbcrlf
   str=str &" var date=new Date();" &vbcrlf
   str=str &" date_arr=date_arr.split(""-"");" &vbcrlf
   str=str & "var year=parseInt(date_arr[0]);" & vbcrlf
   str=str & "var month=parseInt(date_arr[1])-1;" & vbcrlf
   str=str & "var day=0;" & vbcrlf
   str=str & "if (date_arr[2].indexOf("" "")!=-1)" & vbcrlf
   str=str & "day=parseInt(date_arr[2].split("" "")[0]);" & vbcrlf
   str=str & "else" & vbcrlf
   str=str & "day=parseInt(date_arr[2]);" &vbcrlf
   str=str & "var date1=new Date(year,month,day);" & vbcrlf
   str=str & "if(date.valueOf()>date1.valueOf())" & vbcrlf
   str=str &" return false;" &vbcrlf
   str=str &"else" &vbcrlf
   str=str &" return true" & vbcrlf
   str=str &"}" &vbcrlf
   getClicks=str
  end function
  
  Function DggtXs(rst)
    dim str,ttarg,GaoAndKuan,GKCss
	if rst("window")=0 then
		ttarg = "_blank"
	else 
		ttarg = "_top" 
	end if
    if isnumeric(rst("hei")) then
		GaoAndKuan=" height="&rst("hei")&" "
		GKCss="height:" &rst("hei")&"px;"
	else
		
		if right(rst("hei"),1)="%" then
		if isnumeric(Left(rst("hei"),len(rst("hei"))-1))=true then
		 GaoAndKuan=" height="&rst("hei")&" "
		 GKCss="height:" &rst("hei")&";"
		end if
		end if
		
		end if
		
		
		if isnumeric(rst("wid")) then
		GaoAndKuan=GaoAndKuan&" width="&rst("wid")&" "
		GKCss=GKCss&"width:" &rst("wid")&"px;"
		else
		if right(rst("wid"),1)="%" then
		if isnumeric(Left(rst("wid"),len(rst("wid"))-1))=true then 
		GaoAndKuan=GaoAndKuan&" width="&rst("wid")&" "
		GKCss=GKCss&"width:" &rst("wid")&";"
		end if
		end if
	end if	
     Select Case rst("xslei")
		   Case "txt"
		    str="<a title=""" & rst("sitename") & """  href=""" & rst("url") & """ target=""" & ttarg & """>" & rst("sitename") & "</a>"
		   Case "gif"
		    str="<a href=""" &  rst("url") & """ target=""" & ttarg & """><img  alt=""" & rst("sitename") & """  border=""0"" " & GaoAndKuan&" src=""" & rst("gif_url") & """></a>"
		   Case "swf"
		   str="<a href=""" & rst("url") & """ onclick=""addHits" & rst("place")&"(" & rst("clicks") &"," & rst("id") & ")"" target=""" & ttarg & """ hidefocus><button disabled style=""cursor:pointer;border:none;" &GKCss &"""><object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"" "&GaoAndKuan &">"
		   str=str & "<param name=""movie"" value=""" & rst("gif_url") &""" />"
		   str=str & "<param name=""quality"" value=""high"" />"
		   str=str & "<param name=""wmode"" value=""transparent"" />"
		   str=str & "<embed src=""" & rst("gif_url") & """ quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" " &GaoAndKuan &"></embed>"
		   str=str & "</object></button></a>"
		  '  str="<EMBED src=""" & rst("gif_url") & """   " & GaoAndKuan & """  quality=""high"" TYPE=""application/x-shockwave-flash""></EMBED>"
		   Case "dai"
		    str="<iframe marginwidth=""0"" marginheight=""0""  frameborder=""0"" bordercolor=""000000"" scrolling=""no""  name=""���"" src=""" & DomainStr & "plus/ShowA.asp?Action=Daima&id=" & rst("id") & """  " & GaoAndKuan &"></iframe>"
		  Case else
		    str="<a href=""" & rst("url") & """ target=""" & ttarg & """><img alt=""" & rst("sitename") & """  border=""0"" " & GaoAndKuan &" src=""" & rst("gif_url") & """ /></a>"
	End Select
	str=Replace(Replace(Replace(Replace(str, Chr(13)& Chr(10), ""),"'","\'"),"""","\"""),vbcrlf,"") 
	DggtXs=str	
  End Function
  
  
Sub gaokuan(rs,adsrs) 
		if not KS.IsNul(adsrs("hei")) and adsrs("hei")<>"0" then
			if isnumeric(adsrs("hei")) then
			  GaoAndKuan="height="&adsrs("hei")
			else
				 if right(adsrs("hei"),1)="%" then
				   if isnumeric(Left(len(adsrs("hei"))-1))=true then
					 GaoAndKuan="height="&adsrs("hei")
				   end if
				 end if
			end if
		else
		  GaoAndKuan="height="&rs("placehei")
		end if
		
	  If Not KS.IsNul(adsrs("wid")) and adsrs("wid")<>"0" Then
		if isnumeric(adsrs("wid")) then
		   GaoAndKuan=GaoAndKuan&",width="&adsrs("wid")
		else
			if right(adsrs("wid"),1)="%" then
				if isnumeric(Left(len(adsrs("wid"))-1))=true then 
				 GaoAndKuan=GaoAndKuan&",width="&adsrs("wid")
				end if
			end if
		end if
	  Else
	    GaoAndKuan=GaoAndKuan&",width="&rs("placewid")
	  End If
	End Sub
  
  
  Sub AddPlace()
    dim id:id=KS.ChkClng(Request("id"))
	dim placename,show_flag,placelei,placewid,placehei,substr
	If Id<>0 Then
      dim rs:set rs=server.createobject("adodb.recordset")
	  rs.open "select top 1 * from KS_ADPlace where place=" & id,conn,1,1
	  if not rs.eof then
	    placename=rs("placename")
		show_flag=rs("show_flag")
		placewid=rs("placewid")
		placehei=rs("placehei")
		placelei=rs("placelei")
		substr="ȷ���޸�"
	  end if
	  rs.close
	  set rs=nothing
	else
	   id=0 : show_flag=1 : placewid=468 : placehei=60 : placelei=1 : substr="�������λ"
	End If
  %>
  <html>
	<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
	<script language="javascript">
	<!--
	function isok(theform)
	{
		if (theform.placename.value=="")
		{
			alert("����д���λ��ʶ��");
			theform.placename.focus();
			return (false);
		}
	}
	-->
	</script>
	<body>

              <table border=0 width=100% cellspacing=01 cellpadding=2 style="border-collapse: collapse" bordercolor="#111111">
				<form method="POST"  action="KS.Ads.asp" onSubmit="return isok(this)">
                 <input type="hidden" name="action" value="AdwSave">
                 <input type="hidden" name="place" value="<%=id%>">

              <tr class="tdbg"> 
                <td class="clefttitle" width="130" align="right">���λ����:</td>
				<td> 
				 <input type=text name=placename class='textbox' value="<%=placename%>" size=20 maxlength=30><font color="#FF0000">15������</font></td>
			  </tr>
			  <tr class="tdbg">
                <td class="clefttitle" align="right">��ʾ���:</td>
				<td><select class='textbox' name="show_flag">
							   <option value="1"<%If show_flag="1" then response.write " selected"%>>��</option>
							   <option value="0"<%If show_flag="0" then response.write " selected"%>>�ر�</option>
				      </select>
			    </td>
			  </tr>
			  <tr class="tdbg">
			    <td class="clefttitle" align="right">���:</td>
				<td><input type=text class='textbox' name="placewid"  value="<%=placewid%>" size="6" maxlength=30> px</td>
			  </tr>
			  <tr class="tdbg">
			    <td class="clefttitle" align="right">�߶�:</td>
				<td><input class='textbox' type=text name="placehei" value="<%=placehei%>" size="6" maxlength=30> px</td>
			  </tr>
			  <tr class="tdbg">
			    <td class="clefttitle" align="right">����:</td>
				<td><%Call Ggwlei(placelei)%>&nbsp; </td>
			  </tr>
			  <tr class="tdbg">
			      <td colspan="2" height="40" style="text-align:center">
                  <input class="button" type="submit" value="<%=substr%>" name="B1">
				  <input type="button" class="button" onClick="parent.closeWindow();" value="�رմ���">
               </td>
              </tr>
            </form>
          </table>
       
   
  <%
  End Sub
  
  Sub AdwSave()
       Dim SQLStr
       Dim RSObj:set RSObj=server.createobject("adodb.recordset")
		if  KS.ChkClng(request("place"))=0 then
			SqlStr="select top 1 * From KS_ADPlace "
			RSObj.open SqlStr,Conn,1,3
			RSObj.AddNew
		else
			SqlStr="select top 1 * From KS_ADPlace where place="&trim(request("place"))
			RSObj.open SqlStr,Conn,1,3
		end if
			RSObj(1) = trim(request("placename"))
			RSObj(2)= trim(request("placelei"))
			RSObj(3)= trim(request("placehei"))
			RSObj(4)= trim(request("placewid"))
			RSObj(5)=trim(request("show_flag"))
			RSObj.update
			RSOBj.MoveLast
			'����JS�ļ�
			Call CreateJS(rsobj("place"))
			RSObj.close
			set RSObj=nothing
		    if  KS.ChkClng(request("place"))=0 then
			  response.write "<script>if(confirm('���λ��ӳɹ�,���������?')){location.href='?action=AddPlace'}else{parent.Ads.location.href='KS.Ads.asp?Action=Adw';parent.closeWindow()}</script>"
			else
			  response.write "<script>alert('���λ�޸ĳɹ�!');parent.Ads.location.href='KS.Ads.asp?Action=Adw';parent.closeWindow();</script>"
			end if
  End Sub
  
  '�鿴����
  Sub AdsHelp()
  	    With Response
		 .Write "<html>"
		 .Write"<head>"
		 .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		 .Write"<link href=""Include/admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		 .Write"</head>"
		 .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
      End With %>
		<br>
		<div align="center">
		  <center>
		  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="95%" id="AutoNumber1">
			<tr>
			  <td width="100%"><b>һ��ϵͳ�ص㣺</b><ol>
				<li>ͨ����ϵͳ�������ò��������������λ</li>
				<li>�����λ�п����������ѭ�����ŵĹ����</li>
				<li>
				���λ�еĹ��������7����ʾ��ʽ,��&quot;�����ʾһ��&quot;��&quot;������������&quot;��&quot;������������&quot;��&quot;���Ϲ�������&quot;��&quot;�����������&quot;��&quot;�����������&quot;��&quot;�����������&quot;������˵������� 
				<a href="addadw.asp#˵��">���λ��Ŀ�й��λ��ʾ��ʽ˵��</a></li>
				<li>�����������GIF��SWF��Flash�������ı������������ʾ����</li>
				<li>���λ�ϵĹ����Ϊѭ�����ţ�ÿ����ʾ���Ǹù��λ�еȴ�ʱ������Ҵ�������״̬�Ĺ����</li>
				<li>�ɶ�������������ʱִ����ͣ������޸ġ�ɾ���Ȳ���</li>
				<li>ɾ��ĳһ�����ʱ��������ص���ʾ�������¼Ҳ����֮ɾ��</li>
				<li>����ʵ�ֹ��λ��ҳ�淢��,������ġ�<a href="#��">���λ����˵��</a>��</li>
				<li>���ֹ�沥���������ƹ�沥��״̬��������������ơ���ʾ������ơ����ʱ�����Ƶ�</li>
				<li>���ƵĹ����ʼ�¼������ʾ�������ߡ�����ߵ�IP��ַ</li>
				<li>���д������������ʱ����ͨ������������ѯ������Զ�����в���</li>
			  </ol>
			  <p><b>����ʹ��˵����</b></p>
			  <ol>
				<li>�� <font color="#FF0000">�� �� λ</font> һ���ڿ�����¹��λ���޸ġ�ɾ�����й��λ��ʶ����ѯ���λID</li>
				<li>�� <font color="#FF0000">��ӹ�� </font>һ���ڿ�Ϊĳ���λ���һ���¹����</li>
				<li>�� <font color="#FF0000">������� </font>
				һ������ʾ��ǰ���д�����������״̬�Ĺ����������ִ���޸ġ�ɾ������ͣ��Ԥ������</li>
				<li>�� <font color="#FF0000">ͼƬ��� </font>
				һ������ʾ��ǰ���д�����������״̬�ķ��ı������������ִ���޸ġ�ɾ������ͣ��Ԥ������</li>
				<li>�� <font color="#FF0000">�ı���� </font>
				һ������ʾ��ǰ���д�����������״̬�Ĵ��ı������������ִ���޸ġ�ɾ������ͣ��Ԥ������</li>
				<li>�� <font color="#FF0000">������� </font>�� 
				����������Ĳ�ͬ˳����ʾ��������ĵ������������ִ���޸ġ�ɾ������ͣ�����Ԥ������</li>
				<li>�� <font color="#FF0000">��ʾ���� </font>�� 
				����ʾ�����Ĳ�ͬ˳����ʾ�����������ʾ����������ִ���޸ġ�ɾ������ͣ�����Ԥ������</li>
				<li>�� <font color="#FF0000">��ͣ�б� </font>�� 
				��ʾ��ǰ���д�����ͣ����״̬�Ĺ����������ִ���޸ġ�ɾ�������Ԥ������</li>
				<li>�� <font color="#FF0000">ʧЧ�б� </font>�� 
				��ʾ��ǰ�����Ѿ�ʧЧ�Ĺ����������ִ���޸ġ�ɾ�������Ԥ������</li>
				<li>�� <font color="#FF0000">�� �� λ </font>�� 
				ͨ��ĳ���λ���ӣ�����ʾ�ù��λ�µ����й����������ִ���޸ġ�ɾ������ͣ��Ԥ������</li>
			  </ol>
			  <p><b><a name="��">��</a>�����λ����˵����</b></p>
			  <ol>
				<li>ȷ�� <font color="#FF0000">ʵ��ҳ���е�Ԥ�����λ��</font> Ӧ�����ĸ� 
				<font color="#FF0000">ͨ����ϵͳ���õĹ��λ</font> </li><br><br>
				<li>ͨ�� <font color="#FF0000">�� �� λ</font> һ�����õ����� <font color="#FF0000">
				���λID</font></li><br><br>
				<li>Ȼ���±�����ݿ�����Ԥ�����λ�ã�ע�⽫���е� <font color="#FF0000">���λID</font> ��Ӧ��ȷ</li><br><br>
			   
		
				  <input type="text" name="T1" size="100" value='<script type="text/javascript" src="<%=KS.GetDomain%><%=KS.Setting(93)%>���λID.js"></script>'>
</li>
			  </ol>
		
			  <p><b>�ġ�ע�����</b></p>
			  <ol>
				<li>ÿ�����λ�е����й������ʾͼƬ��ȡ��߶�Ӧ��������һ�£���Ӧע������λԤ����ʵ��ҳ��λ�÷��һ��</li><br><br>
				<li>��ʵ��ҳ��Ԥ���Ĳ�ͬ���λ�о�������ʹ�ñ�ϵͳ���õĲ�ͬ���λ�������ɾ����ܶ��Ͷ�Ź��</li><br><br>
				<li>ͬһ���λ��,���ֹ������ͼƬ�����������Ҫ���ʹ��</li>
			  </ol>
			  <p><font color="#FF0000"><b>��ע��ʵ��ҳ���е�Ԥ�����λ�� </b></font>
			  ��ָ��������վҳ����Ҫ���ù���λ�ã���������ͨ����ϵͳ���õĹ��λ����</p>
			  <p>��</td>
			</tr>
		  </table>
		  </center>
		</div>
<%
  End Sub
  
  '���ӹ��λ
  Sub AdsAdw()
		if KS.G("job")="del" then
			if  isnumeric(request("place"))=true then
				set RSObj=server.createobject("adodb.recordset")
				SqlStr="select * From KS_ADPlace where place="&KS.ChkClng(request("place"))
				RSObj.open SqlStr,Conn,3,3
				RSObj.delete
				RSObj.close:set RSObj=nothing
				Conn.Execute("Delete From KS_Advertise Where Place="&KS.ChkClng(request("place")))
				Response.Redirect "?Action=Adw"
			end if
		end if
  
  %>
  <html>
  <head>
  <link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
  <meta http-equiv=Content-Type content="text/html; charset=gb2312">
  </head>
  <body>
  
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="63" align="center" height="23" class="sort"> <font color="#FF0000">���λID</font></td>
                <td width="192" height="20" align="center" class="sort">���λ����</td>
                <td width="62" height="20" align="center" class="sort">���</td>
                <td width="58" height="20" align="center" class="sort">�߶�</td>
                <td width="119" align="center" class="sort">���λ��ʾ��ʽ</td>
                <td width="234" align="center" class="sort">��ʾ���</td>
                <td width="219" align="center" class="sort">�� ��</td>
              </tr>
<%
	CurrentPage=KS.ChkClng(Request("page"))
	if CurrentPage=0 then CurrentPage=1
	MaxPerPage=20

	Dim RSObj:Set RSObj=server.createobject("adodb.recordset")
	Dim SqlStr:SqlStr="select * From KS_ADPlace order by place"
	RSObj.open SqlStr,Conn,1,1
	If RSObj.EOF Then
			response.write ("<tr><td colspan=10 align='center' class='splittd'>û������κι��λ!</td></tr>")
	Else
					        totalPut = RSObj.RecordCount
		
							If CurrentPage < 1 Then	CurrentPage = 1
		
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (totalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
		
							If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RSObj.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
							End If
	 dim i:i=0	
	do while not RSObj.eof 
%>
              <form method="POST" action="?job=add&Action=Adw"  onSubmit="return isok(this)">
                <tr height=25 class='splittd' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='splittd'"> 
                  <td width="63" align="center" class='splittd' nowrap><font color=red><%=RSObj(0)%></font> <input type=hidden name=place value="<%=RSObj(0)%>" >
                  ��</td>
                  <td align="center" nowrap class='splittd'> 
                   <%=RSObj(1)%>
                  </td>
                  <td width="62" align="center" nowrap class='splittd'> 
                   <%=RSObj(4)%></td>
                  <td width="58" align="center" nowrap class='splittd'><%=RSObj(3)%></td>
                  <td width="119" align="center" nowrap class='splittd'>
                      <%=Ggwlx(RSObj("place"))%>
                  </td>
                  <td align="center" class='splittd'> 
                    <%if RSObj(5)=1 then%>
                    <font color="green">����</font>
                    <%else%>
                    <font color="#FF0000">�ر�</font> 
                    <%end  if%>
                  </td>
                  <td width="219" align="center" nowrap class='splittd'> 
                    <a href='javascript:parent.AddPlace(<%=rsobj(0)%>)'>�޸�</a>
                    <a href="?job=del&Action=Adw&place=<%=RSObj(0)%>" onClick="return(confirm('ȷ��ɾ���ù��λ��?'))">ɾ��</a>&nbsp; <a href=?Action=Adslist&type=place&place=<%=RSObj(0)%>>���й����</a> 
                  <a href=KS.Ads.asp?Action=Manage&id=<%=RSObj(0)%>&job=yulanggw>Ԥ��</a>
                  <a href=KS.Ads.asp?Action=CreateJs&id=<%=RSObj(0)%>>����JS</a>
				  
				  </td>
                </tr>
              </form>
              <%
			  i=i+1
			  if i>=maxperpage then exit do
			  RSObj.movenext
      loop
  End If
      %>
	  <tr valign=middle>
	        <td colspan="3">
			 <input type="button" class="button" value="��ӹ��λ" onClick="parent.AddPlace(0)">
			 <input type="button" class="button" value="�������й��λJS" onClick="location.href='?action=CreateJs'"/>
			</td>
			<td colspan="10" align="right">
			<%
			 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			%>
		</td>
		</tr>
 </table>  
<%
      RSObj.close:set RSObj=nothing
%>
  <p align="left">
  <p align="left"><hr color="#808080" size="1">
<p align="left"><font color="#FF0000"><a name="˵��">���λ��ʾ��ʽ˵��</a>��</font></p>
<center>
  </p>
  <ul style="text-align:left">
    <li>�����ʾһ�������ǽ����λֱ������ĳҳ��һ�̶�λ�ã�����ͬһλ�������ʾ���λ�е����������������������ÿˢ��һ�ξͻ������ʾһ���µĹ����
    </li>
    <li>�����������룺���ϵ������Ź��λ�е��������������</li>
    <li>�����������룺�����Һ��Ź��λ�е��������������</li>
    <li>���Ϲ������룺���Ϲ�����ʾ���λ�е��������������</li>
    <li>����������룺���������ʾ���λ�е��������������</li>
    <li>����������ڣ�ҳ���ʱͬʱ����������ڣ�ÿ����������ʾһ��������������������ù��λ�е������������һ��</li>
    <li>  ����������ڣ�ҳ���ʱͬʱ����һ�����ڣ���ͬһ�����������ʾ���λ�е�������棬������ÿˢ��һ�ξͻ��ڵ��������и�����ʾһ���µĹ����
    </li>
  </ul>
  <p align="left"><font color=red> �����뷽����</font>
  <div align=left>
  <li><font color="#FF0000">����1��</font>��ģ��༭���в�����Ӧ�Ĺ��λ��ǩ,��{=GetAdvertise(���λID)}��
  <li><font color="#FF0000">����2��</font>���±����ݷŵ�Ԥ�����λ�ã��������е�<font color="#FF0000">���λID</font>��Ӧ��ȷ 
   <font color="#808080">���ڹ��λ�б��в鿴</font><font color="#FF0000">���λID</font>
  </div>
  <input type="text" name="T1" size="100" value='<script type="text/javascript" src="<%=KS.GetDomain%><%=KS.Setting(93)%>���λID.js"></script>'>
</p>
</body>
</html>
<%End Sub
'���ó��ù��λ���������˵�
Sub Ggwlei(shu) '���ڱ�ʾ���͵���
%>
 <select size=1 name=placelei>
                    <option value=1 <% if shu=1 then%>selected<%end if%>>�����ʾһ��</option>
                    <option value=2 <% if shu=2 then%>selected<%end if%>>������������</option>
                    <option value=3 <% if shu=3 then%>selected<%end if%>>������������</option>
                    <option value=4 <% if shu=4 then%>selected<%end if%>>���Ϲ�������</option>
                    <option value=5 <% if shu=5 then%>selected<%end if%>>�����������</option>
                    <option value=6 <% if shu=6 then%>selected<%end if%>>�����������</option>
                    <option value=7 <% if shu=7 then%>selected<%end if%>>�����������</option>
</select>
<%
  End Sub
  
  '���ӹ��
Sub AdsAddads()
Dim CurrPath:CurrPath = KS.GetCommonUpFilesDir()
%>
<html>
<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
<script src="../KS_Inc/common.js" language="javascript"></script>
<script language="javascript">
<!--
function isok(theform)
{
    if (theform.name.value=="")
    {
        alert("����д������ƣ�");
        theform.name.focus();
        return (false);
    }
    if (theform.url.value=="")
    {
        alert("����д����URL��");
        theform.url.focus();
        return (false);
    }
    return (true);
}
-->
</script>
<%
Dim Ggw,sitename,url,intro,xslei,gif_url,wid,hei,window,classs,clicks,shows,lasttime,flag,AdorderID
Ggw=1:URL="http://":xslei="gif":gif_url="http://":wid="":hei="":clicks=0:window=0:classs="0":flag="Add":AdorderID=1:lasttime=now+30
if KS.G("job")="add" then
	Call  addrk():Exit Sub
ElseIf KS.G("job")="edit" then
 Dim KS_RS_Obj:Set KS_RS_Obj=Server.CreateObject("Adodb.Recordset")
 KS_RS_Obj.Open "Select top 1 * From KS_Advertise where id="&KS.ChkClng(KS.G("id")),Conn,1,1
  If Not KS_RS_Obj.Eof Then
  Ggw      = KS_RS_Obj("Place")
  sitename = KS_RS_Obj("sitename")
  url      = KS_RS_Obj("url")
  intro    = KS_RS_Obj("intro")
  xslei    = KS_RS_Obj("xslei")
  gif_url  = KS_RS_Obj("gif_url")
  wid      = KS_RS_Obj("wid")
  Hei      = KS_RS_Obj("Hei")
  window   = KS_RS_Obj("window")
  classs   = KS_RS_Obj("class")
  clicks   = KS_RS_Obj("clicks")
  shows    = KS_RS_Obj("Shows")
  lasttime = KS_RS_Obj("lasttime")
  AdorderID = KS_RS_Obj("AdorderID")
  End If
  KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
  flag="Edit"
end if
%>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
.STYLE3 {color: #3300FF}
-->
</style>
 <table border=0 width=100% cellspacing=0 cellpadding=0>
            <tr> 
              <td align=center> 
                <%
if KS.G("job")="edit" then
%>
<font size="2" color=red><b>�޸Ĺ����</b></font> 
<%else%>
                <font size="2" color=red><b>����¹����</b></font> 
                <%
end if
%>
     <hr color="#808080" size="1"> 
	        </td>
            </tr>
          </table>
              <table border=0 width=100% cellspacing=1 cellpadding=2  class="ctable">
				<form method="POST"  name="myform"  action="?flag=<%=Flag%>&job=add&Action=Addads&id=<%=KS.G("id")%>" onSubmit="return isok(this)">
				 <input type="hidden" value="<%=request.ServerVariables("http_referer")%>" name="comeurl">
              <tr class='tdbg'> 
                <td class="clefttitle" width='85'>�������λ</td>
                <td colspan="2"> 
                <%
                Call  Ggwxlxx(Ggw) 
				%>              </td>
              </tr>
			  <tr class='tdbg'> 
                <td class="clefttitle">�������</td>
                <td colspan="2"> 
                  <input type="text" class='textbox' name="name" value="<%=sitename%>" size=30 maxlength=30>
                  ������15�����Ļ�30����ĸ����</td>
              </tr>
              <tr class='tdbg'> 
                <td class="clefttitle">����URL</td>
                <td colspan="2"> 
                  <input type=text class='textbox' name=url size=40 value="<%=url%>">
			    </td>
              </tr>
              <tr class='tdbg'> 
                <td class="clefttitle">���/����</td>
                <td width="200"> 
                  <textarea rows="5" class='textbox' name="intro" cols="48" style="height:60"><%=intro%></textarea></td>
                <td> <font color="#FF0000">��ʾ��</font><br>
                  <font color="#808080">�����Ƕ������뽫������������˴� ����URL��Ч<br>
                  �����ʾ���ı�������ʾΪ�������<br>
                  ֻ��GIFͼƬʱURL��д��Ч</font></font>                  </td>
              </tr>
              <tr class='tdbg'> 
                <td class="clefttitle">�������</td>
                <td colspan="2"> 
                  <input name="xslei" type="radio" value="gif" <%if xslei="gif" then response.write " checked"%>>GIFͼƬ 
                  <input type="radio" name="xslei" value="swf" <%if xslei="swf" then response.write " checked"%>><font siz=3 >Flash���� </font>
                  <input type="radio" name="xslei" value="txt" <%if xslei="txt" then response.write " checked"%>><font siz=3 >���ı� </font>    
                  <input type="radio" name="xslei" value="dai" <%if xslei="dai" then response.write " checked"%>>Ƕ�����                </td>
              </tr>
              <tr class='tdbg'> 
                <td class="clefttitle">ͼƬURL</td>
                <td colspan="2"> <input type=text class='textbox' name="gif_url"  size=40 value="<%=gif_url%>">&nbsp;<input type='button' class='button' name='Submit' value='ѡ���ַ...' onClick="OpenThenSetValue('Include/SelectPic.asp?Currpath=<%=CurrPath%>',550,290,window,document.myform.gif_url);">
                <font siz=3 > ��� </font>
                <input type=text name="wid" value="<%=wid%>" size=3 class='textbox' maxlength="4">
                <font siz=3 >�߶� </font> 
                  <input type=text name=hei value="<%=hei%>" size=3 class='textbox'  maxlength="4"><font siz=3 >&nbsp;</font><font color=red siz=3 > �����ǰٷֱȻ��Ĭ��</font> </td>
              </tr>
              <tr class='tdbg'> 
                <td class="clefttitle">���Ӵ򿪷�ʽ</td>
                <td colspan="2"> 
                  <select size=1 name=window>
                    <option value=0<%if window=0 then response.write " selected"%>>�´��ڴ�</option>
                    <option value=1<%if window=1 then response.write " selected"%>>ԭ���ڴ�</option>
                  </select>                </td>
              </tr>
              <tr class='tdbg'> 
                <td class="clefttitle">˳��ID</td>
                <td colspan="2"> 
				<input type=text name="AdorderID" value="<%=AdorderID%>" size=10 class='textbox' maxlength="4">&nbsp;(��ֵС�Ŀ�ǰ)
                 </td>
              </tr>
              <tr class='tdbg'> 
                <td class="clefttitle">��¼�����</td>
                <td colspan="2"> 
				<label><input type="radio" name="clicks" value="1"<%if clicks="1" then response.write" checked"%>>��¼</label>
				<label><input type="radio" name="clicks" value="0"<%if clicks="0" then response.write" checked"%>>����¼</label>
				<br/>
				<font color=green> &nbsp;1����¼������û�ÿ���һ�ι��ͻ��¼����,���Ի�ռ��һЩ���ݿ�ռ�,����û�������˵������ͻ���,���鲻Ҫ�����˹��ܣ�<br/>
&nbsp;2��ֻ��ͼƬ,���ֻ�flash��Ĺ��˴����ò���Ч��
				
				</font>
                 </td>
              </tr>
              <tr class='tdbg'> 
                <td class="clefttitle">ֹͣʱ�����</td>
                <td colspan="2"> 
				 <label><input type="radio" onClick="ss.style.display='none'" name="class"<%if classs="0" then response.write " checked"%> value="0">������</label>
				 <label><input type="radio" onClick="ss.style.display=''" name="class"<%if classs="1" then response.write " checked"%> value="1">����</label>
				 <div id="ss"<%if classs="0" then response.write " style='display:none'"%>>
				 <input type="text" name="lasttime" value="<%=formatdatetime(lasttime,2)%>"/>
				 ��ʽ��yyyy-mm-dd<br/>
				 
				 <font color=green>���ô˹��ܺ󣬵��ں�Ĺ���������ٲ���</font>
				 </div>
                 </td>
              </tr>

			  
             
              <tr class="tdbg"> 
                <td colspan=3 height="45" style="text-align:center"> 
                  <input type=submit class='button' value=' �� �� ' name=B1>
                  <input type=reset class='button' value=' �� д ' name=B2>                </td>
              </tr>
            </form>
          </table>
 </body>
</html>
<%
End Sub
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''�������Ϣ��⺯���������޸ġ�������֣�'''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Sub addrk()
	if KS.G("job")="add" then
	
	dim getname,geturl,getgif,getplace,getwin,getxslei,RSObj,adssql,getclass,getclicks,getshows,gettime,getintro,gethei,getwid,getAdorderID
	getname = Trim(Request("name"))
	geturl = Trim(Request("url"))
	getgif = Trim(Request("gif_url"))
	getplace =trim(Request("place"))
	getwin =trim(Request("window"))
	getxslei = trim(Request("xslei"))
	getclass=trim(Request("class"))
	getintro=trim(Request("intro"))
	getwid=trim(Request("wid"))
	gethei=trim(Request("hei"))
	getAdorderID=KS.ChkClng(Request("AdorderID"))
	
	if getxslei="txt" then
	getwid=0
	gethei=0
	end if
	
	
	getclicks=ks.chkclng(request("clicks"))
	if getclass="1" then
	  gettime=request("lasttime")
	else
	  gettime=now
	end if
	 if not isdate(gettime) then response.write "<script>alert('��ʾ��ֹ���ڣ���ʽ����!');history.back();</script>"
	gettime=formatdatetime(gettime,2)
	
	set RSObj=server.createobject("adodb.recordset")
	if  trim(KS.G("id"))="" then '��������������

	adssql="select top 1 * from KS_Advertise"
	RSObj.open adssql,Conn,1,3
	RSObj.AddNew
	else                                                '������޸Ĺ����
	adssql="select top 1 * from KS_Advertise where id="&KS.ChkClng(KS.G("id"))
	RSObj.open adssql,Conn,1,3
	end if
	RSObj("act") = 1
	RSObj("sitename") = getname
	RSObj("url") = geturl
	RSObj("gif_url") = getgif
	RSObj("place") = getplace
	RSObj("xslei") = getxslei
	RSObj("hei") = gethei
	RSObj("wid") = getwid
	RSObj("window") = getwin
	RSObj("class") = getclass
	RSObj("clicks") = getclicks
	RSObj("shows") = getshows
	RSObj("lasttime") = gettime
	RSObj("regtime") = Now()
	RSObj("time") = now()
	RSObj("intro")=getintro
	RSObj("AdorderID")=getAdorderID
	RSObj.update
	If KS.G("ID")="" Then
	 RSObj.MoveLast
	 Call KS.FileAssociation(1020,RSObj("ID"),getgif,0)
	Else
	 Call KS.FileAssociation(1020,RSObj("ID"),getgif,1)
	End If

	'����js
	Call CreateJs(getplace)
	
	RSObj.close
	set RSObj=nothing
	Conn.close
	set Conn=nothing
	if KS.g("id")<>"" then
	     %>
		 <script>alert('������޸ĳɹ�!');location.href='<%=KS.g("comeurl")%>';</script>"
		 <%
		 response.end
    else
	  %>
	     <script type="text/javascript">
    if (confirm('�������ӳɹ�,���������?')) {
	 location.href='KS.Ads.asp?flag=Add&Action=Addads';
	}else{
	 location.href='KS.Ads.asp?Action=Adw';
	}
   </script>

	  <%	
	end if
	end if
	End Sub
	'�������λ����ѡ��
	
	Sub Ggwxlxx(place) 'place �����ж�Ĭ��ѡ��
	%>
	  <select size=1 name=place>
	<%
	on error resume next
	Dim PRSObj:Set PRSObj=server.createobject("adodb.recordset")
	PRSObj.open "select * From KS_ADPlace",Conn,1,1
	do while not PRSObj.eof
	%>
	<option value="<%=PRSObj(0)%>" <% if PRSObj(0)=place then :Response.Write "selected":end if%>><%=PRSObj(1)%></option>
	 <%PRSObj.movenext
	   loop
	   PRSObj.close
	   Set PRSObj=nothing%>              
	  </select> 
<%
  End Sub
  
  Sub Adslist()
%>
<html>
<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	margin-left: 5px;
	margin-top: 2px;
}
-->
</style>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="1"  width="100%" class=tableBorder >
   <form method=post action="?type=search&Action=Adslist">
    <tr>
      <td width="100%" style="text-align:left">��������=&gt;&gt;
      <select size="1" name="adorder" >
<option value="id">���ID</option>
<option value="name">���ƹؼ���</option>
</select> <input type="text" name="nr" size="20">
<input type="submit" value="�� ѯ" name="B1" class=button>
<input type="button" onclick='location.href="KS.Ads.asp?Action=Adw";' value="�� ��" name="B1" class=button>

</td>
    </tr></form>
  </table>
  </center>
</div>
          <table border=0 width=100% cellspacing=3 cellpadding=3>
            <tr> 
              <td align=center> 
                <%
                  if request("px")="" then
                  px="desc"
                  else
                  px=""
                  end if
                  
                   Select Case KS.G("type")
                   
                          Case "img"
                           adssql="select * from KS_Advertise where act=1 and (xslei='gif' or xslei='swf') order by regtime "&px
                %>
                <b>�������ŵ�ͼƬ�������б�</b>  <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>��</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>��</a>
               
			    <%        Case "txt"
                           adssql="select * from KS_Advertise where act=1 and xslei='txt' order by regtime "&px
                %>
                <b>�������ŵĴ��ı�������б�</b>  <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>��</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>��</a>
                <%
                          Case "close"
                           adssql="select * from KS_Advertise where act=0 order by regtime "&px

                %>
                <b>������ͣ��δʧЧ�Ĺ�����б�</b>  <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>��</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>��</a>
                <%
                          Case "lose"
                           adssql="select * from KS_Advertise where class=1 and  datediff(" & DataPart_D&",lasttime," & sqlnowstring &")>=0 order by regtime "&px
                %>
                <b>�Ѿ�ʧЧ�ĵĹ�����б�</b> <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>��</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>��</a> 
                <%
                          Case "click"
                           adssql="select * from KS_Advertise where act<>2 order by click "&px
                %>
                <b>���������<%if px="desc" then: Response.Write "����":else:Response.Write "����":end if%>����δʧЧ�����</b>  <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>��</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>��</a>
               <%
                          Case "show"
                           adssql="select * from KS_Advertise where act<>2 order by show "&px
                %>
                <b>����ʾ����<%if px="desc" then: Response.Write "����":else:Response.Write "����":end if%>����δʧЧ�����</b> <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>��</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>��</a>
               <%
                          Case "place"
                          
                          if isnumeric(request("place"))=true then
                           adssql="select * from KS_Advertise where act=1 and place="&trim(request("place"))&" order by regtime "&px
						 
		%>
                <b>IDΪ<%=request("place")%>�Ĺ��λ���������ŵĹ����</b>  <a href=?Action=Adslist&type=<%=KS.G("type")%>&place=<%=request("place")%>&px=x>��</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>&place=<%=request("place")%>>��</a>
				 
                <%else
                  adssql="select * from KS_Advertise where act=1 order by regtime "&px
                %>
                <b>�����������ŵĹ�����б�</b>  <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>��</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>��</a>
                        
                <%end if%>
               <%
                          Case "search"
                          if request("adorder")="id" and isnumeric(request("nr"))=true then
                           adssql="select * from KS_Advertise where id="&trim(request("nr"))
                          
                %>
                <b>��ѯ IDΪ<%=request("nr")%> �Ĺ������Ϣ</b>
                <%        else
                  adssql="select * from KS_Advertise where sitename like '%"&request("nr")&"%' order by regtime "&px
                %>
                <b>��ѯ���ƺ��йؼ��֡�<%=request("nr")%>�������</b> <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>��</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>��</a>
                        
                <%end if%>

                <%       
                          Case else
                          adssql="select * from KS_Advertise where act=1 order by regtime "&px
                %>
                <b>�����������ŵĹ�����б�</b>  <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>��</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>��</a>
                <%
                    end Select
                %>
              </td>
            </tr>
          </table>
		   </body>
</html>
<%

if isnumeric(request("page")) then
currentPage=cint(request("page"))
else
currentPage=1
end if
set RSObj=server.createobject("adodb.recordset")

RSObj.open adssql,Conn,1,1
if RSObj.eof and RSObj.bof then
Response.Write "<tr><td bgcolor=#ffffff align=center><BR><BR>û���κ���ؼ�¼<BR><BR><BR><BR>"
else
RSObj.pagesize=10  'ÿҳ��ʾ�ļ�¼��
totalPut=RSObj.recordcount '��¼����
totalPage=RSObj.pagecount
MaxPerPage=RSObj.pagesize
if currentpage<1 then
currentpage=1
end if
if currentpage>totalPage then
currentpage=totalPage
end if
if currentPage=1 then
showContent
showpages
else
if (currentPage-1)*MaxPerPage<totalPut then
RSObj.move  (currentPage-1)*MaxPerPage
dim bookmark
bookmark=RSObj.bookmark '�ƶ�����ʼ��ʾ�ļ�¼λ��
showContent
showpages
end if
end if
RSObj.close:set RSObj=nothing
end if
Conn.close:set Conn=nothing
End Sub

sub showContent
i=0
do while not (RSObj.eof or err)
%>
  <div align="center">
    <center>
    <table border="0" cellpadding="0" cellspacing="1"  width="100%" class="ctable" style="margin-bottom:3px">
		  <input type="hidden" name="id" value="<%=RSObj("id")%>">
     <tr>
        <td width="175" class="clefttitle"><font color="#FF0000">&nbsp;�����ID��<%=RSObj("id")%> </font></td>
        <td width="370" class="clefttitle">&nbsp;���ƣ�<%=RSObj("sitename")%></td>
        <td class="clefttitle" width="275">
       &nbsp;URL�� 
       <%=RSObj("url")%></td>
        <td  width="105" align="center" class="clefttitle">
        <%if RSObj("xslei")="txt" then%>
           <a href=KS.Ads.asp?Action=Manage&id=<%=RSObj("id")%>&job=yulan>Ԥ�����</a>
        <%else
        
        %>
            <a href=KS.Ads.asp?Action=Manage&id=<%=RSObj("id")%>&job=yulan>Ԥ�����</a>
       <%end if%>
��</td>
      </tr>
      <tr>
        <td width="175" height="60" class="tdbg">&nbsp;�򿪣�<%= Ggdklx(RSObj("window"))%><br>&nbsp;��ʾ��<%= Ggxslx(RSObj("xslei"))%><br>
        &nbsp;���ͣ�<%= Ggwlx(RSObj("place"))%></td>
        <td height="60" class="tdbg">&nbsp;����ʱ�䣺<font color=red><%=RSObj("regtime")%></font><br>&nbsp;
		<span style="display:None">������ʾ��<font color=red><%=RSObj("time")%></font></span><br>
        &nbsp;���µ����<font color=red><%=RSObj("lasttime")%></font></td>
        <td height="60" width="272"class="tdbg" >&nbsp;���������<%call  Djcs()%><br>
        &nbsp;�� �� λ��<%= Ggwm(RSObj("place"))%>  ID=<font color=red><%=RSObj("place")%></font></td>
        <td height="60" width="104" align="center" class="tdbg">              <%
if RSObj("act")=1 then
%>                <a href=?Action=Addads&job=edit&id=<%=RSObj("id")%>>�޸�</a>
              <a href=?Action=Manage&id=<%=RSObj("id")%>&job=close>��ͣ</a> 
              <%
else
%>
              <a href=?Action=Manage&id=<%=RSObj("id")%>&job=open>����</a> 
              <%end if%><a href=?Action=Manage&id=<%=RSObj("id")%>&job=delit>ɾ��</a></td>
      </tr>
      <tr>
        <td colspan="3" height="20" style="color:green">�Ƿ����õ������¼��<%if rsobj("clicks")="1" then response.write "����" else response.write "������"%>&nbsp;���Ž�ֹʱ�䣺<%
		if RSObj("class")="0" then%>������
<%else%>
<%=RSObj("lasttime")%>
<%
end if
%></td>
                <td height="20" width="104" align="center"></td>
      </tr>
      </table>
    </center>
</div>
  <%
i=i+1
if i>=MaxPerPage then exit do 'ѭ��ʱ�����β�������˳��������¼�ﵽҳ�����ʾ����Ҳ�˳�
RSObj.movenext
loop
end sub 

sub Showpages()
%>
    
        <table border=0 width=100% cellpadding=2>
            <tr bgcolor=#ffffff> 
              <td align=right colspan=4>
			   <%'��ʾ��ҳ��Ϣ
			  Call KS.ShowPageParamter(totalPut, MaxPerPage, "KS.Ads.asp", True, "��", CurrentPage, KS.QueryParam("page"))
			  %>
              </td>
            </tr>
        </table>
     
<%
end sub



Sub Djcs()%>
 <font color=red><%=RSObj("click")%></font> (<a href=?Action=Listip&id=<%=RSObj("id")%>&ip=cip>�����¼</a>)
<%end sub
	'�����ʾ������
	Function Ggxslx(lx)
	Select Case lx
		  Case "txt":Ggxslx="���ı�"
		  Case "gif":Ggxslx="GIFͼƬ"
		  Case "swf":Ggxslx="Flash����"
		  Case "dai":Ggxslx="Ƕ�����"
	End select
	End Function
	'����������
	Function Ggdklx(lx)
	Select Case lx
		  Case 0:Ggdklx="�´���"
		  Case else:Ggdklx="������"
	End select
	End Function
	'���λ���ͱ�ʾ���ֵ���
	Function Ggwlxsz(place1)
	set PRSObj=server.createobject("adodb.recordset")
	PRSObj.open "select top 1 * From KS_ADPlace where place="&place1,Conn,1,1
	if not PRSObj.eof then
	Ggwlxsz=PRSObj(2)
	else
	Ggwlxsz=0
	end if
	PRSObj.close
	Set PRSObj=nothing
	End Function
	'���λ�������Ƶ���
	Function Ggwlx(place)
	Dim  PRSObj:Set PRSObj=server.createobject("adodb.recordset")
	PRSObj.open "select top 1 * From KS_ADPlace where place="&place,Conn,1,1
	if not PRSObj.eof then
	Ggwlx=PRSObj(2)
	Select Case Ggwlx
		   Case 1:Ggwlx="�����ʾһ��"
		   Case 2:Ggwlx="������������"
		   Case 3:Ggwlx="������������"
		   Case 4:Ggwlx="���Ϲ�������"
		   Case 5:Ggwlx="�����������"
		   Case 6:Ggwlx="�����������"
		   Case 7:Ggwlx="�����������"
	End select
	else
	Ggwlx="���λ��ɾ��"
	end if
	PRSObj.close
	Set PRSObj=nothing
	
	End Function
	'���λ���Ƶ���
	Function Ggwm(place)
	Dim  PRSObj:Set PRSObj=server.createobject("adodb.recordset")
	PRSObj.open "select * From KS_ADPlace where place="&place,Conn,1,1
	if not PRSObj.eof then
	Ggwm=PRSObj(1)
	else
	Ggwm=""
	end if
	PRSObj.close:Set PRSObj=nothing
	End Function
	
	'��ʾIP
	Sub AdsListIP()
	    Dim getadid
	   %>
	    <html>
		<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
		<meta http-equiv=Content-Type content="text/html; charset=gb2312">

		<table border="0" align=center cellpadding="1" cellspacing="1" width="100%">
		<tr><td align=center class="sort">
		<%
		if KS.G("ip")="sip" then
		%>
		IDΪ <%=KS.G("id")%> �Ĺ������ʾ��¼
		<%
		elseif KS.G("ip")="cip" then
		%>
		IDΪ <%=KS.G("id")%> �Ĺ���������¼
		<%
		end if
		%>
		</td>
		<td class="sort" align="right"><input class="button" type="button" name="button1" value="������е�IP��¼" onClick="if (confirm('�˲���������,ȷ��ɾ�����м�¼��')){location.href='?action=IPDel&AdID=<%=KS.G("ID")%>&ip=<%=KS.G("ip")%>';}"></td>
		</tr></table>
		
		
		<table border="0" align=center cellpadding="1" cellspacing="1" class="ctable" width="100%">
		<tr><td align="center" class="sort" height="20">
		��¼ID
		</td><td align=center class="sort" height="20">IP ��ַ</td>
		  <td align=center class="sort" height="20">ʱ����</td></tr>
		<%
		if not isempty(request("page")) then
		 currentPage=cint(request("page"))
		else
		 currentPage=1
		end if
		set adsrs=server.createobject("adodb.recordset")
		
		if KS.G("ip")="sip" then
		getadid=cint(request("id"))
		adssql="select * From KS_Adiplist where adid="&getadid&" and class=1 order by id desc"
		
		elseif KS.G("ip")="cip" then
		getadid=cint(request("id"))
		adssql="select * From KS_Adiplist where adid="&getadid&" and class=2 order by id desc"
		end if
		
		adsrs.open adssql,Conn,1,1
		if adsrs.eof and adsrs.bof then
		Response.Write "<tr align=center><td bgcolor=#ffffff colspan=3>û�м�¼</td></tr></table>"
		else
		adsrs.pagesize=25 'ÿҳ��ʾ�ļ�¼��
		totalPut=adsrs.recordcount '��¼����
		totalPage=adsrs.pagecount
		MaxPerPage=adsrs.pagesize
		if currentpage<1 then
		currentpage=1
		end if
		if currentpage>totalPage then
		currentpage=totalPage
		end if
		if currentPage=1 then
		showIpContent
		else
		if (currentPage-1)*MaxPerPage<totalPut then
		adsrs.move  (currentPage-1)*MaxPerPage
		dim bookmark
		bookmark=adsrs.bookmark '�ƶ�����ʼ��ʾ�ļ�¼λ��
		showIpContent
		end if
		end if
		adsrs.close:set adsrs=nothing
		end if
		Conn.close:set Conn=nothing
		
		End Sub
		
		sub showIpContent
		i=0
		do while not (adsrs.eof or err)
		%>
		<tr align=center><td class="tdbg"><font color=red><%=adsrs("id")%></font>��</td><td align=center class="tdbg"><%=adsrs("ip")%>��</td><td align=center class="tdbg"><%=adsrs("time")%>��</td></tr>
		<%
		i=i+1
		if i>=MaxPerPage then exit do 
		adsrs.movenext
		loop
		showippages
		end sub 
		
		sub showippages()
		dim n
		n=totalPage
		%>
		</table>
		
		<table border="0" align=center cellpadding="1" cellspacing="1" width="100%">
		<tr><td align=right colspan=4 class="forumRowHighlight">
	
		<%
  Call KS.ShowPageParamter(totalPut, MaxPerPage, "KS.Ads.asp", True, "��", CurrentPage, KS.QueryParam("page"))
       %>
		
		</td></tr>
		</table>
		<%
	End Sub
	'ɾ��ip��¼
	Sub IPDel()
	 Conn.Execute("Delete From KS_Adiplist Where Adid=" & KS.ChkClng(KS.G("ADID")))
	 Response.Redirect "?Action=Listip&id=" & KS.G("adid") & "&ip=" & KS.G("IP")
	End Sub
	
	Sub AdsManage()
	    Dim ttarg
		Dim ComeUrl:ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))
		IF ComeUrl="" Then ComeUrl="Ads_List.asp"
	   %>
		<html>
		<link href="Include/admin_Style.CSS" rel="stylesheet" type="text/css">
		<div align=center>
		<center><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		  <tr><td align=center>
		<%
		dim getid,RSObj,adssql
		getid=cint(KS.G("id"))
		
		
		Select Case KS.G("job")
			case "close"
		
		   set RSObj=server.createobject("adodb.recordset")
		   adssql="Select top 1 place,id,sitename,act From KS_Advertise where id="&getid
		   RSObj.open adssql,Conn,1,3
		   RSObj("act")=0
		   RSObj.Update
		   Call CreateJS(RSOBJ("place"))
		   Call KS.Alert("�����[" & RSObj("sitename") & "]����ͣ��", ComeUrl)
		  RSObj.close
		
			case "delit"
		    Call KS.Confirm("ɾ���˹�棬���� IP ��¼Ҳ����ɾ������漰��IP��¼��ɾ�����ָܻ���", "?Action=Manage&ComeUrl1=" & Server.URLEncode(ComeUrl) &"&id=" & getid &"&job=del",ComeUrl)		
			case "del"
			conn.execute("delete from KS_UploadFiles Where ChannelID=1020 And InfoID=" & GetID)
			adssql="delete From KS_Advertise where id="&getid
			Conn.execute(adssql)
			dim adssqldelip
			adssqldelip="delete From KS_Adiplist where adid="&getid
			Conn.execute(adssqldelip)
		     Call KS.Alert("�����ɾ���ɹ���", KS.G("ComeUrl1"))
         
			case "yulan"
			set RSObj=server.createobject("adodb.recordset")
			adssql="Select id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei From KS_Advertise where id="&getid
			RSObj.open adssql,Conn,3,3
			
			RSObj("show")=RSObj("show")+1
			RSObj("time")=now()
			RSObj.Update
			if RSObj("window")=0 then
			ttarg = "_blank"
			else
			ttarg="_top"
			end if
			
			Dim GaoAndKuan
			GaoAndKuan=""
			
			if isnumeric(RSObj("hei")) then
			GaoAndKuan=" height="&RSObj("hei")&" "
			else
			
			if right(RSObj("hei"),1)="%" then
				if isnumeric(Left(rsobj("hei"),len(RSObj("hei"))-1))=true then
				 GaoAndKuan=" height="&RSObj("hei")&" "
				end if
			end if
			
		  end if
		
		
		if isnumeric(RSObj("wid")) then
		GaoAndKuan=GaoAndKuan&" width="&RSObj("wid")&" "
		else
			if right(RSObj("wid"),1)="%" then
				if isnumeric(Left(rsobj("wid"),len(RSObj("wid"))-1))=true then 
				GaoAndKuan=GaoAndKuan&" width="&RSObj("wid")&" "
				end if
			end if
		end if
		Select Case RSObj("xslei")
			
					Case "txt"%><a  title="<%=RSObj("sitename")%>"  href="url.asp?id=<%=RSObj("id")%>" target="<%=ttarg%>"><%=RSObj("intro")%></a>
		<%          Case "gif"%><a href="url.asp?id=<%=RSObj("id")%>" target="<%=ttarg%>"><img art="<%=RSObj("sitename")%>" border=0  <%=GaoAndKuan%> src="<%=RSObj("gif_url")%>"></a> 
		<%          Case "swf"%><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http:/download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0"; <%=GaoAndKuan%>><param name=movie value="<%=RSObj("gif_url")%>"><param name=quality value=high>
		
		  <embed src="<%=RSObj("gif_url")%>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash"  width="<%=RSObj("wid")%>" height="<%=RSObj("hei")%>"></embed></object>
		<%           Case "dai"%><iframe marginwidth=0 marginheight=0  frameborder=0 bordercolor=000000 scrolling=no  name="����WEB������ϵͳ zon.cn" src="daima.asp?id=<%=RSObj("id")%>"  <%=GaoAndKuan%>></iframe>
		
		  <%          Case else%><a href="url.asp?id=<%=RSObj("id")%>" target="<%=ttarg%>"><img art="<%=RSObj("sitename")%>"  border=0  <%=GaoAndKuan%> src="<%=RSObj("gif_url")%>"></a>
		<%
				   End Select
		RSObj.close

		case "yulanggw"
		%>
		<script src="../ks_inc/jquery.js"></script>
		<script language="javascript" src="<%=KS.Setting(3) & KS.Setting(93) %><%=getid%>.js"></script>
			
		<%
		case "open"
			set RSObj=server.createobject("adodb.recordset")
				adssql="Select top 1 place,id,sitename,act From KS_Advertise where id="&getid
				RSObj.open adssql,Conn,1,3
				RSObj("act")=1
				RSObj.Update
				Call CreateJS(RSOBJ("place"))
				Call KS.Alert("�����[" & RSObj("sitename") & "]�����", ComeUrl)
				RSObj.close
			
			End Select
			set RSObj=nothing 
			Conn.close:set Conn=nothing
		%>
		</td></tr><tr height=10 align=center>
		  <td><a href="javascript:this.history.go(-1)">����</a></td>
		</tr></table>
		</center></div>
<%	End Sub
End Class
%> 
