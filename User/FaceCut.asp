<!--#include file="../conn.asp"-->
<!--#include file="../ks_cls/kesion.Membercls.asp"-->
<%
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1

Dim KSUser:Set KSUser=New UserCls
dim KS:Set KS=New PublicCls
IF Cbool(KSUser.UserLoginChecked)=false Then
	KS.Die "<script>top.location.href='Login';</script>"
End If


Dim PhotoUrl:PhotoUrl=Request("photourl")
if left(lcase(photourl),4)="http" then photourl=replace(lcase(photourl),lcase(ks.getdomain),ks.setting(3))
if left(photourl,1)<>"/" and left(photourl,3)<>"../" and left(lcase(photourl),4)<>"http" then photourl="/" & PhotoUrl
if request("action")="docut" then
 docut
elseif request("action")="dosave" then
 dosave
else
 main
end if

sub docut()
Dim Pic:Pic = PhotoUrl
If KS.IsNul(Pic) Then
 KS.Die "<script>alert('��û���ϴ�ͼƬ');</script>"
ElseIf instr(lcase(pic),".gif")=0 and instr(lcase(pic),".jpg")=0 and instr(lcase(pic),".png")=0 and instr(lcase(pic),".jpeg")=0 Then
 KS.Die "<script>alert('��ͼƬ�ļ�!');</script>"
ElseIf left(lcase(pic),4)="http" and instr(lcase(pic),lcase(ks.getdomain))=0 Then
 KS.Die "<script>alert('�Ǳ�վͼƬ���ܴ���!');</script>"
End If
Dim PointX:PointX = KS.ChkClng(KS.S("x"))
Dim PointY:PointY = KS.ChkClng(KS.S("y"))
Dim CutWidth:CutWidth = KS.ChkClng(KS.S("w"))
Dim CutHeight:CutHeight = KS.ChkClng(KS.S("h"))
Dim PicWidth:PicWidth = KS.ChkClng(KS.S("pw"))
Dim PicHeight:PicHeight = KS.ChkClng(KS.S("ph"))

on error resume next
Set Jpeg = Server.CreateObject("Persits.Jpeg")
if err then 
 err.clear
 KS.Die "<script>alert('��������֧��aspJpeg���!');</script>"
end if
Jpeg.Open Server.MapPath(Pic)

'�����и�ͼƬ
Jpeg.Width = PicWidth
Jpeg.Height = PicHeight
Jpeg.Crop PointX, PointY, CutWidth + PointX, CutHeight + PointY

Dim filename:filename=KSUser.GetUserInfo("userid") & ".jpg"


Dim SaveName
SaveName=KS.ReturnChannelUserUpFilesDir(9999,KSUser.UserName) &  filename


Jpeg.Save Server.MapPath(SaveName)        '����ͼƬ������

Conn.Execute("Update KS_User Set UserFace='" & SaveName & "' where username='" &KSUser.UserName &"'")
Conn.Execute("Update KS_GuestComment Set UserFace='" & SaveName & "' where username='" &KSUser.UserName &"'")
KS.Die "<script>alert('��ϲ�����ĸ���������Ƭ�Ѹ���!');top.location.reload();top.location.href='User_EditInfo.asp?Action=face';</script>"
 
'���ͼƬ
'Response.ContentType = "image/jpeg"
'Jpeg.SendBinary

Set KS=Nothing
end sub

sub dosave()
Conn.Execute("Update KS_GuestComment Set UserFace='" & SaveName & "' where username='" &KSUser.UserName &"'")
Conn.Execute("Update KS_User Set UserFace='" & Replace(PhotoUrl,"../","/") & "' where username='" &KSUser.UserName &"'")
KS.Die "<script>alert('��ϲ�����ĸ���������Ƭ�Ѹ���!');top.location.reload();</script>"
end sub

sub main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<META HTTP-EQUIV="pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache, must-revalidate">
<META HTTP-EQUIV="expires" CONTENT="Wed, 26 Feb 1997 08:21:57 GMT">
<title>����ͼƬ�ü�</title>
</head>
<body>
<script type="text/javascript">
var isIE = (document.all) ? true : false;

var isIE6 = isIE && ([/MSIE (\d)\.0/i.exec(navigator.userAgent)][0][1] == 6);

var $$ = function (id) {
	return "string" == typeof id ? document.getElementById(id) : id;
};

var Class = {
	create: function() {
		return function() { this.initialize.apply(this, arguments); }
	}
}

var Extend = function(destination, source) {
	for (var property in source) {
		destination[property] = source[property];
	}
}

var Bind = function(object, fun) {
	return function() {
		return fun.apply(object, arguments);
	}
}

var BindAsEventListener = function(object, fun) {
	var args = Array.prototype.slice.call(arguments).slice(2);
	return function(event) {
		return fun.apply(object, [event || window.event].concat(args));
	}
}

var CurrentStyle = function(element){
	return element.currentStyle || document.defaultView.getComputedStyle(element, null);
}

function addEventHandler(oTarget, sEventType, fnHandler) {
	if (oTarget.addEventListener) {
		oTarget.addEventListener(sEventType, fnHandler, false);
	} else if (oTarget.attachEvent) {
		oTarget.attachEvent("on" + sEventType, fnHandler);
	} else {
		oTarget["on" + sEventType] = fnHandler;
	}
};

function removeEventHandler(oTarget, sEventType, fnHandler) {
    if (oTarget.removeEventListener) {
        oTarget.removeEventListener(sEventType, fnHandler, false);
    } else if (oTarget.detachEvent) {
        oTarget.detachEvent("on" + sEventType, fnHandler);
    } else { 
        oTarget["on" + sEventType] = null;
    }
};
</script>
<script type="text/javascript" src="../ks_inc/imgplus/ImgCropper.js"></script>
<script type="text/javascript" src="../ks_inc/imgplus/Drag.js"></script>
<script type="text/javascript" src="../ks_inc/imgplus/Resize.js"></script>
<script src="../ks_inc/jquery.js"></script>
<style type="text/css">
#rRightDown,#rLeftDown,#rLeftUp,#rRightUp,#rRight,#rLeft,#rUp,#rDown{
	position:absolute;
	background:#FFF;
	border: 1px solid #333;
	width: 6px;
	height: 6px;
	z-index:500;
	font-size:0;
	opacity: 0.5;
	filter:alpha(opacity=50);
}
.button{border-color:#3366cc;margin-right:1em;color:#fff;background:#3366cc;}
.button{border-width:1px;cursor:pointer;padding:.1em 1em;*padding:0 1em;font-size:9pt; line-height:130%; overflow:visible;}

#rLeftDown,#rRightUp{cursor:ne-resize;}
#rRightDown,#rLeftUp{cursor:nw-resize;}
#rRight,#rLeft{cursor:e-resize;}
#rUp,#rDown{cursor:n-resize;}

#rLeftDown{left:0px;bottom:0px;}
#rRightUp{right:0px;top:0px;}
#rRightDown{right:0px;bottom:0px;background-color:#00F;}
#rLeftUp{left:0px;top:0px;}
#rRight{right:0px;top:50%;margin-top:-4px;}
#rLeft{left:0px;top:50%;margin-top:-4px;}
#rUp{top:0px;left:50%;margin-left:-4px;}
#rDown{bottom:0px;left:50%;margin-left:-4px;}

#bgDiv{ border:3px solid #000;position:relative;}
#dragDiv{border:1px dashed #fff; width:133px; height:134px; top:50px; left:50px; cursor:move; }
</style>
<table border="0" width="99%" align="center" cellspacing="0" cellpadding="0">
  <tr>
    <td style="padding:10px;">
	 <div id="bgDiv">
        <div id="dragDiv">
          <div id="rRightDown"> </div>
          <div id="rLeftDown"> </div>
          <div id="rRightUp"> </div>
          <div id="rLeftUp"> </div>
          <div id="rRight"> </div>
          <div id="rLeft"> </div>
          <div id="rUp"> </div>
          <div id="rDown"></div>
        </div>
      </div></td>
    <td valign="top" align="left">
	 <br/><br/>
	 <table border="0">
	  <tr>
	   <td>
	<div style="text-align:left;font-weight:bold;maring:2px">Ч��Ԥ��:</div>
	   </td>
	  </tr>
	  <tr>
	   <td style="height:120px">
	<div id="viewDiv" style="border:3px solid #000;width:133px; height:134px;"> </div>
	   </td>
	  </tr>
	  <tr>
	   <td style="height:40px;color:#ff6600;font-size:12px">
	    <form name="myform" id="myform" action="" method="post">
           <br/><br/>
	       <input name="" type="button" class="button" value="����ü����ͷ��" onClick="Create()" /><br/><br/>
           <input name="" type="button" class="button" value="���ü�ԭͼ����" onClick="DoSave()"/>
        </form>
	   </td>
	  </tr>
	  </table>
	</td>
  </tr>
</table>
<br />
<br />

<Img id="si" src="<%=PhotoUrl%>" style="display:none"/>
<img id="imgCreat" style="display:none;" />

<script>
var h,w,ic;
$(document).ready(function(){
<%if session("urel")="" then
  session("urel")="true"
 %>
 top.location.reload();
<%end if%>
 w=$("#si").width();
 h=$("#si").height();
 if (w>540) w=540;
// if (h>600) h=600;
	  ic = new ImgCropper("bgDiv", "dragDiv", "<%=PhotoUrl%>", {
		Width:w, Height: h, Color: "#999999",
		Resize: true,
		Right: "rRight", Left: "rLeft", Up:	"rUp", Down: "rDown",
		RightDown: "rRightDown", LeftDown: "rLeftDown", RightUp: "rRightUp", LeftUp: "rLeftUp",
		Preview: "viewDiv", viewWidth: 133, viewHeight: 134
	})
});

function Create(){
	var p = ic.Url, o = ic.GetPos();
	x = o.Left,
	y = o.Top,
	w = o.Width,
	h = o.Height,
	pw = ic._layBase.width,
	ph = ic._layBase.height;
	$("#myform").attr("action","FaceCut.asp?action=docut&photourl=" + p + "&x=" + x + "&y=" + y + "&w=" + w + "&h=" + h + "&pw=" + pw + "&ph=" + ph + "&" + Math.random());
	$("#myform").submit();
}
function DoSave(){
	$("#myform").attr("action","FaceCut.asp?action=dosave&photourl=<%=PhotoUrl%>&" + Math.random());
	$("#myform").submit();
}

$(function(){
 $("#bgDiv").width($("#bgDiv").find("img").width());
 $("#bgDiv").height($("#bgDiv").find("img").height());
 
});
</script>

</body>
</html>
<%end sub%>
