<%
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1

Dim PhotoUrl:PhotoUrl=Request("photourl")
Dim del:del=request("del")
if PhotoUrl="" Then
  response.write "<script>alert('��û��ѡ��ͼƬ!');window.close();</script>"
end if
if left(photourl,1)<>"/" and left(lcase(photourl),4)<>"http" then photourl="/" & PhotoUrl
if request("action")="main" then
 call main
else
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>����ͼƬ�ü�</title>
<META HTTP-EQUIV="pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache, must-revalidate">
<META HTTP-EQUIV="expires" CONTENT="Wed, 26 Feb 1997 08:21:57 GMT">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<iframe src="imgcut.asp?action=main&del=<%=del%>&photourl=<%=photourl%>" scrolling="yes" width="100%" height="540"></iframe>
</body>
</html>
<%
end if
sub main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>����ͼƬ�ü�</title>
</head>
<body bgcolor="#000000">
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

#bgDiv{ min-height:400px;border:3px solid #000; position:relative;}
#dragDiv{border:1px dashed #fff; width:150px; height:120px; top:50px; left:50px; cursor:move; }
</style>
<table border="0" width="99%" bgcolor="#666666" align="center" cellspacing="0" cellpadding="0">
  <tr>
    <td style="padding:10px">
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
	<div id="viewDiv" style="border:3px solid #000;width:200px; min-height:120px;"> </div>
	   </td>
	  </tr>
	  <tr>
	   <td style="height:40px;color:#ff6600;font-size:12px">
	    <form name="myform" id="myform" action="" method="post">
		 <%if del="1" then%>
		  <label><input type="checkbox" name="del" value="1">ɾ��ԭͼ</label>
	  <br/> <br/>
		  ���ͼƬ�������ط����õ�,�벻Ҫ��ѡɾ��ԭͼ��
		  <br/>
		 <%end if%>
		  <br/>
	       <input name="" type="button" value="����ͼƬ" onclick="Create()" />
    <input name="" type="button" value="����ʹ��ԭͼ" onclick="top.close()"/>
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
 w=$("#si").width();
 h=$("#si").height();
 if (w>620) w=620;
 if (h>600) h=600;
	  ic = new ImgCropper("bgDiv", "dragDiv", "<%=PhotoUrl%>", {
		Width:w, Height: h, Color: "#999999",
		Resize: true,
		Right: "rRight", Left: "rLeft", Up:	"rUp", Down: "rDown",
		RightDown: "rRightDown", LeftDown: "rLeftDown", RightUp: "rRightUp", LeftUp: "rLeftUp",
		Preview: "viewDiv", viewWidth: 200, viewHeight: 200
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
	$("#myform").attr("action","ImgCutSave.asp?p=" + p + "&x=" + x + "&y=" + y + "&w=" + w + "&h=" + h + "&pw=" + pw + "&ph=" + ph + "&" + Math.random());
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
