<!--#include file="../conn.asp"-->
<!--#include file="../ks_cls/kesion.commoncls.asp"-->
<%
dim MapKey,MapCenterPoint
dim ks:set ks=new publiccls
mapKey=KS.Setting(175)
MapCenterPoint=KS.Setting(176)
If KS.IsNul(MapCenterPoint) Then MapCenterPoint="116.324439,39.961233"
set ks=nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>���ӵ�ͼ��ע</title>
<style type="text/css">
 body{font-size:12px;}
</style>
<script src="../ks_inc/jquery.js"></script>
<script type="text/javascript" src="http://api.map.baidu.com/api?key=<%=MapKey%>&v=1.0&services=true" ></script>
<script type="text/javascript">  

var marker=null;
var map=null;
var infoWindow=null;
function KesionDotccMap(){
	map = new BMap.Map("KesionMap");          // ������ͼʵ��  
	var point = new BMap.Point(<%=MapCenterPoint%>);  // ����������  
	
	map.centerAndZoom(point, 16);                 // ��ʼ����ͼ���������ĵ�����͵�ͼ����  
	
	//������ſؼ�
	map.addControl(new BMap.NavigationControl());  
	map.addControl(new BMap.ScaleControl());  
	map.addControl(new BMap.OverviewMapControl()); 

    showMark();
	
	map.addEventListener("click", function(e){ 
	map.removeOverlay(marker);   
 	var point = new BMap.Point(e.point.lng, e.point.lat);   
	map.centerAndZoom(point, 16);  
	
	marker = new BMap.Marker(point);         // ������ע   
	map.addOverlay(marker);                     // ����ע��ӵ���ͼ��  
	<%if request("action")<>"getcenter" then%>
    var sContent ="<div style='text-align:center;font-size:12px;margin:0 0 5px 0;padding:0.2em 0'>��ǰ����:<font color=#ff6600>"+e.point.lng+","+ e.point.lat+"</font><br/>��ȷ���ڴ�λ������ע��?<br/><br/><input type='button' value='ȷ������' style='background:#f1f1f1;border:1px solid #cccccc' onclick='addMark("+e.point.lng+","+e.point.lat+",true)'/> <input type='button' value='ȷ������' style='background:#f1f1f1;border:1px solid #cccccc' onclick='addMark("+e.point.lng+","+e.point.lat+",false)'/> <input type='button' value='ȡ��' style='background:#f1f1f1;border:1px solid #cccccc' onclick='removeMark()'/></div>"
	<%else%>
    var sContent ="<div style='text-align:center;font-size:12px;margin:0 0 5px 0;padding:0.2em 0'>��ǰ����:<font color=#ff6600>"+e.point.lng+","+ e.point.lat+"</font><br/>��ȷ�����ô���Ϊ���ı�ע��?<br/><br/><input type='button' value='ȷ������' style='background:#f1f1f1;border:1px solid #cccccc' onclick='getCenterBack("+e.point.lng+","+e.point.lat+")'/> <input type='button' value='ȡ��' style='background:#f1f1f1;border:1px solid #cccccc' onclick='removeMark()'/></div>"
	<%end if%>
   infoWindow = new BMap.InfoWindow(sContent);  // ������Ϣ���ڶ���
    
	marker.addEventListener("click", function(){										
   this.openInfoWindow(infoWindow);	}); 
	map.openInfoWindow(infoWindow, map.getCenter());      // ����Ϣ���� 
	window.setTimeout(function(){map.panTo(new BMap.Point(<%=MapCenterPoint%>));}, 2000);
	<%if request("action")<>"getcenter" then%>
    document.getElementById("info").innerHTML ="��ǰ��ͼ�������꣺" +  e.point.lng + ", " + e.point.lat;  
	<%end if%>
}); 
}

function removeMark(){
	map.removeOverlay(marker);  
	infoWindow.close(); 
}
function addMark(x,y,returnflag){
  var mtext=$("#markvalue");
  if (mtext.val().split('|').length>9){
   alert('�Բ���,���ֻ�ܱ�ע10���ط�!');
   return;
  }
  if (mtext.val()=='')
  mtext.val(x+","+y);
  else
  mtext.val(mtext.val()+"|"+x+","+y);
  if (returnflag){
  setOk();
  }
  removeMark();
  showMark();
} 
function showMarkList(v){
  var varr=v.split('|');
  var str='<strong>����ӵı�ע:</strong><br/>';
  for(var i=0;i<varr.length;i++){
     str+=intToLetter(i+1)+"��"+varr[i]+" <a href=javascript:delMark('"+varr[i]+"')><font color='#ff6600'>ɾ</font></a><br/>";
  }
  $("#marklist").html(str);
}
function showMark(){
 var markv=$("#markvalue").val();
 if (markv==''||markv==null) return;
 var varr=markv.split('|');
 for (var i=0;i<varr.length;i++){
  var point = new BMap.Point(varr[i].split(',')[0],varr[i].split(',')[1]);   
   addMarker(point, i);   
 }
 showMarkList(markv);
}
function addMarker(point, index){   
  // ����ͼ�����   
  var myIcon = new BMap.Icon("http://api.map.baidu.com/img/markers.png", new BMap.Size(23, 25), {   
    offset: new BMap.Size(10, 25),                  // ָ����λλ��   
    imageOffset: new BMap.Size(0, 0 - index * 25)   // ����ͼƬƫ��   
  });   
  var marker = new BMap.Marker(point, {icon: myIcon});   
  map.addOverlay(marker);   
}  
function delMark(v){
 if (confirm('ȷ��ɾ����γ��Ϊ'+v+'�ı�ע��')){
    var str='';
	var varr=$("#markvalue").val().split('|');
	for (var i=0;i<varr.length;i++){
	   if ("'"+varr[i]+"'"!="'"+v+"'"){
	      if (str==''){ 
		    str=varr[i];
		  }else{
		    str+='|'+varr[i];
		  }
	   }
	}
	//location.reload();
	location.href="baidumap.asp?MapMark="+escape(str);
 }
} 
function intToLetter(id){
    var k = (--id)%26//26����A~Z 26��Ӣ����ĸ����.
    var str = "";
    while(Math.floor((id=id/26))!=0){
        str = String.fromCharCode(k+65)+str;//65 ����'A'��ASCIIֵ.
        k=(--id)%26;
    }
    //String.fromCharCode(num):���num��ֵ��Ӧ����ĸ.numӦ��ΪASCII�е�ֵ.
    str = String.fromCharCode(k+65)+str;
    return str;
}
function setOk(){
  if ($("#markvalue").val()==''){
    alert('�Բ��𣬻�û������κα�ע��������ͼ����ӣ�');
	return;
  }
  try{
  parent.document.getElementById("MapMark").value=$("#markvalue").val();
  parent.closeWindow();
  }catch(e){
  }
  
}
</script> 
</head>
<body onload="KesionDotccMap();" onkeydown="if(event.keyCode==13)KesionDotccMap()">
<%if request("action")<>"getcenter" then%>
<div style="width:540px;height:420px;border:1px solid gray; float:left" id="KesionMap"></div>
<div style="margin-top:10px; margin-left:10px; float:left">
	<div style="margin-top:10px; margin-left:3px;"><strong>ʹ�÷�����</strong><br/>�϶���Ҫ�鿴�ص㲢������ɱ�ע</div>
	<div id="info" style="margin-top:10px; margin-left:10px;"></div>
	<input type="hidden" name="markvalue" size=20 value="<%=Request("MapMark")%>" id="markvalue" />
	<div id="marklist" style="margin-top:10px; margin-left:10px;"></div>
	<div style="margin-top:10px;text-align:center"><input type='button' value='ȷ���������ϱ�־' onclick='setOk()' style='height:23px;background:#f1f1f1;border:1px solid #cccccc' /></div>
</div>
<%else%>
 <div style="margin:13px">
 ������<input type="text" value="������" name="keyword" id="keyword" /><input onclick="searchMap($('#keyword').val())" type="button" value="������γ����"/>
 <span id="info"></span>
<script type="text/javascript">
 function searchMap(key){
  if(key==''){alert('������ؼ��֣����������!');return;}
  var local = new BMap.LocalSearch(map, {   
	  renderOptions:{map: map}   
	});   
	local.search(key); 
	local.setSearchCompleteCallback(function(searchResult){
			var poi = searchResult.getPoi(0);
			//alert(poi.point.lng+"   "+poi.point.lat);
			document.getElementById("info").innerHTML = "<strong>" + key + "</strong>" + "���꣺" + poi.point.lng + "," + poi.point.lat +"<input type='button' onclick=getCenterBack('"+poi.point.lng+"','"+poi.point.lat+"') value='ʹ�ô�����'/>";
	});  
 }
 
//�õ���������
function getCenterBack(x,y)
{
  parent.document.getElementById("mapcenter").value=x+','+y;
  parent.closeWindow();
}

</script>
</div>
<div style="width:680px;height:360px;border:1px solid gray; float:left" id="KesionMap"></div>
<%end if%>
</body>
</html>
