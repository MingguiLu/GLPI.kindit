document.write('<marquee height=160  width=468 >');
if(0==0 || (0==1 && checkDate56('2010-9-4'))){
document.write("<a href=\"http://www.kesion.com\" onclick=\"addHits56(1,10)\" target=\"_blank\" hidefocus><button disabled style=\"cursor:pointer;border:none\"><object classid=\"clsid:D27CDB6E-AE6D-11cf-96B8-444553540000\" codebase=\"http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0\"  height=250  width=800 ><param name=\"movie\" value=\"http://flash.bdchina.com/swf/kangmeizhilian.swf\" /><param name=\"quality\" value=\"high\" /><param name=\"wmode\" value=\"transparent\" /><embed src=\"http://flash.bdchina.com/swf/kangmeizhilian.swf\" quality=\"high\" pluginspage=\"http://www.macromedia.com/go/getflashplayer\" type=\"application/x-shockwave-flash\"  height=250  width=800 ></embed></object><button></a>&nbsp;");
}
document.write("</marquee>");
function addHits56(c,id){if(c==1){try{jQuery.getScript('http://localhost/plus/ajaxs.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
function checkDate56(date_arr){
 var date=new Date();
 date_arr=date_arr.split("-");
var year=parseInt(date_arr[0]);
var month=parseInt(date_arr[1])-1;
var day=0;
if (date_arr[2].indexOf(" ")!=-1)
day=parseInt(date_arr[2].split(" ")[0]);
else
day=parseInt(date_arr[2]);
var date1=new Date(year,month,day);
if(date.valueOf()>date1.valueOf())
 return false;
else
 return true
}
