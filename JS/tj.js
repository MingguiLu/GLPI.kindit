function randomOrder1 (targetArray)
{
    var arrayLength = targetArray.length
    var tempArray1 = new Array();
    for (var i = 0; i < arrayLength; i ++)
    {
        tempArray1 [i] = i
    }
    var tempArray2 = new Array();
    for (var i = 0; i < arrayLength; i ++)
    {
        tempArray2 [i] = tempArray1.splice (Math.floor (Math.random () * tempArray1.length) , 1)
    }

    var tempArray3 = new Array();
    for (var i = 0; i < arrayLength; i ++)
    {
        tempArray3 [i] = targetArray [tempArray2 [i]]
    }
    return tempArray3
}
tmp=new Array(); 
tmp[0]='<li><img class="app-image" src="/images/ggao/dijiu.gif" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down" target="_blank">极速装机;第九论坛系列</a></span><span class="app-desc">极速、稳定方便、值得推荐</span></li>' 
tmp[1]='<li><img class="app-image" src="/images/ggao/yj.gif" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down " target="_blank">赢家竞技(斗地主、麻将)</a></span><span class="app-desc">每天送百元话费,免费下载!</span></li>' 
tmp[2]='<li><img class="app-image" src="/images/ggao/keepc.gif" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down" target="_blank">KC免费网络电话</a></span><span class="app-desc">电脑、手机使用，通话清晰</span></li>'
tmp[3]='<li><img class="app-image" src="/images/ggao/yqmis.jpg" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down" target="_blank">【8-80岁 人人会编程】</a></span><span class="app-desc"><雅奇MIS>图示化自动编程系统 </span></li>'
tmp[4]='<li><img class="app-image" src="/images/ggao/81box.gif" alt="" width="32" height="32" /><span class="app-name"><a href=" http://www.kesion.com/down" target="_blank">八音盒 SNS交友音乐盒</a></span><span class="app-desc">完美音质 小而强的在线音乐盒</span></li>'
tmp[5]='<li><img class="app-image" src="/images/ggao/gua38.gif" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down" target="_blank">呱呱视频社区2010Beta1</a></span><span class="app-desc">更大屏更清晰视频不可挡 </span></li>'
tmp[6]='<li><img class="app-image" src="/images/ggao/zhcall.gif" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down" target="_blank">中华通网络电话</a></span><span class="app-desc">长途轻松打，短信轻松发！</span></li>'
tmp[7]='<li><img class="app-image" src="/images/ggao/uucall.gif" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down" target="_blank">UUCall免费网络电话</a></span><span class="app-desc">60分钟电话免费打</span></li>'
tmp[8]='<li><img class="app-image" src="/images/ggao/hz.gif" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down" target="_blank">宏正中小企业管理软件</a></span><span class="app-desc">进销存 仓库 软件免费下载</span></li>'
tmp[9]='<li><img class="app-image" src="/images/ggao/97shendu.gif" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down" target="_blank">免激活Windows7系统</a></span><span class="app-desc">最新Windows7旗舰中文激活版</span></li>'

var lastad1 = new Array(10); 
lastad1 = this.randomOrder1(tmp); 
document.writeln("        	<div class=\"cp recommend\">");
document.writeln("                <div class=\"cp-top\">");
document.writeln("                <\/div>");
document.writeln("                <div class=\"cp-main\">");
document.writeln("                <ul>");
for(ids1 = 0; ids1 < 10; ids1++){ 

document.writeln(lastad1[ids1]); 

}
document.writeln("                <\/ul>");
document.writeln("                <\/div>");
document.writeln("        	<\/div>");
