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
tmp[0]='<li><img class="app-image" src="/images/ggao/dijiu.gif" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down" target="_blank">����װ��;�ھ���̳ϵ��</a></span><span class="app-desc">���١��ȶ����㡢ֵ���Ƽ�</span></li>' 
tmp[1]='<li><img class="app-image" src="/images/ggao/yj.gif" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down " target="_blank">Ӯ�Ҿ���(���������齫)</a></span><span class="app-desc">ÿ���Ͱ�Ԫ����,�������!</span></li>' 
tmp[2]='<li><img class="app-image" src="/images/ggao/keepc.gif" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down" target="_blank">KC�������绰</a></span><span class="app-desc">���ԡ��ֻ�ʹ�ã�ͨ������</span></li>'
tmp[3]='<li><img class="app-image" src="/images/ggao/yqmis.jpg" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down" target="_blank">��8-80�� ���˻��̡�</a></span><span class="app-desc"><����MIS>ͼʾ���Զ����ϵͳ </span></li>'
tmp[4]='<li><img class="app-image" src="/images/ggao/81box.gif" alt="" width="32" height="32" /><span class="app-name"><a href=" http://www.kesion.com/down" target="_blank">������ SNS�������ֺ�</a></span><span class="app-desc">�������� С��ǿ���������ֺ�</span></li>'
tmp[5]='<li><img class="app-image" src="/images/ggao/gua38.gif" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down" target="_blank">������Ƶ����2010Beta1</a></span><span class="app-desc">��������������Ƶ���ɵ� </span></li>'
tmp[6]='<li><img class="app-image" src="/images/ggao/zhcall.gif" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down" target="_blank">�л�ͨ����绰</a></span><span class="app-desc">��;���ɴ򣬶������ɷ���</span></li>'
tmp[7]='<li><img class="app-image" src="/images/ggao/uucall.gif" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down" target="_blank">UUCall�������绰</a></span><span class="app-desc">60���ӵ绰��Ѵ�</span></li>'
tmp[8]='<li><img class="app-image" src="/images/ggao/hz.gif" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down" target="_blank">������С��ҵ�������</a></span><span class="app-desc">������ �ֿ� ����������</span></li>'
tmp[9]='<li><img class="app-image" src="/images/ggao/97shendu.gif" alt="" width="32" height="32" /><span class="app-name"><a href="http://www.kesion.com/down" target="_blank">�⼤��Windows7ϵͳ</a></span><span class="app-desc">����Windows7�콢���ļ����</span></li>'

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
