<!--#Include file="../conn.asp"-->
<!--#Include file="../ks_cls/kesion.membercls.asp"-->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312" />
<style>
body{font-size:12px;font-family:verdana;margin:0;padding:0;background-color:#FAFDFF;}
a { color:#0365BF; text-decoration: none; }
a:hover { color:#f60; text-decoration: underline; }
a.addfile{background:url(../images/others/addfile.gif) no-repeat;display:block;float:left;height:20px;margin-top:-1px;position:relative;text-decoration:none;top:0pt;width:80px;cursor:pointer;}
a:hover.addfile{background:url(../images/others/addfile.gif) no-repeat;display:block;float:left;height:20px;margin-top:-1px;position:relative;text-decoration:none;top:0pt;width:80px;cursor:pointer;}
input.addfile{cursor:pointer;height:20px;left:-10px;position:absolute;top:0px;width:1px;filter:alpha(opacity=0);opacity:0;}
#upfile_input_list{font-size:12px;font-family:verdana;}
#upfile_input_msg{font-size:12px;font-family:verdana;}
</style>
<head>
<body>
<%
Dim PostRanNum
Randomize
PostRanNum = Int(900*rnd)+1000
Session("UploadCode") = Cstr(PostRanNum)
Dim ChannelID,BasicType,BoardID,KS,KSUser,Node,BSetting,LoginTF,maxonce,HasUpLoadNum,AddWaterFlag
Set KS=New PublicCls
Set KSUser=New UserCls
LoginTF=cbool(KSUser.UserLoginChecked)
ChannelID=KS.ChkClng(KS.S("ChannelID"))
AddWaterFlag=0
Select Case  ChannelID
  Case 9992  '�ʴ�
   If KS.ASetting(42)<>"1" Then
     KS.Die "&nbsp;�������ϴ���"
   ElseIf LoginTF=false or (not KS.IsNul(KS.ASetting(46)) and KS.FoundInArr(KS.ASetting(46),KSUser.GroupID,",")=false) Then
		  KS.Die "&nbsp;�Բ���,��û���ڴ�Ƶ���ϴ���Ȩ��!"
   End If
   
		 HasUpLoadNum=Conn.Execute("select count(1) From KS_UploadFiles Where ChannelID=" & ChannelID & " and datediff(" & DataPart_D & ",AddDate," & SQLNowString & ")<1 and username='" & KSUser.UserName &"'")(0)  '�������ϴ�����
		 BasicType=9992
		 maxtotal=KS.ChkClng(KS.ASetting(45))
		 maxonce=maxtotal
  Case 9994  '��̳�ϴ��ӿ�
    BoardID=KS.ChkClng(KS.S("BoardID"))
	If BoardID=0 Then
	  KS.Die "&nbsp;�Ƿ�����!"
	Else
		KS.LoadClubBoard
		 Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" &BoardID &"]")
		 If Node Is Nothing Then KS.Die "&nbsp;�Ƿ�����!"
		 BSetting=Node.SelectSingleNode("@settings").text
		 BSetting=BSetting & "$$$$$$0$$0$$0$$0$$0$$0$$0$$0$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"
		 BSetting=Split(BSetting,"$")
		 If KS.ChkClng(BSetting(36))<>1 Then
		  KS.Die "&nbsp;�˰����趨,�������ϴ�����!"
		 End If
		 If LoginTF=false or (not KS.IsNul(BSetting(17)) and KS.FoundInArr(BSetting(17),KSUser.GroupID,",")=false) Then
		  KS.Die "&nbsp;�Բ���,��û���ڴ˰����ϴ���Ȩ��!"
		 End If
		 AddWaterFlag=KS.ChkClng(BSetting(43))
		 HasUpLoadNum=Conn.Execute("select count(1) From KS_UploadFiles Where ClassID=" & BoardID & " and datediff(" & DataPart_D & ",AddDate," & SQLNowString & ")<1 and username='" & KSUser.UserName &"'")(0)  '�������ϴ�����
		 BasicType=9994
		 maxtotal=KS.ChkClng(Bsetting(39))
		 maxonce=maxtotal
	End If
 Case 9993  'д��־
	   If KS.SSetting(26)<>"1" Then
		 KS.Die "&nbsp;�������ϴ���"
	   ElseIf LoginTF=false or (not KS.IsNul(KS.SSetting(30)) and KS.FoundInArr(KS.SSetting(30),KSUser.GroupID,",")=false) Then
			  KS.Die "&nbsp;�Բ���,��û���ڴ�Ƶ���ϴ���Ȩ��!"
	   End If
   
		 HasUpLoadNum=Conn.Execute("select count(1) From KS_UploadFiles Where ChannelID=" & ChannelID & " and datediff(" & DataPart_D & ",AddDate," & SQLNowString & ")<1 and username='" & KSUser.UserName &"'")(0)  '�������ϴ�����
		 BasicType=9993
		 maxtotal=KS.ChkClng(KS.SSetting(29))
		 maxonce=maxtotal
 Case Else
 maxtotal=10:maxonce=10 : HasUpLoadNum=0
 BasicType=KS.C_S(ChannelID,6)
End Select
If maxtotal=0 Then maxonce=10
Set KS=Nothing
Set KSUser=Nothing
CloseConn

%>


<script language="javascript">
<!--
var UploadFileInput={
	$$:function(d){return document.getElementById(d);},
	isFF:function(){var a=navigator.userAgent;return a.indexOf('Gecko')!=-1&&!(a.indexOf('KHTML')>-1||a.indexOf('Konqueror')>-1||a.indexOf('AppleWebKit')>-1);},
	ae:function(o,t,h){if (o.addEventListener){o.addEventListener(t,h,false);}else if(o.attachEvent){o.attachEvent('on'+t,h);}else{try{o['on'+t]=h;}catch(e){;}}},
	count:0,
	realcount:0,
	uped:0,//�����Ѿ��ϴ�����
	max:1,//�������ϴ����ٸ�
	once:1,//�����ͬʱ�ϴ����ٸ�
	uploadcode:0,
	readme:'',
	add:function(){
		if (UploadFileInput.chkre()){
			UploadFileInput_OnEcho('<font color=red><b>���Ѿ���ӹ����ļ���!</b></font>');
		}else if(UploadFileInput.max-UploadFileInput.uped<=0 && <%=maxtotal%>!=0){
			UploadFileInput_OnEcho('<font color=red><b>�Բ������������ϴ��ļ���</b></font>');
		}
		else if (UploadFileInput.realcount>=UploadFileInput.max && <%=maxtotal%>!=0){
			UploadFileInput_OnEcho('<font color=red><b>�����ֻ���ϴ�'+UploadFileInput.max+'���ļ���</b></font>');
		}else if (UploadFileInput.realcount>=UploadFileInput.once){
			UploadFileInput_OnEcho('<font color=red><b>��һ�����ֻ���ϴ�'+UploadFileInput.once+'���ļ���</b></font>');
		}else{
			UploadFileInput_OnEcho('<font color=blue>�����<font color=red>'+(UploadFileInput.count+1)+'</font>��,���Լ�����Ӹ�����</font>');
			var o=UploadFileInput.$$('upfile_input_'+UploadFileInput.count);
			++UploadFileInput.count;
			++UploadFileInput.realcount;
			UploadFileInput_OnResize();
			var oInput=document.createElement('input');
			oInput.type='file';
			oInput.id='upfile_input_'+UploadFileInput.count;
			oInput.name='upfile_input_'+UploadFileInput.count;
			oInput.size=1;
			oInput.className='addfile';
			UploadFileInput.ae(oInput,'change',function(){UploadFileInput.add();});
			o.parentNode.appendChild(oInput);
			o.blur();
			o.style.display='none';
			UploadFileInput.show();
		}
	},
	chkre:function(){
		var c=UploadFileInput.$$('upfile_input_'+UploadFileInput.count).value;
		for (var i=UploadFileInput.count-1; i>=0; --i){
			var o=UploadFileInput.$$('upfile_input_'+i);
			if (o&&o.value==c&&UploadFileInput.realcount>0){return true}
		}
		return false;
	},
	filename:function(u){
		var p=u.lastIndexOf('\\');
		return (p==-1?u:u.substr(p+1));
	},
	show:function(){
		var oDiv=document.createElement('div');
		var oBtn=document.createElement('img');
		var i=UploadFileInput.count-1;
		oBtn.id='upfile_input_btn_'+i;
        oBtn.src='../images/default/filedel.gif';
        oBtn.alt='ɾ��';
		oBtn.style.cursor='pointer';
		var o=UploadFileInput.$$('upfile_input_'+i);
		UploadFileInput.ae(oBtn,'click',function(){
			UploadFileInput.remove(i);
        });
		if (o.value.length>70){
        oDiv.innerHTML=' <font color=gray>'+o.value.substr(0,70)+'...</font> ';
		}else{
        oDiv.innerHTML=' <font color=gray>'+o.value+'</font> ';
		}
        oDiv.appendChild(oBtn);
        UploadFileInput.$$('upfile_input_show').appendChild(oDiv);
	},
	remove:function(i){
		var oa=UploadFileInput.$$('upfile_input_'+i);
		var ob=UploadFileInput.$$('upfile_input_btn_'+i);
		if(oa&&i>0){oa.parentNode.removeChild(oa);}
		if(ob){ob.parentNode.parentNode.removeChild(ob.parentNode);}
		if(0==i){UploadFileInput.$$('upfile_input_0').disabled=true;}
		if(0==UploadFileInput.realcount){UploadFileInput.clear();}else{--UploadFileInput.realcount;}
		UploadFileInput_OnResize();
	},
	init:function(){
		var a=document;
		a.writeln('<form id="batchupfileform" name="batchupfileform" action="upfilesave.asp"  method="post" enctype="multipart/form-data" style="margin:0;padding:0;"><input name="AddWaterFlag" type="hidden" id="AddWaterFlag" value="<%=AddWaterFlag%>"><input type="hidden" id="UploadCode" name="UploadCode" value="'+UploadFileInput.uploadcode+'" /><input type="hidden" name="AutoReName" value="4"><input name="Type"" value="File" type="hidden"><input name="BasicType"" value="<%=BasicType%>" type="hidden"><input name="ChannelID" value="<%=ChannelID%>" type="hidden"><input name="BoardID" value="<%=BoardID %>" type="hidden"><div id="batchupfileformarea"><img src="../images/default/fileitem.gif" alt="���������Ӹ���" border="0" /> <a href="javascript:;">��Ӹ���<input id="upfile_input_0" name="upfile_input_0" class="addfile" size="1" type="file" onchange="UploadFileInput.add();" /></a> <span id="upfile_input_upbtn"><a href="javascript:UploadFileInput.send();">�ϴ�����</a></span> <span id="upfile_input_msg"></span> '+UploadFileInput.readme+'</div></form></div><div id="upfile_input_show"></div>');
	},
	send:function(){
		if (UploadFileInput.realcount>0){
			UploadFileInput.$$('upfile_input_'+UploadFileInput.count).disabled=true;
			UploadFileInput.$$('upfile_input_upbtn').innerHTML='�ϴ��У����Ե�..';
			UploadFileInput.$$('batchupfileform').submit();
		}else{
			alert('������Ӹ������ϴ���');
		}
	},
	clear:function(){
		for (var i=UploadFileInput.count; i>0; --i){
			UploadFileInput.remove(i);
		}
		UploadFileInput.$$('batchupfileform').reset();
		var o=UploadFileInput.$$('upfile_input_btn_0');
		if(o){o.parentNode.parentNode.removeChild(o.parentNode);}
		UploadFileInput.$$('upfile_input_0').disabled=false;
		UploadFileInput.$$('upfile_input_0').style.display='';
		UploadFileInput.count=0;
		UploadFileInput.realcount=0;
	}
}
UploadFileInput_OnResize=function(){
	var o=parent.document.getElementById("upiframe");
	(o.style||o).height=(parseInt(UploadFileInput.realcount)*16+18)+'px';
}

UploadFileInput_OnEcho=function(str){
	UploadFileInput.$$('upfile_input_msg').innerHTML=str;
}
UploadFileInput_OnMsgSuc=function(str){
	UploadFileInput_OnEcho(str);
	UploadFileInput.clear();
}
UploadFileInput_OnMsgFail=function(str){
	UploadFileInput_OnEcho(str);
	UploadFileInput.clear();
}
UploadFileInput_OnUpdateRndCode=function(str){
	UploadFileInput.$$('UploadCode').value=str;
}
//-->
</script>

<script language="javascript">
<!--
UploadFileInput.uploadcode='<%=PostRanNum%>';
UploadFileInput.uped=parseInt('<%=HasUpLoadNum%>');
UploadFileInput.max=parseInt('<%=maxtotal%>');   //��������ϴ�����
UploadFileInput.once=parseInt('<%=maxonce%>');  //һ���ϴ���������
<%If maxtotal<>0 then%>
UploadFileInput.readme='���컹���ϴ�'+(UploadFileInput.max-UploadFileInput.uped)+'�����ơ�';
<%else%>
UploadFileInput.readme='�ϴ���������';
<%end if%>
UploadFileInput.init();	
UploadFileInput_OnResize();
//-->
</script>
</body>
</html>
