//��ʾ���ӻ��ֺ�����,����:bgdir ����ͼ·��,tipstr ��ʾ��Ϣ,�ö��ŷֿ�
function popShowMessage(bgdir,tipstr){
	if(document.readyState=="complete"){ 
	if (tipstr==null || tipstr=='')return;
	var p=new KesionPopup()
	p.BgColor="url("+bgdir+"/images/popupbg.gif) no-repeat";
	p.QuoteByAdmin=false;
	p.MsgBorder=0;
	p.ShowBackground=false;
	p.ShowClose=false;
	p.popup('','<div id="tipsmessage" style="color:#fff;margin-top:12px;height:45px;text-align:center">...</div>',244);
	//$("#mesWindowContent").hide();
	//$("#mesWindowContent").slideToggle('fast');
	showtips(0,tipstr);
	}else{ 
  setTimeout(function(){popShowMessage(bgdir,tipstr);},10); 
  }
}
function showtips(n,tipstr){
  var tipsarr=tipstr.split(',')
  $("#tipsmessage").html(tipsarr[n]);
  n++;
  if (n>tipsarr.length) {
  $("#mesWindowContent").slideToggle('fast');
  //closeWindow();
  return;
  }
  setTimeout(function () { showtips(n,tipstr); }, n==1?2200:1000);
}
//ͶƱ
function doVote(dir,voteid,votetype){
 var VoteOption='';
 if (votetype=='Single'){
	VoteOption=$("input[@name=VoteOption][checked]").val(); 
 }else{
	 $("input[name=VoteOption]").each(function(){
					if ($(this).attr("checked")==true){
						if (VoteOption==''){
							 VoteOption=$(this).val();
						}else{
							 VoteOption+=","+$(this).val();
						}
					}
	});
 }
 if (VoteOption==undefined||VoteOption==''){
	 alert('��ѡ����ͶƱ��!');
	 return false;
 }
 	 $.get("../"+dir+"/ajax.asp",{action:"dovote",voteid:voteid,VoteOption:VoteOption},function(r){
		  var rstr=unescape(r);
		  if (rstr.substring(0,7)=='success'){
			   $("#showvote").html(rstr.split('@@@')[1]);
		  }else{
			   alert(rstr);
		  }
	});

}
function showVoteUser(dir,voteid){
	var p=new KesionPopup()
	p.popupIframe('�鿴ͶƱ����',"../"+dir+"/showvoteuser.asp?voteid="+voteid,330,330,"auto")
}
function movetopic(dir,topicid,title){
		  new KesionPopup().popup("�����ƶ�","<form name='moveform' action='../"+dir+"/ajax.asp' method='get'><img src='../"+dir+"/images/p_up.gif'><b>�ƶ����ӣ�</b>"+title+"<br/><br/><b>�Ƶ����棺</b><span id='showboardselect'></span><div style='text-align:center;margin:20px'><input type='submit' value='ȷ���ƶ�' class='btn'><input type='hidden' value="+topicid+" name='id' id='id'><input type='hidden' value='movetopic' name='action'><input type='button' value=' ȡ �� ' onclick='closeWindow()' class='btn'></div></form>",500);
		  $.get("../plus/ajaxs.asp",{action:"GetClubBoardOption"},function(r){
		    $("#showboardselect").html(unescape(r));
	});
}
 //����
function Posted(){
	var p=new KesionPopup();
	p.MsgBorder=1;
	p.ShowBackground=false;
	p.BgColor='#fff';
	popTopHeight=200;
	p.TitleCss="font-size:14px;background:#1B76B7;color:#fff;height:22px;";
    var tips='<div style="background:url(../user/images/loginbg.png) repeat-x;padding:5px;">���浼��<span id="navlist1"></span><span id="navlist2"></span><br/><div id="boardlist"><img src="../images/loading.gif" /></div></div>';
    p.popup('<img src="../user/images/icon11.png" align="absmiddle"> ����ѡ����淢��',tips,600);
   $.get("../plus/ajaxs.asp",{action:"getclubboard",anticache:Math.floor(Math.random()*1000)},function(d){
    $("#boardlist").html(d);
   });
}
function checklength(cobj,cmax)
{   
    var star='';
    if (PresetPoint!=null){
	 for(var k=0;k<PresetPoint.length;k++){
		 if ($("#star"+k).val()!=''){
			 if (star==''){
				  star=$("#star"+k).val();
			 }else{
				  star+=' '+$("#star"+k).val();
			 }
		 }
	 }
	}
	if (cmax-cobj.value.length-star.length<0) {
	 cobj.value = cobj.value.substring(0,cmax);
	 alert("�����������ܳ���"+cmax+"���ַ�!");
	}
	else {
	 $('#cmax').html(cmax-cobj.value.length-star.length);
	}
}
//����
function comments(dir,topicid,replayid,boardid,n,userId){
	var p=new KesionPopup();
	p.MsgBorder=1;
	p.ShowBackground=false;
	p.BgColor='#fff';
	p.TitleCss="font-size:14px;background:#1B76B7;color:#fff;height:22px;";
    var tips='<div style="background:url(../user/images/loginbg.png) repeat-x;padding:5px;"><div id="comts"></div><textarea name="comment" onkeydown="checklength(this,255);" onkeyup="checklength(this,255);" id="comment" cols="50" rows="4" style="border:1px solid #ccc;color:#666;width:450px;height:90px"></textarea><div style="margin-top:6px;margin-bottom:10px;">����<select id="Prestige"><option value="-1">-1</option><option value="-2">-2</option><option value="0">0</option><option value="1" selected>+1</option><option value="2">+2</option></select>&nbsp;&nbsp;<input type="button" onclick="saveComments(\''+dir+'\','+topicid+','+replayid+','+boardid+','+n+','+userId+')" class="btn" value="����"/> <span style="color:#999">Tips:������������<span id="cmax">255</span>���ַ�!</span></div></div>';
    p.popup('<img src="../user/images/icon11.png" align="absmiddle"> ����',tips,500);
	$.ajax({type:"get",url:"../"+dir+"/ajax.asp?action=checkcomments&userId="+userId+"&n="+n+"&topicid="+topicid+"&id="+replayid+"&boardid="+boardid+"&anticache=" + Math.floor(Math.random()*1000),cache:false,dataType:"html",success:function(d){
	    var rstr=unescape(d).split('|');
		if (rstr[0]=='success'){
			showPresetPoint(rstr[1]);
		}else{
			alert(rstr[1]);
			closeWindow();
		}
		
	}
  });
}
var startnum = 5;	//�ǵĸ���
var selectedcolor = "#ff6600";	//ѡ�ϵ���ɫ
var uselectedcolor = "#999999";//δѡ����ɫ
var PresetPoint=null;
function setstar(k,index)
{
	for(var i=1;i<=index;i++){
		$("#s"+k+i)[0].style.color=selectedcolor;
		$("#s"+k+i)[0].style.cursor="hand";
	}
	for(var i=(index+1);i<=startnum;i++){
		$("#s"+k+i)[0].style.color=uselectedcolor;
		$("#s"+k+i)[0].style.cursor="hand";
	}
}
function clickstar(presetpoint,k,index)
{   
    $("#star"+k).val(presetpoint+'��<i>'+index+'</i>');
	checklength($("#comment")[0],255);
}
function showPresetPoint(s){
 if (s=='') return;
 var str='';
 PresetPoint=s.split(',');
 for (var k=0;k<PresetPoint.length;k++){
        str+=PresetPoint[k]+':<input type="hidden" name="star'+k+'" id="star'+k+'">';
	 for(var i=1;i<=startnum;i++){
			str+=('<span id="s'+k+i+'" style="color:#999;font-size:14px;" onclick="clickstar(\''+PresetPoint[k]+'\','+k+','+i+')" title="'+i+'��" onmouseout="setstar('+k+','+i+')" onmouseover="setstar('+k+','+i+')">��</span>');
		}
		str+="&nbsp;"
 }
  $("#comts").html(str);
}
function saveComments(dir,topicid,replayid,boardid,n,userId){
	var star='';
	var c=$("#comment").val();
	if(c==''){alert('�������������!');$("#comment").focus();return;}
	if (PresetPoint!=null){
	 for(k=0;k<PresetPoint.length;k++){
		 if ($("#star"+k).val()!=''){
			 if (star==''){
				  star=$("#star"+k).val();
			 }else{
				  star+=' '+$("#star"+k).val();
			 }
		 }
	 }
	}
	if (star!='') c=star+"<br/>"+c
 	$.get("../"+dir+"/ajax.asp",{action:"comments",n:n,userId:userId,Prestige:$("#Prestige option:selected").val(),comment:escape(c),topicid:topicid,boardid:boardid,id:replayid},function(r){
	    var rstr=unescape(r).split('|');
		if (rstr[0]=="success"){
			closeWindow();
			$("#comment_"+replayid).html(rstr[1]);
		}else{
			alert(rstr[1]);
		}
	 });
}
//������ҳ��ʾ
function ShowCmtPage(dir,p,pid,boardid){
	$("#comment_"+pid).html("������...");
	$.ajax({type:"get",url:"../"+dir+"/ajax.asp?action=getcommentpage&p="+p+"&pid="+pid+"&boardid="+boardid+"&anticache=" + Math.floor(Math.random()*1000),cache:false,dataType:"html",success:function(d){
			$("#comment_"+pid).html(d);																																																									   }
 });
}
//ɾ������
function delCmt(dir,id,pid,boardid,p){
	if (confirm('ɾ���󣬲��ɻָ���ȷ��ɾ����')){
	$.ajax({type:"get",url:"../"+dir+"/ajax.asp?action=delcomment&p="+p+"&id="+id+"&boardid="+boardid+"&anticache=" + Math.floor(Math.random()*1000),cache:false,dataType:"html",success:function(d){
	  if (d=="success"){	
	    alert('��ϲ��ɾ���ɹ�!');
		ShowCmtPage(dir,p,pid,boardid);	
	  }else
	    alert(d);
	  }
 });
	}
}


function loadBoard(v){
  if (v==''||v=='0') return;
  var str=$("#pid>option:selected").text();
   $("#navlist1").html("->"+str);
   $("#navlist2").html("");
  $.get("../plus/ajaxs.asp",{action:"getclubboard",pid:v},function(d){
    $("#boardlist").html(d);
   });
}
function toBoard(){
 var bid=$('#bid>option:selected').val();
 if (bid!='' && bid!=undefined)
 location.href='?boardid='+bid;
 else
  alert('��ѡ��Ҫ������Ӱ���!');
}
function toPost(dir){
 if (dir!='') dir='../'+dir;
 var bid=$('#bid>option:selected').val();
 if (bid!='' && bid!=undefined)
 location.href=dir+'post.asp?bid='+bid;
 else
  alert('��ѡ��Ҫ���뷢���Ӱ���!');
}
function insertHiddenContent(ev){
	new KesionPopup().mousepop("������������","������Ҫ�ظ����ܲ鿴������<br /><textarea name='message' id='hidmessage' style='width:440px;height:120px'></textarea><div style='text-align:center;margin:10px'><input type='button' onclick='insertHidden();' value=' ȷ �� ' class='btn'><input type='button' value=' ȡ �� ' onclick='closeWindow()' class='btn'></div>",450);
	return false;
}

function insertHidden(){
	var str="[replyview]"+document.getElementById("hidmessage").value+"[/replyview]";
	Editor.insertText(str);
	closeWindow();
}
function popUserInfo(obj,n){
	jQuery('#user'+n).show();
	jQuery('#f'+n).html(obj.innerHTML);
}
function showPopUserInfo(n){
		jQuery('#user'+n).show();
}
function hidPopUserInfo(n){
	jQuery('#user'+n).hide();
}
var selectId='';
var bstr='';
function showmanage(c,v,dir,boardid){
	if (c){
		if (selectId==''){
		 selectId=v;
		}else{
		  var sarr=selectId.split(',');
		  var fv=false;
		  for(var i=0;i<sarr.length;i++){
			   if (sarr[i]==v){
				   fv=true;
				   break;
			   }
		  }
		  if (fv==false){
			  selectId+=","+v;
		  }
		}
	}else{
		var sarr=selectId.split(',');
		var nstr=''
		for(var i=0;i<sarr.length;i++){
			if (sarr[i]!=v){
				if (nstr==''){
					nstr=sarr[i]
				}else{
					nstr+=','+sarr[i];
				}
			}
		}
		selectId=nstr;
	}
	var p=new KesionPopup();
	p.MsgBorder=5;
	p.BgColor='#fff';
	p.ShowBackground=false;
	p.popup("������������","<strong><label><input type='checkbox' id='checkall' onclick='checkall()'/>ȫѡ</label> ��ѡ������ID����:</strong><span id='selids'>"+selectId+"</span><br /><br /><div><a href='javascript:void(0)' onclick=\"verifictopic(1,selectId,'"+dir+"',"+boardid+")\">�������ѡ������</a> | <a href='javascript:void(0)' onclick=\"verifictopic(0,selectId,'"+dir+"',"+boardid+")\">����ȡ�����</a> | <a href='javascript:void(0)' onclick=\"verifictopic(2,selectId,'"+dir+"',"+boardid+")\">��������ѡ������</a> | <a href='javascript:void(0)' onclick=\"delsubject(selectId,'"+dir+"',"+boardid+")\">����ɾ��ѡ�е�����</a><br/> <a href='javascript:void(0)' onclick=\"settop(selectId,'"+dir+"',"+boardid+",1)\">�����ö�ѡ�е�����</a> | <a href='javascript:void(0)' onclick=\"canceltop(selectId,'"+dir+"',"+boardid+")\">����ȡ���ö�</a> | <a href='javascript:void(0)' onclick=\"setbest(selectId,'"+dir+"',"+boardid+")\">�������þ���</a> | <a href='javascript:void(0)' onclick=\"cancelbest(selectId,'"+dir+"',"+boardid+")\">����ȡ������</a><br/><br/><strong>��ѡ�������ƶ�������</strong><br/><form name='moveform' action='../"+dir+"/ajax.asp' method='get'></b><span id='showboardselect'></span><input type='submit' value='ȷ���ƶ�' class='btn'><input type='hidden' value="+selectId+" name='id' id='id'>&nbsp;<input type='hidden' value='movetopic' name='action'></form></div><br/>",420);
	      if (bstr==''){
		  $.get("../plus/ajaxs.asp",{action:"GetClubBoardOption"},function(r){
		    $("#showboardselect").html(unescape(r));});
		  	      bstr=$("#showboardselect").html();

		  }else{
			  $("#showboardselect").html(bstr);
		  }
	      
	
	if (selectId==''){
		closeWindow();
	}
}
function checkall(){
	if ($("#checkall")[0].checked){
	selectId='';
	$(document).find("input[type=checkbox]").not("#checkall").each(function(){
           $(this).attr("checked",true);
	       if (selectId=='') {selectId=$(this).val()}else{selectId+=','+$(this).val();}
		   $("#selids").html(selectId);
																			});
	}else{
	$(document).find("input[type=checkbox]").not("#checkall").attr("checked",false);
	selectId='';
	closeWindow();
	}
}
function verifictopic(v,id,dir,boardid){
   	$.get("../"+dir+"/ajax.asp",{action:"verifictopic",v:v,id:id,boardid:boardid},function(r){
		if (r=="success"){switch(v){
			case 0 :alert('��ϲ,��ѡ������ȡ����˵Ĳ����ɹ���');break;
		    case 1:	alert('��ϲ,��ѡ������������˵Ĳ����ɹ���');break;
			case 2:	alert('��ϲ,��ѡ���������������Ĳ����ɹ���');break;
		}
		location.reload();}else{alert(r);}
	});
}
function settop(id,dir,boardid,v){
	if (!confirm('ȷ����Ϊ�ö���')) return;
   	$.get("../"+dir+"/ajax.asp",{action:"settop",id:id,boardid:boardid,v:v},function(r){
		if (r=="success"){
			alert('��ϲ,��ѡ�������ö������ɹ���')
		   location.reload();}else{alert(r);}
	});
}
function canceltop(id,dir,boardid){
	if (!confirm('ȷ��ȡ���ö���')) return;
   	$.get("../"+dir+"/ajax.asp",{action:"canceltop",id:id,boardid:boardid},function(r){
		if (r=="success"){
			alert('��ϲ,��ѡ������ȡ���ö������ɹ���')
		   location.reload();}else{alert(r);}
	});
}
function setbest(id,dir,boardid){
	if (!confirm('ȷ����Ϊ��������')) return;
   	$.get("../"+dir+"/ajax.asp",{action:"setbest",id:id,boardid:boardid},function(r){
		if (r=="success"){
			alert('��ϲ,��ѡ��������Ϊ�����������ɹ���')
		   location.reload();}else{alert(r);}
	});
}
function cancelbest(id,dir,boardid){
	if (!confirm('ȷ��ȡ����������')) return;
   	$.get("../"+dir+"/ajax.asp",{action:"cancelbest",id:id,boardid:boardid},function(r){
		if (r=="success"){
			alert('��ϲ,��ѡ������ȡ�������������ɹ���')
		   location.reload();}else{alert(r);}
	});
}
function topicfav(id,dir,boardid){
	$.get("../"+dir+"/ajax.asp",{action:"fav",id:id,topicid:id,boardid:boardid},function(r){
		if (r=="success"){alert('��ϲ,���ղأ�');}else{alert(r);}
		
	});
}
function locked(id,dir,boardid){
	 $.get("../"+dir+"/ajax.asp",{action:"locked",id:id,topicid:id,boardid:boardid},function(r){
			if (r=="success"){location.reload();}else{alert(r);}
	});
}
function unlocked(id,dir,boardid){
	 $.get("../"+dir+"/ajax.asp",{action:"unlocked",id:id,topicid:id,boardid:boardid},function(r){
			if (r=="success"){location.reload();}else{alert(r);}
	});
}
function delsubject(id,dir,bid){
	var p=new KesionPopup();
	p.MsgBorder=5;
	p.BgColor='#fff';
	p.ShowBackground=false;
	p.popup("����ɾ����ʾ","<strong>ɾ��ѡ�</strong><label><input type='hidden' value='"+id+"' id='did' name='did'/><input onclick=\"$('#oprzm').hide();\" type='radio' value='0' name='deltype' checked>�������վ</label> <label><input type='radio' value='1' name='deltype' onclick=\"$('#oprzm').show();\">����ɾ��</label><br/><div id='oprzm' style='display:none'><strong>�� ֤ �룺</strong><input type='text' name='rzm' id='rzm'> <br/><font color='#999999'>tips:����ɾ����Ҫ������֤�룬��֤��λ��conn.asp���趨�� </font></div><br/><input type='submit' id='delbtn' value='ȷ��ɾ��' class='btn' onclick=\"dodelsubject('"+dir+"',"+bid+");\"></div><br/>",420);
}
function dodelsubject(dir,bid){
	var id=$("#did").val();
	var deltype=$("input[name=deltype][checked]").val();
	var rzm=$("#rzm").val();
	if (parseInt(deltype)==1 && rzm==''){
		 alert('����ɾ���������������֤��!');
		 return;
	}
  if (parseInt(deltype)==1 && !(confirm('ɾ�����⣬���еĻظ���ɾ����ȷ��ִ��ɾ��������'))){
	  closeWindow();
  }
  $("#delbtn").attr("value","����ɾ����...");
  $("#delbtn").attr("disabled",true);
   $.get("../"+dir+"/ajax.asp",{action:"delsubject",id:id,boardid:bid,deltype:deltype,rzm:rzm},function(r){
  			if (r=="success"){alert('��ϲ,ɾ���ɹ�!');location.href='../'+dir+'/index.asp?boardid='+bid;}else{alert(r);$("#delbtn").attr("disabled",false);$("#delbtn").attr("value",'ȷ��ɾ��');$("#rzm").val('');}
  	});
}
function delreply(dir,topicid,replyid,boardid){
	var p=new KesionPopup();
	p.MsgBorder=5;
	p.BgColor='#fff';
	p.ShowBackground=false;
	p.popup("ɾ���ظ�","<strong>ɾ��ѡ�</strong><label><input onclick=\"$('#oprzm').hide();\" type='radio' value='0' name='deltype' checked>�������վ</label> <label><input type='radio' value='1' name='deltype' onclick=\"$('#oprzm').show();\">����ɾ��</label><br/><div id='oprzm' style='display:none'><strong>�� ֤ �룺</strong><input type='text' name='rzm' id='rzm'> <br/><font color='#999999'>tips:ɾ���û����Ӳ�����Ҫ������֤�룬��֤��λ��conn.asp���趨�� </font></div><br/><input type='submit' id='delbtn' value='ȷ��ɾ��' class='btn' onclick=\"dodelreply('"+dir+"',"+topicid+","+replyid+","+boardid+");\"></div><br/>",420);
}
function dodelreply(dir,topicid,replyid,boardid){
	var deltype=$("input[name=deltype][checked]").val();
	var rzm=$("#rzm").val();
	if (parseInt(deltype)==1 && rzm==''){
		 alert('����ɾ���������������֤��!');
		 return;
	}

  $("#delbtn").attr("value","����ɾ����...");
  $("#delbtn").attr("disabled",true);
  $.get("../"+dir+"/ajax.asp",{action:"delreply",deltype:deltype,id:topicid,replyid:replyid,boardid:boardid,rzm:escape(rzm)},function(r){
  			if (r=="success"){alert('��ϲ,ɾ���ɹ�!');location.reload()}else{alert(r);$("#delbtn").attr("disabled",false);$("#delbtn").attr("value",'ȷ��ɾ��');$("#rzm").val('');}
	});
}
function delusertopic(topicid,page,n,postusername,boardid,dir){
	var p=new KesionPopup();
	p.MsgBorder=5;
	p.BgColor='#fff';
	p.ShowBackground=false;
	p.popup("ɾ���û�[<font color=#ff6600>"+postusername+"</font>]�����з���","<strong>ɾ��ѡ�</strong><label><input onclick=\"$('#oprzm').hide();\" type='radio' value='0' name='deltype' checked>�������վ</label> <label><input type='radio' value='1' name='deltype' onclick=\"$('#oprzm').show();\">����ɾ��</label><br/><div id='oprzm' style='display:none'><strong>�� ֤ �룺</strong><input type='text' name='rzm' id='rzm'> <br/><font color='#999999'>tips:ɾ���û����Ӳ�����Ҫ������֤�룬��֤��λ��conn.asp���趨�� </font></div><br/><input type='submit' id='delbtn' value='ȷ��ɾ��' class='btn' onclick=\"dodelusertopic('"+dir+"',"+topicid+","+page+","+n+",'"+postusername+"',"+boardid+");\"></div><br/>",420);
}
function dodelusertopic(dir,topicid,page,n,postusername,boardid){
	var deltype=$("input[name=deltype][checked]").val();
	var rzm=$("#rzm").val();
	if (parseInt(deltype)==1 && rzm==''){
		 alert('����ɾ���������������֤��!');
		 return;
	}
  if (parseInt(deltype)==1 && !(confirm('ɾ�����⣬���еĻظ���ɾ����ȷ��ִ��ɾ��������'))){
	  closeWindow();
  }
  $("#delbtn").attr("value","����ɾ����...");
  $("#delbtn").attr("disabled",true);
  $.get("../"+dir+"/ajax.asp",{action:"delusertopic",deltype:deltype,topicid:topicid,page:page,n:n,username:escape(postusername),boardid:boardid,rzm:escape(rzm)},function(r){
																																												                var rstr=r.split('|');
				if (rstr[0]=="succ"){alert(rstr[1]);location.href=rstr[2];}else{alert(rstr[1]);$("#delbtn").attr("disabled",false);$("#delbtn").attr("value",'ȷ��ɾ��');$("#rzm").val('');}
	});
}

function support(topicid,id,dir){
	 $.get("../"+dir+"/ajax.asp",{action:"support",id:id,topicid:topicid},function(r){
			if (r=="error"){
				alert('����Ͷ��Ʊ��!');
			}else{
				$("#supportnum"+id).html(r);
			}
	});
}
function opposition(topicid,id,dir){
	 $.get("../"+dir+"/ajax.asp",{action:"opposition",id:id,topicid:topicid},function(r){
			if (r=="error"){
				alert('����Ͷ��Ʊ��!');
			}else{
				$("#oppositionnum"+id).html(r);
			}
	});
}
function checkmsg()
 {   var message=escape($("#message").val());
	 var username=escape($("#username").val());
	 if (username==''){
			  alert('�������ݳ���!');
			  closeWindow();
	 }
	 if (message==''){
			   alert('��������Ϣ����!');
			   $("#message").focus();
			   return false;
	 }
	 $("#sendmsgbtn").attr("disabled",true);
	 $.get("../plus/ajaxs.asp",{action:"SendMsg",username:username,message:message},function(r){
			   r=unescape(r);
	             $("#sendmsgbtn").attr("disabled",false);
			   if (r!='success'){
				alert(r);
			   }else{
				 alert('��ϲ��������Ϣ�ѷ���!');
				 closeWindow();
			   }
			 });
 }
function sendMsg(ev,username){
	new KesionPopup().popup("<img src='../images/user/mail.gif' align='absmiddle'>������Ϣ","�Է���¼����Կ���������Ϣ(������255���ַ�)<br /><textarea name='message' id='message' style='width:340px;height:80px'></textarea><div style='text-align:center;margin:10px'><input type='button' id='sendmsgbtn' onclick='return(checkmsg())' value=' ȷ �� ' class='btn'><input type='hidden' value="+username+" name='username' id='username'> <input type='button' value=' ȡ �� ' onclick='closeWindow()' class='btn'></div>",350);
		  $.get("../plus/ajaxs.asp",{action:"CheckLogin"},function(r){
		   if (r!='true'){
			 ShowLogin();
			}
		   });
}
function check()
{
		 var message=escape($("#message").val());
		 var username=escape($("#username").val());
		 if (username==''){
		  alert('�������ݳ���!');
		  closeWindow();
		 }
		 if (message==''){
		   alert('�����븽��!');
		   $("#message").focus();
		   return false;
		 }
		 $.get("../plus/ajaxs.asp",{action:"AddFriend",username:username,message:message},function(r){
		   r=unescape(r);
		   if (r!='success'){
		    alert(r);
		   }else{
		     alert('���������ѷ���,��ȴ��Է���ȷ��!');
			 closeWindow();
		   }
		 });
}
function addF(ev,username)
{ 
		 show(ev,username);
		 var isMyFriend=false;
		 $.get("../plus/ajaxs.asp",{action:"CheckMyFriend",username:escape(username)},function(b){
		    if (b=='nologin'){
			  closeWindow();
			  ShowLogin();
			}else if (b=='true'){
			  closeWindow();
			  alert('�û�['+username+']�Ѿ������ĺ����ˣ�');
			  return false;
			 }else if(b=='verify'){
			  closeWindow();
			  alert('���������['+username+'],��ȴ��Է�����֤!');
			  return false;
			 }else{
			 }
		 })
}
function show(ev,username){
	new KesionPopup().popup("<img src='../images/user/log/106.gif'>��Ӻ���","ͨ���Է���֤���ܳ�Ϊ����(������255���ַ�)<br /><textarea name='message' id='message' style='width:340px;height:80px'></textarea><div style='text-align:center;margin:10px'><input type='button' onclick='return(check())' value=' ȷ �� ' class='btn'><input type='hidden' value="+username+" name='username' id='username'> <input type='button' value=' ȡ �� ' onclick='closeWindow()' class='btn'></div>",350);
}
function ShowLogin()
{ 
	if(document.readyState=="complete"){ 
    var p=new KesionPopup();
	p.MsgBorder=1;
	p.ShowBackground=false;
	p.BgColor='#fff';
	p.TitleCss="font-size:14px;background:#1B76B7;color:#fff;height:22px;";
    p.popupIframe('<img src="/user/images/icon18.png" align="absmiddle">��Ա��¼','/user/userlogin.asp?Action=Poplogin',430,204,'no');}else{
		setTimeout(function(){ShowLogin();},10); 
	}
}
function checksearch()
{
     if ($("#keyword").val()=="")
	 {
	  alert('������ؼ���!');
	  $('#keyword').focus();
	  return false;
	 }
}
/*
*����Ie && Firefox ��CopyToClipBoard
*
*/
function copyToClipBoard(txt) {
    if (window.clipboardData) {
        window.clipboardData.clearData();
        window.clipboardData.setData("Text", txt);
    } else if (navigator.userAgent.indexOf("Opera") != -1) {
    } else if (window.netscape) {
        try {
            netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");
        } catch (e) {
            alert("��������ܾ���\n�����������ַ������'about:config'���س�\nȻ�� 'signed.applets.codebase_principal_support'����Ϊ'true'");
        }
        var clip = Components.classes['@mozilla.org/widget/clipboard;1'].createInstance(Components.interfaces.nsIClipboard);
        if (!clip)   return;
        var trans = Components.classes['@mozilla.org/widget/transferable;1'].createInstance(Components.interfaces.nsITransferable);
        if (!trans) return;
        trans.addDataFlavor('text/unicode');
        var str = new Object();
        var len = new Object();
        var str = Components.classes["@mozilla.org/supports-string;1"].createInstance(Components.interfaces.nsISupportsString);
        var copytext = txt;
        str.data = copytext;
        trans.setTransferData("text/unicode", str, copytext.length * 2);
        var clipid = Components.interfaces.nsIClipboard;
        if (!clip)   return false;
        clip.setData(trans, null, clipid.kGlobalClipboard);
    }
    alert("���Ѿ��ɹ����Ʊ���ַ����ֱ��ճ���Ƽ����������!");
}
function showOnlneList(){
	if ($("#onlineText").html()=='��ϸ�����б�'){
		$("#onlineText").html('�ر������б�');
		 $("#showOnline").fadeIn('slow');
		  $.get("../plus/ajaxs.asp",{action:"getonlinelist"},function(d){
			$("#showOnline").html(d);
			onlineList(1);
		   });
	}else{
		$("#onlineText").html('��ϸ�����б�');
		$("#showOnline").fadeOut('fast');
	}
}
function onlineList(p){
	  $.get("../plus/ajaxs.asp",{action:"getonlinelist",page:p},function(d){
			$("#showOnline").html(d);
		   });
}

function CopyCode(obj) {
	if (typeof obj != 'object') {
		if (document.all) {
			window.clipboardData.setData("Text",obj);
			alert('���Ƴɹ�!');
		} else {
			prompt('��Ctrl+C��������', obj);
		}
	} else if (document.all) {
		var js = document.body.createTextRange();
		js.moveToElementText(obj);
		js.select();
		js.execCommand("Copy");
		alert('���Ƴɹ�!');
	}
	return false;
}


