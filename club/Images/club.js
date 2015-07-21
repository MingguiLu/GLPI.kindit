//提示增加积分和威望,参数:bgdir 背景图路径,tipstr 提示信息,用逗号分开
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
//投票
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
	 alert('请选项择投票项!');
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
	p.popupIframe('查看投票详情',"../"+dir+"/showvoteuser.asp?voteid="+voteid,330,330,"auto")
}
function movetopic(dir,topicid,title){
		  new KesionPopup().popup("帖子移动","<form name='moveform' action='../"+dir+"/ajax.asp' method='get'><img src='../"+dir+"/images/p_up.gif'><b>移动帖子：</b>"+title+"<br/><br/><b>移到版面：</b><span id='showboardselect'></span><div style='text-align:center;margin:20px'><input type='submit' value='确定移动' class='btn'><input type='hidden' value="+topicid+" name='id' id='id'><input type='hidden' value='movetopic' name='action'><input type='button' value=' 取 消 ' onclick='closeWindow()' class='btn'></div></form>",500);
		  $.get("../plus/ajaxs.asp",{action:"GetClubBoardOption"},function(r){
		    $("#showboardselect").html(unescape(r));
	});
}
 //发帖
function Posted(){
	var p=new KesionPopup();
	p.MsgBorder=1;
	p.ShowBackground=false;
	p.BgColor='#fff';
	popTopHeight=200;
	p.TitleCss="font-size:14px;background:#1B76B7;color:#fff;height:22px;";
    var tips='<div style="background:url(../user/images/loginbg.png) repeat-x;padding:5px;">版面导航<span id="navlist1"></span><span id="navlist2"></span><br/><div id="boardlist"><img src="../images/loading.gif" /></div></div>';
    p.popup('<img src="../user/images/icon11.png" align="absmiddle"> 快速选择版面发帖',tips,600);
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
	 alert("点评字数不能超过"+cmax+"个字符!");
	}
	else {
	 $('#cmax').html(cmax-cobj.value.length-star.length);
	}
}
//点评
function comments(dir,topicid,replayid,boardid,n,userId){
	var p=new KesionPopup();
	p.MsgBorder=1;
	p.ShowBackground=false;
	p.BgColor='#fff';
	p.TitleCss="font-size:14px;background:#1B76B7;color:#fff;height:22px;";
    var tips='<div style="background:url(../user/images/loginbg.png) repeat-x;padding:5px;"><div id="comts"></div><textarea name="comment" onkeydown="checklength(this,255);" onkeyup="checklength(this,255);" id="comment" cols="50" rows="4" style="border:1px solid #ccc;color:#666;width:450px;height:90px"></textarea><div style="margin-top:6px;margin-bottom:10px;">威望<select id="Prestige"><option value="-1">-1</option><option value="-2">-2</option><option value="0">0</option><option value="1" selected>+1</option><option value="2">+2</option></select>&nbsp;&nbsp;<input type="button" onclick="saveComments(\''+dir+'\','+topicid+','+replayid+','+boardid+','+n+','+userId+')" class="btn" value="发表"/> <span style="color:#999">Tips:您还可以输入<span id="cmax">255</span>个字符!</span></div></div>';
    p.popup('<img src="../user/images/icon11.png" align="absmiddle"> 点评',tips,500);
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
var startnum = 5;	//星的个数
var selectedcolor = "#ff6600";	//选上的颜色
var uselectedcolor = "#999999";//未选的颜色
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
    $("#star"+k).val(presetpoint+'：<i>'+index+'</i>');
	checklength($("#comment")[0],255);
}
function showPresetPoint(s){
 if (s=='') return;
 var str='';
 PresetPoint=s.split(',');
 for (var k=0;k<PresetPoint.length;k++){
        str+=PresetPoint[k]+':<input type="hidden" name="star'+k+'" id="star'+k+'">';
	 for(var i=1;i<=startnum;i++){
			str+=('<span id="s'+k+i+'" style="color:#999;font-size:14px;" onclick="clickstar(\''+PresetPoint[k]+'\','+k+','+i+')" title="'+i+'星" onmouseout="setstar('+k+','+i+')" onmouseover="setstar('+k+','+i+')">★</span>');
		}
		str+="&nbsp;"
 }
  $("#comts").html(str);
}
function saveComments(dir,topicid,replayid,boardid,n,userId){
	var star='';
	var c=$("#comment").val();
	if(c==''){alert('请输入点评内容!');$("#comment").focus();return;}
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
//点评翻页显示
function ShowCmtPage(dir,p,pid,boardid){
	$("#comment_"+pid).html("加载中...");
	$.ajax({type:"get",url:"../"+dir+"/ajax.asp?action=getcommentpage&p="+p+"&pid="+pid+"&boardid="+boardid+"&anticache=" + Math.floor(Math.random()*1000),cache:false,dataType:"html",success:function(d){
			$("#comment_"+pid).html(d);																																																									   }
 });
}
//删除点评
function delCmt(dir,id,pid,boardid,p){
	if (confirm('删除后，不可恢复，确定删除吗？')){
	$.ajax({type:"get",url:"../"+dir+"/ajax.asp?action=delcomment&p="+p+"&id="+id+"&boardid="+boardid+"&anticache=" + Math.floor(Math.random()*1000),cache:false,dataType:"html",success:function(d){
	  if (d=="success"){	
	    alert('恭喜，删除成功!');
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
  alert('请选择要进入的子版面!');
}
function toPost(dir){
 if (dir!='') dir='../'+dir;
 var bid=$('#bid>option:selected').val();
 if (bid!='' && bid!=undefined)
 location.href=dir+'post.asp?bid='+bid;
 else
  alert('请选择要进入发帖子版面!');
}
function insertHiddenContent(ev){
	new KesionPopup().mousepop("插入隐藏内容","插入需要回复才能查看的内容<br /><textarea name='message' id='hidmessage' style='width:440px;height:120px'></textarea><div style='text-align:center;margin:10px'><input type='button' onclick='insertHidden();' value=' 确 定 ' class='btn'><input type='button' value=' 取 消 ' onclick='closeWindow()' class='btn'></div>",450);
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
	p.popup("帖子批量管理","<strong><label><input type='checkbox' id='checkall' onclick='checkall()'/>全选</label> 已选择帖子ID如下:</strong><span id='selids'>"+selectId+"</span><br /><br /><div><a href='javascript:void(0)' onclick=\"verifictopic(1,selectId,'"+dir+"',"+boardid+")\">批量审核选中主题</a> | <a href='javascript:void(0)' onclick=\"verifictopic(0,selectId,'"+dir+"',"+boardid+")\">批量取消审核</a> | <a href='javascript:void(0)' onclick=\"verifictopic(2,selectId,'"+dir+"',"+boardid+")\">批量锁定选中主题</a> | <a href='javascript:void(0)' onclick=\"delsubject(selectId,'"+dir+"',"+boardid+")\">批量删除选中的主题</a><br/> <a href='javascript:void(0)' onclick=\"settop(selectId,'"+dir+"',"+boardid+",1)\">批量置顶选中的主题</a> | <a href='javascript:void(0)' onclick=\"canceltop(selectId,'"+dir+"',"+boardid+")\">批量取消置顶</a> | <a href='javascript:void(0)' onclick=\"setbest(selectId,'"+dir+"',"+boardid+")\">批量设置精华</a> | <a href='javascript:void(0)' onclick=\"cancelbest(selectId,'"+dir+"',"+boardid+")\">批量取消精华</a><br/><br/><strong>将选中主题移动到版面</strong><br/><form name='moveform' action='../"+dir+"/ajax.asp' method='get'></b><span id='showboardselect'></span><input type='submit' value='确定移动' class='btn'><input type='hidden' value="+selectId+" name='id' id='id'>&nbsp;<input type='hidden' value='movetopic' name='action'></form></div><br/>",420);
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
			case 0 :alert('恭喜,对选中帖子取消审核的操作成功！');break;
		    case 1:	alert('恭喜,对选中帖子批量审核的操作成功！');break;
			case 2:	alert('恭喜,对选中帖子批量锁定的操作成功！');break;
		}
		location.reload();}else{alert(r);}
	});
}
function settop(id,dir,boardid,v){
	if (!confirm('确定设为置顶吗？')) return;
   	$.get("../"+dir+"/ajax.asp",{action:"settop",id:id,boardid:boardid,v:v},function(r){
		if (r=="success"){
			alert('恭喜,对选中主题置顶操作成功！')
		   location.reload();}else{alert(r);}
	});
}
function canceltop(id,dir,boardid){
	if (!confirm('确定取消置顶吗？')) return;
   	$.get("../"+dir+"/ajax.asp",{action:"canceltop",id:id,boardid:boardid},function(r){
		if (r=="success"){
			alert('恭喜,对选中主题取消置顶操作成功！')
		   location.reload();}else{alert(r);}
	});
}
function setbest(id,dir,boardid){
	if (!confirm('确定设为精华帖吗？')) return;
   	$.get("../"+dir+"/ajax.asp",{action:"setbest",id:id,boardid:boardid},function(r){
		if (r=="success"){
			alert('恭喜,对选中主题设为精华帖操作成功！')
		   location.reload();}else{alert(r);}
	});
}
function cancelbest(id,dir,boardid){
	if (!confirm('确定取消精华帖吗？')) return;
   	$.get("../"+dir+"/ajax.asp",{action:"cancelbest",id:id,boardid:boardid},function(r){
		if (r=="success"){
			alert('恭喜,对选中主题取消精华帖操作成功！')
		   location.reload();}else{alert(r);}
	});
}
function topicfav(id,dir,boardid){
	$.get("../"+dir+"/ajax.asp",{action:"fav",id:id,topicid:id,boardid:boardid},function(r){
		if (r=="success"){alert('恭喜,已收藏！');}else{alert(r);}
		
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
	p.popup("帖子删除提示","<strong>删除选项：</strong><label><input type='hidden' value='"+id+"' id='did' name='did'/><input onclick=\"$('#oprzm').hide();\" type='radio' value='0' name='deltype' checked>放入回收站</label> <label><input type='radio' value='1' name='deltype' onclick=\"$('#oprzm').show();\">彻底删除</label><br/><div id='oprzm' style='display:none'><strong>认 证 码：</strong><input type='text' name='rzm' id='rzm'> <br/><font color='#999999'>tips:彻底删除需要输入认证码，认证码位于conn.asp里设定。 </font></div><br/><input type='submit' id='delbtn' value='确定删除' class='btn' onclick=\"dodelsubject('"+dir+"',"+bid+");\"></div><br/>",420);
}
function dodelsubject(dir,bid){
	var id=$("#did").val();
	var deltype=$("input[name=deltype][checked]").val();
	var rzm=$("#rzm").val();
	if (parseInt(deltype)==1 && rzm==''){
		 alert('彻底删除，请输入操作认证码!');
		 return;
	}
  if (parseInt(deltype)==1 && !(confirm('删除主题，所有的回复将删除，确定执行删除操作吗？'))){
	  closeWindow();
  }
  $("#delbtn").attr("value","正在删除中...");
  $("#delbtn").attr("disabled",true);
   $.get("../"+dir+"/ajax.asp",{action:"delsubject",id:id,boardid:bid,deltype:deltype,rzm:rzm},function(r){
  			if (r=="success"){alert('恭喜,删除成功!');location.href='../'+dir+'/index.asp?boardid='+bid;}else{alert(r);$("#delbtn").attr("disabled",false);$("#delbtn").attr("value",'确定删除');$("#rzm").val('');}
  	});
}
function delreply(dir,topicid,replyid,boardid){
	var p=new KesionPopup();
	p.MsgBorder=5;
	p.BgColor='#fff';
	p.ShowBackground=false;
	p.popup("删除回复","<strong>删除选项：</strong><label><input onclick=\"$('#oprzm').hide();\" type='radio' value='0' name='deltype' checked>放入回收站</label> <label><input type='radio' value='1' name='deltype' onclick=\"$('#oprzm').show();\">彻底删除</label><br/><div id='oprzm' style='display:none'><strong>认 证 码：</strong><input type='text' name='rzm' id='rzm'> <br/><font color='#999999'>tips:删除用户帖子操作需要输入认证码，认证码位于conn.asp里设定。 </font></div><br/><input type='submit' id='delbtn' value='确定删除' class='btn' onclick=\"dodelreply('"+dir+"',"+topicid+","+replyid+","+boardid+");\"></div><br/>",420);
}
function dodelreply(dir,topicid,replyid,boardid){
	var deltype=$("input[name=deltype][checked]").val();
	var rzm=$("#rzm").val();
	if (parseInt(deltype)==1 && rzm==''){
		 alert('彻底删除，请输入操作认证码!');
		 return;
	}

  $("#delbtn").attr("value","正在删除中...");
  $("#delbtn").attr("disabled",true);
  $.get("../"+dir+"/ajax.asp",{action:"delreply",deltype:deltype,id:topicid,replyid:replyid,boardid:boardid,rzm:escape(rzm)},function(r){
  			if (r=="success"){alert('恭喜,删除成功!');location.reload()}else{alert(r);$("#delbtn").attr("disabled",false);$("#delbtn").attr("value",'确定删除');$("#rzm").val('');}
	});
}
function delusertopic(topicid,page,n,postusername,boardid,dir){
	var p=new KesionPopup();
	p.MsgBorder=5;
	p.BgColor='#fff';
	p.ShowBackground=false;
	p.popup("删除用户[<font color=#ff6600>"+postusername+"</font>]的所有发帖","<strong>删除选项：</strong><label><input onclick=\"$('#oprzm').hide();\" type='radio' value='0' name='deltype' checked>放入回收站</label> <label><input type='radio' value='1' name='deltype' onclick=\"$('#oprzm').show();\">彻底删除</label><br/><div id='oprzm' style='display:none'><strong>认 证 码：</strong><input type='text' name='rzm' id='rzm'> <br/><font color='#999999'>tips:删除用户帖子操作需要输入认证码，认证码位于conn.asp里设定。 </font></div><br/><input type='submit' id='delbtn' value='确定删除' class='btn' onclick=\"dodelusertopic('"+dir+"',"+topicid+","+page+","+n+",'"+postusername+"',"+boardid+");\"></div><br/>",420);
}
function dodelusertopic(dir,topicid,page,n,postusername,boardid){
	var deltype=$("input[name=deltype][checked]").val();
	var rzm=$("#rzm").val();
	if (parseInt(deltype)==1 && rzm==''){
		 alert('彻底删除，请输入操作认证码!');
		 return;
	}
  if (parseInt(deltype)==1 && !(confirm('删除主题，所有的回复将删除，确定执行删除操作吗？'))){
	  closeWindow();
  }
  $("#delbtn").attr("value","正在删除中...");
  $("#delbtn").attr("disabled",true);
  $.get("../"+dir+"/ajax.asp",{action:"delusertopic",deltype:deltype,topicid:topicid,page:page,n:n,username:escape(postusername),boardid:boardid,rzm:escape(rzm)},function(r){
																																												                var rstr=r.split('|');
				if (rstr[0]=="succ"){alert(rstr[1]);location.href=rstr[2];}else{alert(rstr[1]);$("#delbtn").attr("disabled",false);$("#delbtn").attr("value",'确定删除');$("#rzm").val('');}
	});
}

function support(topicid,id,dir){
	 $.get("../"+dir+"/ajax.asp",{action:"support",id:id,topicid:topicid},function(r){
			if (r=="error"){
				alert('您已投过票了!');
			}else{
				$("#supportnum"+id).html(r);
			}
	});
}
function opposition(topicid,id,dir){
	 $.get("../"+dir+"/ajax.asp",{action:"opposition",id:id,topicid:topicid},function(r){
			if (r=="error"){
				alert('您已投过票了!');
			}else{
				$("#oppositionnum"+id).html(r);
			}
	});
}
function checkmsg()
 {   var message=escape($("#message").val());
	 var username=escape($("#username").val());
	 if (username==''){
			  alert('参数传递出错!');
			  closeWindow();
	 }
	 if (message==''){
			   alert('请输入消息内容!');
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
				 alert('恭喜，您的消息已发送!');
				 closeWindow();
			   }
			 });
 }
function sendMsg(ev,username){
	new KesionPopup().popup("<img src='../images/user/mail.gif' align='absmiddle'>发送消息","对方登录后可以看到您的消息(可输入255个字符)<br /><textarea name='message' id='message' style='width:340px;height:80px'></textarea><div style='text-align:center;margin:10px'><input type='button' id='sendmsgbtn' onclick='return(checkmsg())' value=' 确 定 ' class='btn'><input type='hidden' value="+username+" name='username' id='username'> <input type='button' value=' 取 消 ' onclick='closeWindow()' class='btn'></div>",350);
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
		  alert('参数传递出错!');
		  closeWindow();
		 }
		 if (message==''){
		   alert('请输入附言!');
		   $("#message").focus();
		   return false;
		 }
		 $.get("../plus/ajaxs.asp",{action:"AddFriend",username:username,message:message},function(r){
		   r=unescape(r);
		   if (r!='success'){
		    alert(r);
		   }else{
		     alert('您的请求已发送,请等待对方的确认!');
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
			  alert('用户['+username+']已经是您的好友了！');
			  return false;
			 }else if(b=='verify'){
			  closeWindow();
			  alert('您已邀请过['+username+'],请等待对方的认证!');
			  return false;
			 }else{
			 }
		 })
}
function show(ev,username){
	new KesionPopup().popup("<img src='../images/user/log/106.gif'>添加好友","通过对方验证才能成为好友(可输入255个字符)<br /><textarea name='message' id='message' style='width:340px;height:80px'></textarea><div style='text-align:center;margin:10px'><input type='button' onclick='return(check())' value=' 确 定 ' class='btn'><input type='hidden' value="+username+" name='username' id='username'> <input type='button' value=' 取 消 ' onclick='closeWindow()' class='btn'></div>",350);
}
function ShowLogin()
{ 
	if(document.readyState=="complete"){ 
    var p=new KesionPopup();
	p.MsgBorder=1;
	p.ShowBackground=false;
	p.BgColor='#fff';
	p.TitleCss="font-size:14px;background:#1B76B7;color:#fff;height:22px;";
    p.popupIframe('<img src="/user/images/icon18.png" align="absmiddle">会员登录','/user/userlogin.asp?Action=Poplogin',430,204,'no');}else{
		setTimeout(function(){ShowLogin();},10); 
	}
}
function checksearch()
{
     if ($("#keyword").val()=="")
	 {
	  alert('请输入关键字!');
	  $('#keyword').focus();
	  return false;
	 }
}
/*
*兼容Ie && Firefox 的CopyToClipBoard
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
            alert("被浏览器拒绝！\n请在浏览器地址栏输入'about:config'并回车\n然后将 'signed.applets.codebase_principal_support'设置为'true'");
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
    alert("你已经成功复制本地址，请直接粘贴推荐给你的朋友!");
}
function showOnlneList(){
	if ($("#onlineText").html()=='详细在线列表'){
		$("#onlineText").html('关闭在线列表');
		 $("#showOnline").fadeIn('slow');
		  $.get("../plus/ajaxs.asp",{action:"getonlinelist"},function(d){
			$("#showOnline").html(d);
			onlineList(1);
		   });
	}else{
		$("#onlineText").html('详细在线列表');
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
			alert('复制成功!');
		} else {
			prompt('按Ctrl+C复制内容', obj);
		}
	} else if (document.all) {
		var js = document.body.createTextRange();
		js.moveToElementText(obj);
		js.select();
		js.execCommand("Copy");
		alert('复制成功!');
	}
	return false;
}


