$(document).ready(function(){loadDocRating();});
function loadDocRating(){
		  $.getScript(url+"plus/DocRating.asp?id="+itemid+"&m_id="+channelid+"&c_id="+infoid+"&title="+title,function(){
		    $("#DocRating").html(data.str); });
}
function PopRating(){
		   var str='<div id="popshow"><img src="'+url+'images/loading.gif"/>������...</div>';
		  	 new KesionPopup().popup("��Ҫ��������",str,400);
			  $.getScript(url+"plus/DocRating.asp?action=ShowPopup&id="+itemid+"&m_id="+channelid+"&c_id="+infoid,function(){
			   if (popu.islogin=='false'){
			    closeWindow();alert('���ȵ�¼!');
				new KesionPopup().popupIframe('��Ա��¼',url+'user/userlogin.asp?Action=Poplogin',397,184,'no');
			   }else{
			   $("#popshow").html(popu.str);	
			   }
			});
}
function PostMyScore(){
		  var score=$("#myscore option:selected").val();
		  var myitem=$("input[name=myitem][checked=true]").val()
		  $.getScript(url+"plus/DocRating.asp?score="+score+"&itemid="+myitem+"&action=hits&id="+itemid+"&m_id="+channelid+"&c_id="+infoid+"&title="+title,function(){
		     switch(vote.status){
				  case "nologin":
				   alert('�Բ���,����û��¼���ܴ��!');
				   break;
				  case "standoff":
				   alert('���ѱ�̬����, �����ظ����!');
				   break;
				  case "lock":
				   alert('����ѹر�!');
				   break;
				  case "errstartdate":
				   alert('δ�����ʱ��!');
				   break;
				  case "errexpireddate":
				   alert('���ʱ���ѹ�!');
				   break;
				  case "errgroupid":
				   alert('�����ڵ��û���û�д�ֵ�Ȩ��!');
				   break;
				  case "noinfo":
				   alert('�Ҳ�����Ҫ��ֵ���Ϣ!');
				   break;
				  default:
				   closeWindow();
				   alert('��ϲ,���ѳɹ����!');
				   loadDocRating();
				   break;
				 }
		  });
 }