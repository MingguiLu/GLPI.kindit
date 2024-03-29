<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New GetRelativeList
KSCls.Kesion()
Set KSCls = Nothing

Class GetRelativeList
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim InstallDir, CurrPath, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript, LabelFlag
		Dim ChannelID,ClassID, IncludeSubClass, ShowClassName, OpenType, DocProperty, Num, RowHeight, TitleLen, OrderStr, ColNumber, ShowPicFlag, NavType, Navi, SplitPic, DateRule, DateAlign, TitleCss, DateCss,ShowNewFlag,ShowHotFlag, PrintType,AjaxOut,LabelStyle,IntroLen,RelativeType,relativeText
		Dim PicWidth,PicHeight,PicStyle,PicBorderColor,PicSpacing
		Dim ButtonType,PriceType,ProductType,Discount
		Dim TypeID,ShowGQType
		FolderID = Request("FolderID")
		CurrPath = KS.GetCommonUpFilesDir()
		ChannelID=KS.ChkCLng(Request("ChannelID"))
		
		
		With KS
		'判断是否编辑
		LabelID = Trim(Request.QueryString("LabelID"))
		If LabelID = "" Then
		  ClassID = "0"
		  Action = "Add"
		Else
		  Action = "Edit"
		  Dim LabelRS, LabelName
		  Set LabelRS = Server.CreateObject("Adodb.Recordset")
		  LabelRS.Open "Select * From KS_Label Where ID='" & LabelID & "'", Conn, 1, 1
		  If LabelRS.EOF And LabelRS.BOF Then
			 LabelRS.Close
			 Set LabelRS = Nothing
			 .echo ("<Script>alert('参数传递出错!');history.back();</Script>")
			 Exit Sub
		  End If
			LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
			FolderID = LabelRS("FolderID")
			LabelContent = LabelRS("LabelContent")
			LabelFlag = LabelRS("LabelFlag")
			LabelRS.Close:Set LabelRS = Nothing
			LabelStyle         = KS.GetTagLoop(LabelContent)
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetRelativeList", ""),"}" & LabelStyle&"{/Tag}", "")
            ' response.write LabelContent
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('标签加载出错!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
			  ChannelID          = Node.getAttribute("modelid")
			  ClassID            = Node.getAttribute("classid")
			  IncludeSubClass    = Node.getAttribute("includesubclass")
			  showclassname      = Node.getAttribute("showclassname")
			  DocProperty        = Node.getAttribute("docproperty")
			  OpenType           = Node.getAttribute("opentype")
			  Num                = Node.getAttribute("num")
			  RowHeight          = Node.getAttribute("rowheight")
			  TitleLen           = Node.getAttribute("titlelen")
			  IntroLen           = Node.getAttribute("introlen")
			  OrderStr           = Node.getAttribute("orderstr")
			  ColNumber          = Node.getAttribute("col")
			  ShowPicFlag        = Node.getAttribute("showpicflag")
			  NavType            = Node.getAttribute("navtype")
			  Navi               = Node.getAttribute("nav")
			  SplitPic           = Node.getAttribute("splitpic")
			  DateRule           = Node.getAttribute("daterule")
			  DateAlign          = Node.getAttribute("datealign")
			  TitleCss           = Node.getAttribute("titlecss")
			  DateCss            = Node.getAttribute("datecss")
			  ShowNewFlag        = Node.getAttribute("shownewflag")
			  ShowHotFlag        = Node.getAttribute("showhotflag")
			  PrintType          = Node.getAttribute("printtype")
			  AjaxOut            = Node.getAttribute("ajaxout")
			  
			  PicWidth           = Node.getAttribute("picwidth")
			  PicHeight          = Node.getAttribute("picheight")
			  PicStyle           = Node.getAttribute("picstyle")
			  PicBorderColor     = Node.getAttribute("picbordercolor")
			  PicSpacing         = Node.getAttribute("picspacing")
			  
			  ButtonType         = Node.getAttribute("buttontype")
			  PriceType          = Node.getAttribute("pricetype")
			  ProductType        = Node.getAttribute("producttype")
			  Discount           = Node.getAttribute("discount")
			  
			  TypeID             = Node.getAttribute("typeid")
			  ShowGQType         = Node.getAttribute("showgqtype")
			  RelativeType       = Node.getAttribute("relativetype")
			  relativeText       = Node.getAttribute("relativetext")

			End If
            
			Set Node=Nothing
			Set XMLDoc=Nothing
		End If
		If PrintType="" Then PrintType=1
		If Num = "" Then Num = 10
		If DocProperty = "" Then DocProperty = "00000"
		If RowHeight = "" Then RowHeight = 20
		If TitleLen = "" Then TitleLen = 30
		If IntroLen = "" Then IntroLen = 50
		If ColNumber = "" Then ColNumber = 1
		If ShowNewFlag="" Then ShowNewFlag=False
		If ShowHotFlag="" Then ShowHotFlag=False
		If PicWidth="" Then PicWidth=130
		If PicHeight="" Then PicHeight=90
		If PicStyle="" Then PicStyle=1
		If PicSpacing="" Then PicSpacing=2
		If ButtonType="" Then ButtonType=4
		If PriceType="" Then PriceType=0
		If ProductType="" Then ProductType=0
		If Discount="" or IsNull(Discount) Then Discount=true
		If TypeID="" Then TypeID=0
		If ShowGQType="" Or IsNull(ShowGQType) Then ShowGQType=true
		If AjaxOut="" Then AjaxOut=false
		If KS.IsNul(RelativeType) Then RelativeType=0
		If LabelStyle="" Then LabelStyle="[loop={@num}] " & vbcrlf & "<li><a href=""{@linkurl}"" target=""_blank"">{@title}</a></li>" & vbcrlf & "[/loop]"
		.echo "<html>"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.echo "<link href=""../admin_style.css"" rel=""stylesheet"">"
		.echo "<script src=""../../../ks_inc/Common.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../ks_inc/Jquery.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../ks_inc/kesion.box.js"" language=""JavaScript""></script>"
		%>
		<style>
		 .field{width:720px;}
		 .field li{cursor:pointer;float:left;border:1px solid #DEEFFA;background-color:#F7FBFE;height:18px;line-height:18px;margin:3px 1px 0px;padding:2px}
		 .field li.diyfield{border:1px solid #f9c943;background:#FFFFF6}
		</style>
		<script>
		var TempFieldStr='';
		var TempDateStr='';
		var TempTitleCss='';
		var GenericPicStyleOption="<option value='1'>①:仅显示缩略图</option><option value='2'>②:缩略图+名称:上下</option><option value='3'>③:缩略图+(名称+简介:上下):左右</option><option value='4'>④:(名称+简介:上下)+缩略图:左右</option>";
						 
		$(document).ready(function(){
		  $("#ChannelID").change(function(){
			  SetField($("#ChannelID").val());
			  SetPicStyle($("#ChannelID").val());
			  SetModelParam($("#ChannelID").val());
		    });
		   
		  
		  SetPicStyle($("#ChannelID").val()); //填充样式选项
		  $("#PicStyle").change(function(){
		    $("#ViewStylePicArea").html('<img style="border:1px solid #ccc;margin:5px" src="../../Images/View/S'+$(this).val()+'.gif" height="100" width="180" border="0">');
			if ($(this).val()==1){
			 if ($("#ShowPicTitleCss").html()!=null)	TempTitleStr=$("#ShowPicTitleCss").html();
			 $("#ShowPicTitleCss").empty();
			}else{
			$("#ShowPicTitleCss").html(TempTitleStr);
			}
		  });
		  $("#ViewStylePicArea").html('<img style="border:1px outset #ccc;margin:5px" src="../../Images/View/S<%=PicStyle%>.gif" height="100" width="180" border="0">');
		  try{
		  $("#PicStyle>option[value=<%=PicStyle%>]").attr("selected",true);
		  }catch(e){
		  }
		 
		  <%
		  If LabelID <> "" Then
		   .echo "$('#ChannelID').attr('disabled',true);"
		  End If
		  %>
		  TempFieldStr=$("#ShowFieldArea").html();
		  TempDateStr=$("#ShowTableDate").html();
		  TempTitleStr=$("#ShowTitleCss").html();
		  ChangeOutArea($("#PrintType>option[selected=true]").val());
		})
		
		function SetField(channelid)
		{  
		   switch (parseInt(channelid)){
		    case 5:
		     $("#ShowFieldArea").html(TempFieldStr+"<li onclick=\"InsertLabel('{@bigphoto}')\" title=\"商品大图\">商品大图</li><li onclick=\"InsertLabel('{@price_market}')\" title=\"市场价格\">市 场 价</li><li onclick=\"InsertLabel('{@price_member}')\" title=\"会员价\">会 员 价</li><li title=\"当前零售价\" onclick=\"InsertLabel('{@price}')\">当前零售价</li><li title=\"折扣率\" onclick=\"InsertLabel('{@discount}')\">折扣率</li><li title=\"品牌ID\" onclick=\"InsertLabel('{@brandid}')\">品牌ID号</li><li title=\"品牌名称\" onclick=\"InsertLabel('{@brandname}')\">品牌名称</li><li title=\"品牌英文名\" onclick=\"InsertLabel('{@brandename}')\">品牌英文名</li><li style=\"width:55px\" title=\"商品型号\" onclick=\"InsertLabel('{@promodel}')\">商品型号</li><li title=\"赠送积分\" onclick=\"InsertLabel('{@point}')\">赠送积分</li>");
			 break;
		    case 7:
		     $("#ShowFieldArea").html(TempFieldStr+"<li onclick=\"InsertLabel('{@movieact}')\" title=\"主要演员\">主要演员</li><li onclick=\"InsertLabel('{@moviedy}')\" title=\"影片导演\">影片导演</li><li title=\"播放时间\" onclick=\"InsertLabel('{@movietime}')\">播放时间</li><li title=\"影片语言\" onclick=\"InsertLabel('{@movieyy}')\">影片语言</li><li title=\"出产地区\" onclick=\"InsertLabel('{@moviedq}')\">出产地区</li><li title=\"所需点数\" onclick=\"InsertLabel('{@readpoint}')\">所需点数</li>");
		     break;
		    case 8:
		     $("#ShowFieldArea").html(TempFieldStr+"<li onclick=\"InsertLabel('{@validdate}')\" title=\"有效期\">有 效 期</li><li onclick=\"InsertLabel('{@typeid}')\" title=\"交易类别\">交易类别</li><li title=\"联系人\" onclick=\"InsertLabel('{@contactman}')\">联 系 人</li><li title=\"公司名称\" onclick=\"InsertLabel('{@companyname}')\">公司名称</li><li title=\"所在省份\" onclick=\"InsertLabel('{@province}')\">所在省份</li><li title=\"所在城市\" onclick=\"InsertLabel('{@city}')\">所在城市</li><li title=\"详细地址\" onclick=\"InsertLabel('{@address}')\">详细地址<li title=\"联系电话\" onclick=\"InsertLabel('{@tel}')\">联系电话</li></li>");
		     break;
			
		   default:
		     $("#ShowFieldArea").html(TempFieldStr);
		   }
		   
		   if ($("#PrintType").val()==4){
		      $(top.frames['FrameTop'].document).find('#ajaxmsg').toggle();
		  	  $.get('../../../plus/ajaxs.asp',{action:'GetFieldOption',channelid:channelid},function(data){
			  $("#ShowFieldArea").html($("#ShowFieldArea").html()+data)
			  $(top.frames['FrameTop'].document).find('#ajaxmsg').toggle();
			 });

		 }
		}
		
		function SetPicStyle(channelid)
		{ 
		   switch (parseInt(channelid))
		   { case 0:
		     case 1:
			 case 2:
			 case 3:
			   $("#PicStyle").empty();
			   $("#PicStyle").append(GenericPicStyleOption);
			  break;
			 case 4:
			   $("#PicStyle").empty();
			   $("#PicStyle").append(GenericPicStyleOption);
			   $("#PicStyle").append("<option value='5'>⑤:缩略图+(名称+类别+作者+时间:上下):左右</option>");
			   $("#PicStyle").append("<option value='6'>⑥:缩略图+(名称+介绍:上下+人气等):左右</option>");
			   break;
			 case 5:
			   $("#PicStyle").empty();
			   $("#PicStyle").append(GenericPicStyleOption);
			   $("#PicStyle").append("<option value='7'>⑤:缩略图+按钮</option>");
			   $("#PicStyle").append("<option value='8'>⑥:缩略图+名称+按钮:上下</option>");
			   $("#PicStyle").append("<option value='9'>⑦:缩略图+名称+价格+按钮:上下</option>");
			   $("#PicStyle").append("<option value='10'>⑧:缩略图+(价格+按钮:上下):左右</option>");
			   $("#PicStyle").append("<option value='11'>⑨:(缩略图+名称)+(价格+按钮):左右</option>");
			   $("#PicStyle").append("<option value='12'>⑩:缩略图+(名称+价格+按钮):左右</option>");
			   break;
			 case 7:
			   $("#PicStyle").empty();
			   $("#PicStyle").append(GenericPicStyleOption);
			   $("#PicStyle").append("<option value='13'>⑤:缩略图+(名称+主演+简介+按钮):左右</option>");
			   $("#PicStyle").append("<option value='14'>⑥:缩略图+(名称+简介+属性):左右</option>");
			   $("#PicStyle").append("<option value='15'>⑦:缩略图+(名称+主演+导演+简介+按钮):左右</option>");
			   break;
			 case 8:
			   $("#PicStyle").empty();
			   $("#PicStyle").append(GenericPicStyleOption);
			   $("#PicStyle").append("<option value='16'>⑤:缩略图+[(标题+地区+时间)+简介]:左右</option>");
			   $("#PicStyle").append("<option value='17'>⑥:缩略图+(名称+简介+属性):左右</option>");
			   break;
			 default:
			  break;
		   }
		}
		
		function SetModelParam(channelid)
		{
		  if (parseInt(channelid)<=1) 
		    $("#twbz").show() 
		  else $("#twbz").hide();
		  
		  if (parseInt(channelid)==5){
		   if (parseInt($("#PrintType").val())==2)   
		    $("#ModelParamArea").show();
		   else
		    $("#ModelParamArea").hide();
		   $("#ModelParamArea").empty();
		   $("#ModelParamArea").append("<tr class='tdbg'><td colspan='2'>按钮样式 <select style='width:160px' name='ButtonType' id='ButtonType'><option value='0'>不显示</option><option value='1'>显示购买按钮</option><option value='2'>显示收藏按钮</option><option value='3'>显示详情按钮</option><option value='4' selected>显示购买+收藏按钮</option><option value='5'>显示购买+详情按钮</option><option value='6'>显示收藏+详情按钮</option><option value='7'>显示购买+详情+收藏按钮</option></select> 价格样式 <select style='width:160px' class='textbox' name='PriceType' id='PriceType'><option value='0' selected>自动显示</option><option value='8'>只显示会员价</option><option value='1'>只显示原始零售价</option><option value='2'>只显示当前零售价</option><option value='3'>原始零售价+会员价</option><option value='4'>当前零售价+会员价</option><option value='5'>显示市场价+当前零售价</option><option value='6'>市场价+原始零售价+会员价</option><option value='7'>市场价+原价+当前价+会员价</option></select> 销售类型<input name='ProductType' type='radio' value='0' Checked>不限<input name='ProductType'  type='radio' value='1'>正常 <input name='ProductType' type='radio' value='2'>涨价 <input name='ProductType' type='radio' id='ProductType' value='3'>降价 <label><input type='checkbox' name='Discount' id='Discount' value='true'><font color=blue>显示折扣</font></label></td></tr>");
		   $("#ButtonType>option[value=<%=ButtonType%>]").attr("selected",true);
		   $("#PriceType>option[value=<%=PriceType%>]").attr("selected",true);
		   $("input[name=ProductType][value=<%=ProductType%>]").attr("checked",true);
		   <%if Channelid=5 and cbool(Discount)=true then .echo "$('#Discount').attr('checked',true);" %>
		  }
		 else if(parseInt(channelid)==8){
		  $("#ModelParamArea").show();
		  $("#ModelParamArea").empty();
		  
		  $("#ModelParamArea").append("<tr class='tdbg'><td colspan='2'>交易类型 <%= Replace(Replace(KS.ReturnGQType(TypeID,1),"""","\"""),vbcrlf,"\n")%>  <label><input type='checkbox' name='ShowGQType' id='ShowGQType'>显示交易类型</label></td></tr>");
		  $("#TypeID").css("width",120);
		  <%if ChannelID=8 Then%>
		  $("#TypeID>option[value=<%=ButtonType%>]").attr("selected",true);
		  <%if cbool(ShowGQType)=true then .echo "$('#ShowGQType').attr('checked',true);"%>
		  <%End If%>
		 }else{
		   $("#ModelParamArea").hide()
		  }
		}
		
		function InsertLabel(label)
		{
		  InsertValue(label);
		}
		var pos=null;
		 function setPos()
		 { if (document.all){
				$("#LabelStyle").focus();
				pos = document.selection.createRange();
			  }else{
				pos = document.getElementById("LabelStyle").selectionStart;
			  }
		 }
		 //插入
		function InsertValue(Val)
		{  if (pos==null) {alert('请先定位要插入的位置!');return false;}
			if (document.all){
				  pos.text=Val;
			}else{
				   var obj=$("#LabelStyle");
				   var lstr=obj.val().substring(0,pos);
				   var rstr=obj.val().substring(pos);
				   obj.val(lstr+Val+rstr);
			}
		 }
		
      function FieldInsertCode(fieldname,dbtype,dbname)
		{ 
		   if(pos==null) {alert('请先定位插入位置!');return false;}
		   var link="../FieldParam.asp?fieldname=" + fieldname + "&dbtype="+ dbtype + "&dbname=" + dbname+"&datasourcetype=0";
		  var p=new KesionPopup()
		  p.PopupImgDir="../../";
		  p.PopupCenterIframe('插入字段标签',link,350,230,'no');
		}		 
		
		function ChangeOutArea(Val)
		{
		 SetModelParam($("#ChannelID").val());
		 switch (parseInt(Val)){
		  case 2:
		   $("#DiyArea").hide();
		   $("#TableArea").hide();
		   $("#PicArea").show();
		   $("#ShowIntroArea").show();
   		   
		     $("#ShowPicTitleCss").html(TempTitleStr);
		     $("#ShowTitleCss").empty();
		   $("#ViewStylePicArea").html('<img style="border:1px outset #ccc;margin:5px" src="../../Images/View/S'+$("#PicStyle").val()+'.gif" height="100" width="180" border="0">');
		   break;
		  case 3:
		  case 4:
		  $("#DiyArea").show();
		  $("#TableArea").hide();
		  $("#PicArea").hide();
		  $("#ShowDiyDate").html(TempDateStr);
		  $("#ShowTableDate").html('')
		  $("#DateRule").attr("style","width:130px");
		  $("#ShowIntroArea").show();
		  break;
		  default :
		  $("#DiyArea").hide();
		  $("#PicArea").hide();
		  $("#TableArea").show();
		  $("#ShowTableDate").html(TempDateStr);
		  $("#ShowDiyDate").html('')
		  $("#DateRule").attr("style","width:268px");
		  $("#ShowIntroArea").hide();
		  $("#ShowTitleCss").html(TempTitleStr);
		  $("#ShowPicTitleCss").html('');
		  break;
		 }
		 SetField($("#ChannelID").val());
		 
		}
		function SetNavStatus()
		{
		  if ($("select[name=NavType]").val()==0)
		   {$("#NavWord").show();
			$("#NavPic").hide();
			}else{
		   $("#NavWord").hide();
		   $("#NavPic").show();}
		}
		
		function SetLabelFlag(Obj)
		{
		 if (Obj.value=='-1')
		  $("#LabelFlag").val(1);
		  else
		  $("#LabelFlag").val(0);
		}
		
		function CheckForm()
		{
		    if ($("input[name=LabelName]").val()=='')
			 {
			  alert('请输入标签名称');
			  $("input[name=LabelName]").focus(); 
			  return false
			  }
			var ChannelID=$("#ChannelID").val();
			 var DocProperty='';
			 $("input[name=DocProperty]").each(function(){
			     if ($(this).attr("checked")==true){
				  DocProperty=DocProperty+'1'
				 }else{
				  DocProperty=DocProperty+'0'
				 }      
			 })

			var NavType=1;
			var ShowClassName,ShowPicFlag,ShowNewFlag,ShowHotFlag;
			var OpenType=$("#OpenType").val();
			var Num= $("#Num").val();
			var RowHeight=$("input[name=RowHeight]").val();
			var TitleLen=$("input[name=TitleLen]").val();
			var IntroLen=$("input[name=IntroLen]").val();
			var OrderStr=$("#OrderStr").val();
			var ColNumber=$("input[name=ColNumber]").val();
			var Nav,NavType=$("select[name=NavType]").val();
			var SplitPic=$("input[name=SplitPic]").val();
			var DateRule= $("#DateRule").val();
			var DateAlign=$("select[name=DateAlign]").val();
			var TitleCss=$("input[name=TitleCss]").val();
			var DateCss=$("input[name=DateCss]").val();
			var PicWidth=$("input[name=PicWidth]").val();
			var PicHeight=$("input[name=PicHeight]").val();
			var PicStyle=$("#PicStyle").val();
			var PicBorderColor=$("input[name=PicBorderColor]").val();
			var PicSpacing=$("input[name=PicSpacing]").val();
			
			var PrintType=$("#PrintType").val();
			var AjaxOut=false;
			if ($("#AjaxOut").attr("checked")==true){AjaxOut=true}
			var IncludeSubClass=false;
			if ($("#IncludeSubClass").attr("checked")==true) IncludeSubClass=true;
			var ShowClassName =false;
			if ($("#ShowClassName").attr("checked")==true) ShowClassName = true
			var ShowPicFlag=false;
			if ($("#ShowPicFlag").attr("checked")==true)  ShowPicFlag= true
			var ShowHotFlag=false;
			if ($("#ShowHotFlag").attr("checked")==true)  ShowHotFlag= true
			var ShowNewFlag=false;
			if ($("#ShowNewFlag").attr("checked")==true)  ShowNewFlag= true
			var RelativeType=$("#RelativeType>option:selected").val();
			   
			if  (Num=='')  Num=10;
			if (RowHeight=='') RowHeight=20
			if  (TitleLen=='') TitleLen=30;
			if  (ColNumber=='') ColNumber=1;
			if  (NavType==0) Nav=$("input[name=TxtNavi]").val();
			 else  Nav=$("input[name=NaviPic]").val();
			
			var tagVal='{Tag:GetRelativeList labelid="0" printtype="'+PrintType+'" ajaxout="'+AjaxOut+'" modelid="'+ChannelID+'"  docproperty="'+DocProperty+'" relativetext="'+$("#relativeText option:selected").val()+'" relativetype="'+RelativeType+'" orderstr="'+OrderStr+'" opentype="'+OpenType+'" num="'+Num+'" titlelen="'+TitleLen+'" introlen="'+IntroLen+'" rowheight="'+RowHeight+'" col="'+ColNumber+'" showclassname="'+ShowClassName+'" showpicflag="'+ShowPicFlag+'" shownewflag="'+ShowNewFlag+'" showhotflag="'+ShowHotFlag+'" navtype="'+NavType+'" nav="'+Nav+'"  splitpic="'+SplitPic+'" daterule="'+DateRule+'" datealign="'+DateAlign+'" titlecss="'+TitleCss+'" datecss="'+DateCss+'" picwidth="'+PicWidth+'" picheight="'+PicHeight+'" picstyle="'+PicStyle+'" picbordercolor="'+PicBorderColor+'" picspacing="'+PicSpacing+'"';
			if (ChannelID==5){
			 var ButtonType=$("#ButtonType").val();
			 var PriceType =$("#PriceType").val();
			 var ProductType=$("input[name=ProductType][checked=true]").val();
			 var Discount=false;
			 if ($("#Discount").attr("checked")==true)  Discount= true;
			 tagVal += ' buttontype="'+ButtonType+'" pricetype="'+PriceType+'" producttype="'+ProductType+'" discount="' + Discount + '"';
			}else if(ChannelID==8){
			 var TypeID=$("#TypeID").val();
			 var ShowGQType=false;
			 if($("#ShowGQType").attr("checked")==true) ShowGQType=true;
			 tagVal += ' typeid="'+TypeID+'" showgqtype="'+ShowGQType+'"';
			}
			tagVal  +='}'+$("#LabelStyle").val()+'{/Tag}';
			
			$("input[name=LabelContent]").val(tagVal);
			
			$("#myform").submit();
		}
		</script>
		<%
		.echo "</head>"
		.echo "<body topmargin=""0"" leftmargin=""0"">"
		.echo "<div align=""center"">"
		.echo "<iframe src='about:blank' name='_hiddenframe' id='_hiddenframe' width='0' height='0'></iframe>"
		.echo "<form  method=""post"" id=""myform"" name=""myform"" action=""AddLabelSave.asp"" target='_hiddenframe'>"
		.echo " <input type=""hidden"" name=""LabelContent"">"
		.echo " <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""" & LabelFlag & """>"
		.echo " <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.echo " <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetRelativeList.asp"">"
		.echo KS.ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		.echo "            <tr id=""ClassArea"" class=tdbg>"
		.echo "              <td width=""50%"" height=""24"" colspan=""2"">选择范围"
		.echo "                <select name=""ChannelID"" id=""ChannelID"">"
		.echo "                 <option value=""0"">-不区分模型-</option>"
        .LoadChannelOption ChannelID
		.echo "                </select>"
		
		.echo "         <span style='color:green'>不区分模型可以关联非本模型下的信息,如文章可以关联出相关的商品</span>"			
			
		.echo "           <span style='display:none'>属性控制 <label><input name=""DocProperty"" type=""checkbox"" value=""1"""
		If mid(DocProperty,1,1) = 1 Then .echo (" Checked")
		.echo ">推荐</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox""  value=""2"""
		If mid(DocProperty,2,1) = 1 Then .echo (" Checked")
		  .echo ">滚动</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""3"""
		If mid(DocProperty,3,1) = 1 Then .echo (" Checked")
		  .echo ">头条</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""4"""
		If mid(DocProperty,4,1) = 1 Then .echo (" Checked")
		  .echo ">热门</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""5"""
		If mid(DocProperty,5,1) = 1 Then .echo (" Checked")
		  .echo ">幻灯</label></span> </div></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" width=""50%"" colspan=3>关联方式"
		.echo "                <select style=""width:200px;"" class='textbox' name=""RelativeType"" id=""RelativeType"">"
		                       If RelativeType="0" Then
		                        .echo "<option value='0' selected style='color:red'>手工添加关联</option>"
							   Else
		                        .echo "<option value='0' style='color:red'>手工添加关联</option>"
							   End If
							   If RelativeType="1" Then
		                        .echo "<option value='1' selected>关键词关联</option>"
							   Else
		                        .echo "<option value='1'>关键词关联</option>"
							   End If
							   If RelativeType="2" Then
		                        .echo "<option value='2' selected>录入者关联</option>"
							   Else
		                        .echo "<option value='2'>录入者关联</option>"
							   End If
				
		.echo "                </select>"
		
		.echo " 分类:<select id='relativeText'>"
		.echo "<option value=''>--不限--</option>"
		dim rs:set rs=conn.execute("select relativeText from ks_iteminfor Where Relativetext<>'' group by relativeText")
		do while not rs.eof
		  if trim(relativeText)=trim(rs(0)) then
		  .echo "<option value='"& rs(0) & "' selected>" & rs(0) &"</option>"
		  else
		  .echo "<option value='"& rs(0) & "'>" & rs(0) &"</option>"
		  end if
		rs.movenext
		loop
		rs.close :set rs=nothing
		.echo "</select>"
					Conn.Close:Set Conn = Nothing

		.echo " <br/><font color=green>1.""手工添加关联""指的是按在后台添加文章章时,将需要关联的文章手工关联进来<br/>2.""关键词关联""指的是按当前文章的关键词自动关联相关文章出来(较占用资源) <br/>3.""录入者关联""指的关联出与当前文章相同录入者的文章<br/> "
		.echo "              </td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" width=""50%"">排序方法"
		.echo "                <select style=""width:70%;"" class='textbox' name=""OrderStr"" id=""OrderStr"">"
					If OrderStr = "ID Desc" Then
					.echo ("<option value='ID Desc' selected>文档ID(降序)</option>")
					Else
					.echo ("<option value='ID Desc'>文档ID(降序)</option>")
					End If
					If OrderStr = "ID Asc" Then
					.echo ("<option value='ID Asc' selected>文档ID(升序)</option>")
					Else
					.echo ("<option value='ID Asc'>文档ID(升序)</option>")
					End If
					If OrderStr = "Rnd" Then
					.echo ("<option value='Rnd' style='color:blue' selected>随机显示</option>")
					Else
					.echo ("<option value='Rnd' style='color:blue'>随机显示</option>")
					End If
					
					If OrderStr = "AddDate Asc" Then
					.echo ("<option value='AddDate Asc' selected>更新时间(升序)</option>")
					Else
					.echo ("<option value='AddDate Asc'>更新时间(升序)</option>")
					End If
					If OrderStr = "AddDate Desc" Then
					 .echo ("<option value='AddDate Desc' selected>更新时间(降序)</option>")
					Else
					 .echo ("<option value='AddDate Desc'>更新时间(降序)</option>")
					End If
					If OrderStr = "Hits Asc" Then
					 .echo ("<option value='Hits Asc' selected>点击数(升序)</option>")
					Else
					 .echo ("<option value='Hits Asc'>点击数(升序)</option>")
					End If
					If OrderStr = "Hits Desc" Then
					  .echo ("<option value='Hits Desc' selected>点击数(降序)</option>")
					Else
					  .echo ("<option value='Hits Desc'>点击数(降序)</option>")
					End If

		.echo "         </select></td>"
		.echo "              <td height=""24"">" & KS.ReturnOpenTypeStr(OpenType) & "</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" colspan='2'>文档数量"
		.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num""    style=""width:40px;text-align:center"" onBlur=""CheckNumber(this,'文档数量');"" value=""" & Num & """>条 标题字数<input name=""TitleLen"" class=""textbox"" onBlur=""CheckNumber(this,'标题字数');"" type=""text""    style=""width:40px;;text-align:center"" value=""" & TitleLen & """> 行距"
		.echo "                <input name=""RowHeight"" class=""textbox"" type=""text"" id=""RowHeight2""    style=""width:40px;;text-align:center"" onBlur=""CheckNumber(this,'文档行距');"" value=""" & RowHeight & """>px 列数<input type=""text"" class=""textbox"" onBlur=""CheckNumber(this,'排列列数');""  style=""width:40px;text-align:center"" value=""" & ColNumber & """ name=""ColNumber""> <font color=red>Tips:若自定义样式输出,行距和列数请在您的样式里控制</font></td>"
		.echo "            </tr>"
		
		
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">输出格式"
		.echo " <select class='textbox'  name=""PrintType"" style='width:200px' id=""PrintType"" onChange=""ChangeOutArea(this.options[this.selectedIndex].value);"">"
        .echo "  <option value=""1"""
		If PrintType=1 Then .echo " selected"
		.echo ">文本列表样式(Table)</option>"
        .echo "  <option value=""2"""
		If PrintType=2 Then .echo " selected"
		.echo ">图片列表样式(Table)</option>"
        .echo "  <option value=""3"""
		If PrintType=3 Then .echo " selected"
		.echo ">自定义输出样式(不带自定义字段)</option>"
        .echo "  <option style='color:green' value=""4"""
		If PrintType=4 Then .echo " selected"
		.echo ">自定义输出样式(带自定义字段)</option>"
        .echo "</select>"
		.echo "            <label><input type='checkbox' name='AjaxOut' id='AjaxOut' value='1'"
		If AjaxOut="true" Then .echo " checked"
		.echo ">采用Ajax输出</label></td>"
		.echo "              <td><span id='ShowDiyDate'></span> <span id='ShowIntroArea'>简介字数<input type='text' class='textbox' style='text-align:center' name='IntroLen' id='IntroLen' value='" & IntroLen & "' size='4'></span></td>"
		.echo "            </tr>"
		.echo "            <tbody id=""DiyArea"">"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@autoid}')"">行 号</li><li onclick=""InsertLabel('{@linkurl}')"">链接URL</li> <li onclick=""InsertLabel('{@id}')"">文档ID</li><li onclick=""InsertLabel('{@title}')"">标 题</li><li onclick=""InsertLabel('{@fulltitle}')"" style='color:red'>不截断标题</li> <li onclick=""InsertLabel('{@intro}')"">简要介绍</li><li onclick=""InsertLabel('{@photourl}')"">图片地址</li><li onclick=""InsertLabel('{@adddate}')"">添加时间</li><li onclick=""InsertLabel('{@inputer}')"">录入员</li><li onclick=""InsertLabel('{@hits}')"">点击数</li><li onclick=""InsertLabel('{@newimg}')"" title='显示新信息图片标志' style='color:red;'>最新图标志</li><li onclick=""InsertLabel('{@hotimg}')"" title='显示热门信息图片标志' style='color:red;'>热门图标志</li><li onclick=""InsertLabel('{@classname}')"">当前栏目名称</li><li onclick=""InsertLabel('{@classurl}')"">当前栏目URL</li><li onclick=""InsertLabel('{@topclassname}')"">一级栏目名称</li><li onclick=""InsertLabel('{@topclassurl}')"">一级栏目URL</li></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'><textarea name='LabelStyle' onkeyup='setPos()' onclick='setPos()' id='LabelStyle' style='width:95%;height:150px'>" & LabelStyle & "</textarea></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='attention'><strong><font color=red>使用说明 :</font></strong><br />1、循环标签[loop=n][/loop]对可以省略,也可以平行出现多对；<br />2、输出格式选择不带自定义字段输出的运行效率高于带自定义字段,如果列表没有用到自定义字段请选择不带自定义字段输出</font></td>"
		.echo "            </tr>"
		.echo "           </tbody>"
		
		
		.echo "           <tbody id='ModelParamArea'></tbody>"

		
		.echo "          <tbody id='TableArea'>"
		.echo "           <tr class=tdbg>"
		 .echo "             <td colspan=2 height=""30"">附加显示 "
				   If cbool(ShowClassName) = true Then
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowClassName"" name=""ShowClassName"" checked>显示栏目</label>")
				   Else
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowClassName"" name=""ShowClassName"">显示栏目</label>")
				   End If
                    .echo "&nbsp;&nbsp;&nbsp;"
					 If cbool(ShowPicFlag) = True Then
					  .echo ("<label id='twbz'><input type=""checkbox"" value=""true"" id=""ShowPicFlag"" name=""ShowPicFlag"" checked>“图文”标志</label>")
					 Else
					  .echo ("<label id='twbz'><input type=""checkbox"" value=""true"" id=""ShowPicFlag"" name=""ShowPicFlag"">“图文”标志</label>")
					 End If
				   .echo "&nbsp;&nbsp;&nbsp;"
					 If  cbool(ShowNewFlag) = True Then
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowNewFlag"" name=""ShowNewFlag"" checked>最新文档标志</label>")
					 Else
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowNewFlag"" name=""ShowNewFlag"">最新文档标志</label>")
					 End If
				 .echo "&nbsp;&nbsp;&nbsp;"
					 If  cbool(ShowHotFlag) = True Then
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowHotFlag"" name=""ShowHotFlag"" checked>显示热门文档标志</label>")
					 Else
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowHotFlag"" name=""ShowHotFlag"">显示热门文档标志</label>")
					 End If
			   
		.echo "       　</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">导航类型"
		.echo "                <select name=""NavType"" style=""width:70%;"" class='textbox' onchange=""SetNavStatus()"">"
				   If LabelID = "" Or CStr(NavType) = "0" Then
					.echo ("<option value=""0"" selected>文字导航</option>")
					.echo ("<option value=""1"">图片导航</option>")
				   Else
					.echo ("<option value=""0"">文字导航</option>")
					.echo ("<option value=""1"" selected>图片导航</option>")
				   End If
		 .echo "               </select></td>"
		 .echo "             <td width=""50%"" height=""24"">"
			   If LabelID = "" Or CStr(NavType) = "0" Then
				  .echo ("<div align=""left"" id=""NavWord""> ")
				  .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" style=""width:70%;"" value=""" & Navi & """> 支持HTML语法")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"">")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""选择图片..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				  .echo ("</div>")
				Else
				  .echo ("<div align=""left"" id=""NavWord"" style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" style=""width:70%;""> 支持HTML语法")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"" value=""" & Navi & """>")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""选择图片..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				  .echo ("</div>")
				End If
		 .echo "             </td>"
		 .echo "           </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" colspan=""2"">分隔图片"
		.echo "                <input name=""SplitPic"" class=""textbox"" type=""text"" id=""SplitPic"" style=""width:61%;"" value=""" & SplitPic & """ readonly>"
		.echo "                <input class='button' name=""SubmitPic"" class='button' onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.SplitPic);"" type=""button"" id=""SubmitPic2"" value=""选择图片..."">"
		.echo "                <span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.SplitPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>"
		.echo "                <div align=""left""> </div></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" id='ShowTableDate'>日期格式"
		.echo "                <select class='textbox' style=""width:70%;"" name=""DateRule"" id=""DateRule"">"
		.echo KS.ReturnDateFormat(DateRule)
		.echo "                </select> </td>"
		.echo "              <td height=""24"">"
		.echo "                <div align=""left"">日期对齐"
		.echo "                  <select class=""textbox"" name=""DateAlign"" id=""select3"" style=""width:70%;"">"
							
					If LabelID = "" Or CStr(DateAlign) = "left" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 .echo ("<option value=""left""" & Str & ">左对齐</option>")
					If CStr(DateAlign) = "center" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 .echo ("<option value=""center""" & Str & ">居中对齐</option>")
					If CStr(DateAlign) = "right" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 .echo ("<option value=""right""" & Str & ">右对齐</option>")
					 
		.echo "                  </select>"
		.echo "                </div></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" id=""ShowTitleCss"">标题样式"
		.echo "                <input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:70%;"" value=""" & TitleCss & """></td>"
		.echo "              <td height=""24"">日期样式"
		.echo "                <input name=""DateCss"" class=""textbox"" type=""text"" id=""DateCss"" style=""width:70%;"" value=""" & DateCss & """></td>"
		.echo "            </tr>"
		.echo "              </tbody>"



		.echo "           <tbody id='PicArea'>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"">图片设置 宽"
		.echo "                <input name=""PicWidth"" class=""textbox"" type=""text"" id=""PicWidth"" size='4' value=""" & PicWidth & """>px 高<input name=""PicHeight"" class=""textbox"" type=""text"" id=""PicHeight"" size='4' value=""" & PicHeight & """>px</td>"
		.echo "                <td colspan='2' rowspan='5' id='ViewStylePicArea'>图片显示</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"">显示样式"
		.echo "                <select class='textbox' style='width:230px' name=""PicStyle"" id=""PicStyle"">"
							.echo ("<option value=""1"">①:仅显示缩略图</option>")
							.echo ("<option value=""2"">②:缩略图+名称:上下</option>")
							.echo ("<option value=""3"">③:缩略图+(名称+简介:上下):左右</option>")
							.echo ("<option value=""4"">④:(名称+简介:上下)+缩略图:左右</option>")
						 
		.echo "                </select> <font color=""#FF0000""> =>右边效果预览</font></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"">边框颜色 <input type=""text"" id=""PicBorderColor"" class=""textbox"" name=""PicBorderColor"" style=""width:120;"" value=""" & PicBorderColor & """><img border=0 id=""ColorThumbsBorderShow"" src=""../../images/rect.gif"" style=""cursor:pointer;background-Color:" & PicBorderColor & ";"" onClick=""Getcolor(this,'../../../ks_editor/SelectColor.asp','PicBorderColor');"" title=""选取颜色""> 可留空</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"">图片间距:<input type='text' class='textbox' name='PicSpacing' id='PicSpacing' value='" & PicSpacing & "' size='8' style='text-align:center'> px</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" id=""ShowPicTitleCss""></td>"
		.echo "            </tr>"
		.echo "           </tbody>"

		.echo "         </table>"			 
		.echo "    </form>"
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
		End With
		End Sub
End Class
%> 
