<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
Dim KSCls
Set KSCls = New Admin_ShopOrder
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_ShopOrder
        Private KS,KSCls
		Private totalPut, CurrentPage, MaxPerPage,DomainStr
		Private SqlStr,PageTotalMoney1,PageTotalMoney2,SqlTotalMoney,RS,SqlParam,SearchType
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		  DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		Public Sub Kesion()
		
		   If Not KS.ReturnPowerResult(5, "M510012") Then                  '权限检查
			Call KS.ReturnErr(1, "")   
			Response.End()
		  End iF

			SearchType=KS.ChkClng(KS.G("SearchType"))
		%>
<html>
<head><title>订单处理</title>
<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<link href="include/Admin_Style.css" type=text/css rel=stylesheet>
<script src="../ks_inc/jquery.js"></script>
<script src="../ks_inc/common.js"></script>
<script src="../ks_inc/kesion.box.js"></script>
<script type="text/javascript">
  function modifyPrice(ev,title,orderid,id,price)
  {
    new KesionPopup().mousepop("<b>商品价格</b>","<iframe style='display:none' src='about:blank' id='_framehidden' name='_framehidden' width='0' height='0'></iframe><form name='rform' target='_framehidden' action='KS.ShopOrder.asp?action=ModifyPrice' method='post'>商品名称:"+title+"<br/><input type='hidden' value='"+price+"' name='oprice'><input type='hidden' name='orderId' value='"+orderid+"'><input type='hidden' name='Id' value='"+id+"'>实收价格:<input type='text' value='"+price+"' name='price' style='width:40px;text-align:center'>元<br /><input style='margin-top:7px' class='button' type='submit' value='确定修改'></form>",240);
  }
  function modifytotalprice(id,moneytotal){
    new KesionPopup().mousepop("<b>修改订单总价</b>","<iframe style='display:none' src='about:blank' id='_framehidden' name='_framehidden' width='0' height='0'></iframe><form name='rform' target='_framehidden' action='KS.ShopOrder.asp?action=ModifyTotalPrice' method='post'>当前价格:￥"+moneytotal+"元<br/><input type='hidden' value='"+moneytotal+"' name='oprice'><input type='hidden' name='Id' value='"+id+"'>将订单总价格改为:<input type='text' value='"+moneytotal+"' name='price' style='width:60px;text-align:center'>元<br /><input style='margin-top:7px' class='button' type='submit' value='确定修改'></form>",240);
  }
  function modifyInfo(id)
  {
    new KesionPopup().PopupCenterIframe('修改送货资料','KS.ShopOrder.asp?action=modifyinfo&id='+id,650,400,'auto')
  }
  function modifyproduct(id){
    new KesionPopup().PopupCenterIframe('修改/添加商品','KS.ShopOrder.asp?action=modifyproduct&id='+id,750,440,'auto')
  }
</script>
</head>
<body leftMargin=2 topMargin=0 marginheight="0" marginwidth="0">
 <%
   If KS.G("Action")="PrintOrder" Then
     Call PrintOrder()
     Response.end
   End IF
   If KS.G("Action")<>"modifyinfo" and KS.G("Action")<>"modifyproduct" Then
  %>
  <div class="topdashed" style="padding:4px;">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
<FORM name=form1 action=KS.ShopOrder.asp method=get>
      <td><strong>订单处理：</strong></td>
      <td valign="top">快速查询： 
<Select onchange=javascript:submit() size=1 name=SearchType> 
  <Option value=0<%If SearchType="0" Then Response.write " selected"%>>所有订单</Option> 
  <Option value=1<%If SearchType="1" Then Response.write " selected"%>>24小时之内的新订单</Option> 
  <Option value=2<%If SearchType="2" Then Response.write " selected"%>>最近10天内的新订单</Option> 
  <Option value=3<%If SearchType="3" Then Response.write " selected"%>>最近一月内的新订单</Option> 
  <Option value=4<%If SearchType="4" Then Response.write " selected"%>>未确认的订单</Option> 
  <Option value=5<%If SearchType="5" Then Response.write " selected"%>>未付款的订单</Option> 
  <Option value=6<%If SearchType="6" Then Response.write " selected"%>>未付清的订单</Option> 
  <Option value=7<%If SearchType="7" Then Response.write " selected"%>>未送货的订单</Option> 
  <Option value=8<%If SearchType="8" Then Response.write " selected"%>>未签收的订单</Option> 
  <Option value=9<%If SearchType="9" Then Response.write " selected"%>>未开发票的订单</Option> 
  <Option value=11<%If SearchType="11" Then Response.write " selected"%>>未结清的订单</Option> 
  <Option value=12<%If SearchType="12" Then Response.write " selected"%>>已结清的订单</Option>
      </Select></td></FORM>
<FORM name=form2 action=KS.ShopOrder.asp method=post>
      <td><B>高级查询：</B> 
	<Select id="Field" name="Field"> 
  <Option value=1>订单编号</Option> 
  <Option value=2>收货人</Option> 
  <Option value=3>用户名</Option> 
  <Option value=4>联系地址</Option> 
  <Option value=5>联系电话</Option> 
  <Option value=6>下单时间</Option>
  <Option value=7>推荐人</Option>
</Select> 
  <Input class='textbox' id=Keyword maxLength=30 name=Keyword> 
  <Input type=submit value=" 查 询 " class='button' name=Submit2> 
        <Input id=SearchType type=hidden value=10 name=SearchType> </td></FORM>
    </tr>
  </table>
  </div>
  <%
   End If
  
  
		  Select Case KS.G("Action")
		   Case "ModifyTotalPrice"
		    Call ModifyTotalPrice()
		   Case "modifyinfo"
		    Call modifyinfo()
		   Case "DoModifyInfoSave"
		    Call DoModifyInfoSave()
		   Case "modifyproduct"
		    Call modifyproduct()
		   Case "doModifyProductSave"
		    Call doModifyProductSave()
		   Case "ProAddToOrder"
		    Call ProAddToOrder()
		   Case "delproduct"
		    Call delproduct()
		   Case "ShowOrder"
		    Call ShowOrder()
		   Case "DelOrder"
		    Call DelOrder()
		   Case "OrderConfirm"
		    Call OrderConfirm()
		   Case "BankPay"     '付款
		    Call BankPay() 
		   Case "DoBankPay"    '银行付款操作
		    Call DoBankPay()
		   Case "BankRefund"    '退款
		    Call BankRefund()
		   Case "DoRefundMoney" '退款操作
		    Call DoRefundMoney()
		   Case "DeliverGoods"  '发货
		    Call DeliverGoods()
		   Case "DoDeliverGoods" '发货操作 
		    Call DoDeliverGoods()
		   Case "BackGoods"     '退货
		    Call BackGoods()
		   Case "SaveBack"     '退货操作
		     Call SaveBack()
		   Case "PayMoney"      '支付货款给卖方
		    Call PayMoney()
		   Case "DoPayMoney"    '支付货款
		    Call DoPayMoney()
		   Case "Invoice"   '开发票
		    Call Invoice()
		   Case "DoSaveInvoice"
		    Call DoSaveInvoice()
		   Case "ClientSignUp"   '已签收商品
		    Call ClientSignUp()
		   Case "FinishOrer"     '结算清单
		    Conn.Execute("Update KS_Order Set Status=2 Where ID=" & KS.G("ID"))
			Response.Redirect "KS.ShopOrder.asp?Action=ShowOrder&ID=" & KS.G("ID")
		   Case "ModifyPrice"    '修改指定价
		    Call ModifyPrice()
		   Case Else
		    Call OrderList
		  End Select
		End Sub
		'修改商品
		Sub modifyproduct()
		  If KS.ReturnPowerResult(0, "M520013")=false Then
		    response.write "<script>alert('对不起，您没有权限修改订单!');parent.closeWindow();</script>"
			response.end
		  End If
		 Dim RSI,OrderID
		 OrderID=KS.G("ID")
		 %>
		 <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'> 		  
		 <tr align='center' class='title' height='25'>  		   
		   <td><b>商 品 名 称</b></td> 		   
		   <td width='45'><b>单位</b></td>  		   
		   <td width='55'><b>数量</b></td>  		   
		   <td width='65'><b>原价</b></td>  		   
		   <td width='65'><b>实价</b></td>  		   
		   <td width='85'><b>小计</b></td>   		   
		   <td width='45'><b>操作</b></td>  		  
		  </tr> 
		  <form name="myform" action="KS.ShopOrder.asp" method="post">
		  <input type="hidden" name="action" value="doModifyProductSave"/>
		  <input type="hidden" name="orderid" value="<%=orderid%>"/>
		 <%
		 Dim SQLStr
		 SQLStr="Select i.*,P.Title,P.Unit From KS_OrderItem I Left Join KS_Product P  On I.ProID=P.ID Where I.SaleType<>5 and I.SaleType<>6 and I.OrderID='" & OrderID & "' order by i.ischangedbuy,i.id"
		 Set RSI=Server.CreateObject("ADODB.RECORDSET")
		 RSI.Open sqlstr,conn,1,1
		 If RSI.Eof And RSI.Bof Then
		    Response.Write "<tr class='tdbg'><td colspan=10 align='center'>该订单里没有商品!</td></tr>"
		 Else
		   Do While Not RSI.Eof
		   %>
		   <tr valign='middle' class='tdbg' height='20'>	  
		    <td width='*'><%=RSI("Title")%></td> 
			<td width='45' align=center><%=RSI("Unit")%></td>
			<td width='55' style='text-align:center'>
			 <input type="hidden" value="<%=rsi("id")%>" name="id" />
			<input type="text" name="amount<%=rsi("id")%>" value="<%=RSI("Amount")%>" size="4" style="text-align:center"></td>
			<td width='65' style='text-align:center'><input type="text" name="price_original<%=rsi("id")%>" value="<%=RSI("Price_Original")%>" size="5"/></td>    	   
			<td width='65' style='text-align:center'><input type="text" name="realprice<%=rsi("id")%>" value="<%=rsi("RealPrice")%>" size="5"/></td>    	   
			<td width='85' align='right'><%=formatnumber(rsi("realprice")*rsi("amount"),2,-1,-1)%> 元</td>
			 <td  style='text-align:center' width='45'>
			  <a href="?action=delproduct&orderid=<%=rsi("orderid")%>&id=<%=rsi("id")%>" onClick="return(confirm('确定将该商品从本订单中移除吗?'))">删除</a>
			 </td>  	   
			 </tr>
		   <%
		   RSI.MoveNext
		   Loop
		 End If
		 %>
		 <tr class="tdbg">
		   <td colspan=8>
		     <input type="submit" value="批量修改" class="button" /> <font color="blue">说明：批量修改将会重新计算订单的运费，订单总额等。</font>
		   </td>
		  </tr>
		  </form>
		 </table>
		 
		<script type="text/javascript">
		  function getProduct()
		  {			 
		     $(parent.parent.frames["FrameTop"].document).find("#ajaxmsg").toggle("fast");
			 var key=escape($('input[name=key]').val());
			 var tid=$('#tid>option:selected').val();
			 var priceType=$('#PriceType>option:selected').val();
			 var minPrice=$("#minPrice").val();
			 var maxPrice=$("#maxPrice").val();
			 var str='';
			 if (key!=''){
			   str='商品名称:'+key;
			 } 
			 if (tid!=''){
			   str+=' 栏目:'+$('#tid>option:selected').get(0).text
			 }
			 if (priceType!=0){
			   str+= minPrice +' 元';
			   switch (parseInt(priceType)){
			     case 1 :
				  str+='<=当前零售价<=';
				  break;
			     case 2 :
				   str+='<=会员价<=';
				   break;
			     case 3 :
				  str+='<=原始零售价<=';
				  break;
			   }
			   str+= maxPrice +' 元';
			   
			 }
			 if (str!='') str='<strong>条件:</strong><font color=red>'+str+'</font>';
			 $("#keyarea").html(str);
			 
			 $.get("../plus/ajaxs.asp", { action: "GetPackagePro", proid:$("#proids").val(),pricetype:priceType,key: key,tid:tid,minPrice:minPrice,maxPrice:maxPrice},
			 function(data){
					$(parent.parent.frames["FrameTop"].document).find("#ajaxmsg").toggle("fast");
					$("#prolist").empty().append(data);
			  });
		  }
		</script>
		<div style="border:1px dashed #cccccc;margin:3px;padding:4px">
		<table width="100%" border="0">
		  <tr>
			<td style="text-align:left">
			  &nbsp;<strong>快速搜索=></strong>
			  <br/>
			   &nbsp;商品编号: <input type="text" class="textbox" name="proids" id="proids" size='15'> 可留空<br/>
			 &nbsp;商品名称: <input type="text" class='textbox' name="key">
			 <br/>&nbsp;所属栏目: <select size='1' name='tid' id='tid'><option value=''>--栏目不限--</option><%=KS.LoadClassOption(5,false)%></select>
			 <br/>&nbsp;价格范围:
			<input type='text' name='minPrice' size='5' style='text-align:center' id='minPrice' value='10'> 元
			<= <select name="PriceType" id="PriceType">
			  <option value=0>--不限制--</option>
			  <option value=1>当前零售价</option>
			  <option value=2>会员价</option>
			  <option value=3>原始零售价</option>
			 </select>
			 <= <input type='text' name='maxPrice' size='5' style='text-align:center' id='maxPrice' value='100'> 元
			  
			  <br/> <br/>
			  &nbsp;<input type="button" onClick="getProduct()" value="开始搜索" class="button" name="s1">
			
			</td>
			<form name="myform" id="myform" action="KS.ShopOrder.asp?action=ProAddToOrder" method="post">
		  <input type="hidden" name="orderid" value="<%=orderid%>"/>
			<td>
			<div id='keyarea'></div>
			<strong>查询到的商品:</strong>			
			<br/>
			 <select name="prolist" size="5" style="width:260px;height:140px" multiple="multiple" id="prolist"></select>
			 <br/>
			 <input type="submit" value="将选中的商品加入到本订单" class="button">
			</td>
			</form>
		  </tr>
		</table>
		 </div>
		 <%RSI.Close
		 Set RSI=Nothing
		End Sub
		
		'保存修改
		Sub doModifyProductSave()
		 dim orderid:orderid=ks.s("orderid")
		 dim id:id=ks.filterids(ks.s("id"))
		 if id="" then ks.alerthintscript "没有商品!"
		 dim idarr,i
		 idarr=split(id,",")
		 for i=0 to ubound(idarr)
		    conn.execute("update ks_orderitem set amount=" & KS.G("amount" & trim(IDArr(i))) & ",price_original=" & KS.G("price_original"&Trim(IDArr(i))) &",realprice=" & KS.G("realprice"&Trim(IDArr(i))) & " Where ID=" & IDArr(i))
		 next
		 call updateorderprice(orderid)
		 KS.Die "<script>alert('恭喜，订单商品修改成功');parent.location.reload();</script>"
		 
		End Sub
		
		'商品加入订单
		Sub ProAddToOrder()
		 dim orderid:orderid=ks.g("orderid")
		 dim prolist:prolist=ks.filterids(ks.g("prolist"))
		 if orderid="" then ks.die "error!"
		 if ks.isnul(prolist) then ks.alerthintscript "对不起，您没有选择商品!"
		 dim rs:set rs=server.CreateObject("adodb.recordset")
		 rs.open "select * from ks_product where id in("&prolist&")",conn,1,1
		 if not rs.eof then
			 do while not rs.eof 
			  dim rsi:set rsi=server.CreateObject("adodb.recordset")
			  rsi.open "select top 1 * from ks_orderitem where proid=" & rs("id"),conn,1,3
			  if rsi.eof then
				  rsi.addnew
				  rsi("orderid")=orderid
				  rsi("proid")=rs("id")
				  rsi("SaleType")=RS("ProductType")
				  rsi("Price_Original")=RS("Price_Original")
				  rsi("Price")=RS("Price")
				  rsi("IsChangedBuy")=0
				  rsi("LimitBuyTaskID")=0
				  rsi("IsLimitBuy")=0
				  rsi("RealPrice")=RS("Price_Member")
				  rsi("Amount")=1
				  rsi("AttributeCart")=""
				  rsi("TotalPrice")=RS("Price_Member")
				  rsi("BeginDate")=Now
				  rsi("ServiceTerm")=RS("ServiceTerm")
				  rsi("PackID")=0
				  rsi("BundleSaleProID")=0
				  rsi.update
			 end if
			 rsi.close:set rsi=nothing
			 rs.movenext
			 loop 
			 call updateorderprice(orderid)
		 end if
			 rs.close
			 set rs=nothing
		 ks.alertHintscript "恭喜，已成功将选中的商品加入订单中!"
		End Sub
		
		Sub delproduct()
		 If KS.ReturnPowerResult(0, "M520013")=false Then
		    response.write "<script>alert('对不起，您没有权限修改订单!');parent.closeWindow();</script>"
			response.end
		  End If
		  dim id:id=KS.ChkClng(KS.S("ID"))
		  dim orderid:orderid=ks.s("orderid")
		  Conn.Execute("Delete From KS_OrderItem Where ID=" & ID)
			 call updateorderprice(orderid)
		 ks.alertHintscript "恭喜，已成功将选中的商品从订单中移除!"
		End Sub
		
		'修改送货信息
		Sub modifyinfo()
		If KS.ReturnPowerResult(0, "M520013")=false Then
		    response.write "<script>alert('对不起，您没有权限修改订单!');parent.closeWindow();</script>"
			response.end
		  End If
		 dim id:id=KS.ChkClng(Request("id"))
		 if id=0 then ks.die "error!"
		 dim rs:set rs=server.CreateObject("adodb.recordset")
		 rs.open "select top 1 * from ks_order where id=" & id,conn,1,1
		 if rs.eof and rs.bof then
		   rs.close : set rs=nothing
		   ks.die "error!"
		 end if
		%>
			<table border="0" cellpadding="2" cellspacing="1" class="border" width="100%">
			<form name="myform" action="KS.ShopOrder.asp" method="post">
				<tr align="middle" class="title">
					<td colspan="2" height="25">
						<b>修 改 送 货 资 料</b></td>
				</tr>
				<tr class="tdbg">
					<td align="right" width="15%">
						收货人：</td>
					<td><input type="text" name="contactman" maxlength="20" value="<%=rs("contactman")%>"/></td>
				</tr>
				<tr class="tdbg">
					<td align="right" width="15%">
						收货地址：</td>
					<td><input type="text" name="address" maxlength="120" value="<%=rs("address")%>"/>
					
					邮政编码：<input type="text" name="zipcode" maxlength="20"  size="10" value="<%=rs("zipcode")%>"/>
					</td>
				</tr>

				<tr class="tdbg">
					<td align="right" width="15%">
						联系电话：</td>
					<td>
						<input type="text" name="phone" maxlength="20" value="<%=rs("phone")%>"/>
						
						联系手机：<input type="text" name="mobile" maxlength="20" value="<%=rs("phone")%>"/></td>
				</tr>

				<tr class="tdbg">
					<td align="right" width="15%">
						电子邮件：</td>
					<td>
						<input type="text" name="email" maxlength="60" value="<%=rs("email")%>"/>
						联系QQ：<input type="text" name="qq" maxlength="20" value="<%=rs("qq")%>"/>
						</td>
				</tr>

				<tr class="tdbg">
					<td align="right" width="15%">
						发货方式：</td>
					<td>
					   <style>
					   	  .provincename{color:#ff6600}
						  .tocity{border:1px solid #006699;text-align:center;background:#C6E7FA;height:23px;width:130px;}
						  .showcity{position:absolute;background:#C6E7FA;border:#278BC6 1px solid;width:340px;display:none;height:230px;overflow-y:scroll;overflow-x:hidden;} 
						  .delivery{width:530px;padding:5px;border:1px solid #cccccc;background:#f1f1f1}
						  .jgxx{color:#ff3300}
						  .jgxx span{color:blue}
						 </style>
							 <script type="text/javascript">
								  function ajshowdata(city)
									{ 
											  $.get("../shop/ajax.getdate.asp",{city:escape(city),expressid:$("#DeliverType option:selected").val()},function(d){
											  var r=unescape(d).split('|');
											  if (r[0]=='error'){
											   alert(r[1]);
											   $("#jgxx").html('选择发往路线确定运费!');
											   $("#tocity").val('');
											  }else{ 
											   $("#jgxx").html(r[1]);
											   $("#tocity").val(city);
											   }
											  });
									} 
                                   $(document).ready(function(){
								   ajshowdata('<%=rs("tocity")%>');
								   })
							  </script>
						<div class="delivery">			  
						<%=GetDeliveryTypeStr(rs("DeliverType"),rs("tocity"))%>
						</div>
						</td>
				</tr>
				<tr class="tdbg">
					<td align="right" width="15%">
						付款方式：</td>
					<td>
						<%=GetPaymentTypeStr(rs("PaymentType"))%></td>
				</tr>
				<tr class="tdbg">
					<td align="right" width="15%">
						发票信息：</td>
					<td>
						<input type="radio" name="NeedInvoice" <%if rs("NeedInvoice")=0 then response.write " checked"%> value=0>不需要发票
						<input type="radio" name="NeedInvoice" <%if rs("NeedInvoice")=1 then response.write " checked"%> value="1">需要发票
						<br/>
						<textarea name="InvoiceContent" cols="40" rows="3"><%=rs("InvoiceContent")%></textarea>
						
						</td>
				</tr>
				<tr class="tdbg">
					<td align="right" width="15%">
						备注留言：</td>
					<td>
						<textarea name="Remark" cols="40" rows="3"><%=rs("Remark")%></textarea></td>
				</tr>

				<tr align="middle" class="tdbg">
					<td colspan="2" height="30" style="text-align:center">
						<input id="Action" name="Action" type="hidden" value="DoModifyInfoSave" /> <input id="ID" name="ID" type="hidden" value="<%=id%>" /> <input class="button" name="Submit" type="submit" value="确定保存修改" />&nbsp;<input class="button" name="Submit" onClick="javascript:parent.closeWindow();" type="button" value="关闭取消" /></td>
				</tr>
			</form>
		</table>
		<%
		End Sub
		
		'保存修改
      Sub DoModifyInfoSave()
		Dim ID:ID=KS.ChkClng(KS.G("id"))
		If id=0 Then KS.Die "error!"
		Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select top 1 * From KS_Order Where ID=" & ID,Conn,1,3
		If Not RS.Eof Then
		  RS("ContactMan")=KS.G("ContactMan")
		  RS("Address")=KS.G("Address")
		  RS("ZipCode")=KS.G("ZipCode")
		  RS("Phone")=KS.G("Phone")
		  RS("Mobile")=KS.G("Mobile")
		  RS("Email")=KS.G("Email")
		  RS("qq")=KS.G("qq")
		  RS("PaymentType")=KS.ChkClng(KS.G("PaymentType"))
		  RS("DeliverType")=KS.ChKClng(KS.G("DeliverType"))
		  RS("ToCity")=KS.G("ToCity")
		  RS("NeedInvoice")=KS.ChKClng(KS.G("NeedInvoice"))
		  RS("ToCity")=KS.G("ToCity")
		  RS("InvoiceContent")=KS.G("InvoiceContent")
		  RS("Remark")=KS.G("Remark")
		  RS.Update
		End If
		RS.Close :Set RS=Nothing
		KS.Die "<script>alert('恭喜，修改成功!');parent.location.reload();</script>"
  End Sub
		
  '付款方式
  Function GetPaymentTypeStr(PaymentType)
   Dim DiscountStr,SQL,I,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "select TypeID,TypeName,IsDefault,Discount from KS_PaymentType order by orderid",conn,1,1
   If Not RS.Eof Then
     SQL=RS.GetRows(-1)
   End IF
   RS.Close:Set RS=Nothing
   GetPaymentTypeStr="<select name='PaymentType'>"
   For I=0 To UBound(SQL,2)
     If SQL(3,I)<>100 Then
	  DiscountStr="折扣率 " & SQL(3,I) & "%"
	 Else
	  DiscountStr=""
	 End iF
     If trim(SQL(0,I))=trim(PaymentType) Then
    GetPaymentTypeStr=GetPaymentTypeStr& "<option value='" & SQL(0,I) & "' selected>"  &SQL(1,I) & " " & DiscountStr & "</option>"
	 Else
    GetPaymentTypeStr=GetPaymentTypeStr& "<option value='" & SQL(0,I) & "'>"  &SQL(1,I) & " " & DiscountStr & "</option>"
	End If
   Next
   GetPaymentTypeStr=GetPaymentTypeStr & "</select>"
  End Function
	
	 '发货方式
  Function GetDeliveryTypeStr(typeid,tocity)
   Dim j,rss,rsss
   Dim DiscountStr,SQL,I,RS


   Set RS=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "select TypeID,TypeName,IsDefault from KS_DeliveryType order by orderid,TypeID",conn,1,1
   If Not RS.Eof Then
     SQL=RS.GetRows(-1)
   End IF
   RS.Close:Set RS=Nothing
   GetDeliveryTypeStr="<strong>快递公司：</strong><select name='DeliverType' id='DeliverType'>"
   For I=0 To UBound(SQL,2)
     If trim(typeid)=trim(sql(0,i)) Then
    GetDeliveryTypeStr=GetDeliveryTypeStr& "<option value='" & SQL(0,I) & "' selected>"  &SQL(1,I) & "</option>"
	 Else
    GetDeliveryTypeStr=GetDeliveryTypeStr& "<option value='" & SQL(0,I) & "'>"  &SQL(1,I) & "</option>"
	End If
   Next
   GetDeliveryTypeStr=GetDeliveryTypeStr & "</select>"
   if tocity="" then tocity="选择送货地点"

   GetDeliveryTypeStr=GetDeliveryTypeStr & "<br/> <input type=""hidden"" name=""tocity"" id=""tocity""/> <span style='position:relative'><input class=""tocity"" style='text-align;left' name='' id='choosecity' type='button' value='" & tocity & "'  onclick=""showprovn.style.display='block';if(this.getBoundingClientRect().top>300){showprovn.style.top=(this.offsetHeight-showprovn.offsetHeight)}else{showprovn.style.top='0'}""><span id='showprovn' onclick=""this.style.display='none'"" class='showcity'>"&_
			 "<table width='92%' align='center' border='0' cellspacing='0' cellpadding='0'>"&vbcrlf
			        dim pxml,node,pnode
			        set rss=conn.execute("select id,City,parentid from KS_Province order by orderid asc,id")
					if not rss.eof then
					  set pxml=KS.RsToXml(rss,"row","")
					end if
					rss.close  : Set RSS=Nothing
					If IsObject(Pxml) Then
	  				 For Each Node In pxml.DocumentElement.SelectNodes("row[@parentid=0]")
					    GetDeliveryTypeStr=GetDeliveryTypeStr&"<tr><td colspan='5' class='provincename'><strong>" & Node.SelectSingleNode("@city").text &"</td></tr>"
						j=1
						For Each pnode in Pxml.DocumentElement.SelectNodes("row[@parentid=" & Node.SelectSingleNode("@id").text & "]")
						IF (j MOD 4) = 1 THEN GetDeliveryTypeStr=GetDeliveryTypeStr&"<tr>"&vbcrlf
						GetDeliveryTypeStr=GetDeliveryTypeStr&"<td id='ccity' onclick=""choosecity.value=this.innerHTML;ajshowdata(this.innerHTML)"" style='cursor:hand' onmouseover=""this.style.color='red'"" onmouseout=""this.style.color=''"">"&pnode.selectsinglenode("@city").text&"</td>"&vbcrlf
						if (j mod 4)=0 then GetDeliveryTypeStr=GetDeliveryTypeStr&"</tr>"&vbcrlf
						j=j+1
						Next
						
					 Next
					End If
 
			        
					 
			 GetDeliveryTypeStr=GetDeliveryTypeStr&"</table>"&vbcrlf&_
			"</span></span>"&_
		" <span id='jgxx' class='jgxx'>选择送货路线确定运费！</span>"&vbcrlf


  End Function
	
  '修改总价
  Sub ModifyTotalPrice()
          If KS.ReturnPowerResult(0, "M520013")=false Then
		    response.write "<script>alert('对不起，您没有权限修改订单!');parent.closeWindow();</script>"
			response.end
		  End If
		  dim id:id=ks.chkclng(request("id"))
		  dim price:price=request("price")
		  dim oprice:oprice=request("oprice")
		  if id=0 then
		    response.write "<script>alert('参数出错!');</script>"
			response.end
		  end if
		  if not isnumeric(price) then
		    response.write "<script>alert('输入的价格不对,请输入正确的数字!');</script>"
			response.end
		  end if
		  if oprice=price then
		    response.write "<script>alert('价格与修改前一样,没有更新!');</script>"
			response.end
		  end if
		  conn.execute("update ks_order set moneytotal=" & price  & " where id=" & id)
		  response.write "<script>alert('恭喜,订单总价修改成功!');parent.location.replace(document.referrer);</script>"
  End Sub	

  '更新订单价格
  sub updateorderprice(orderid)
          dim totalrealprice:totalrealprice=0
		  Dim totalweight:totalweight=0
		  dim rs:set rs=server.CreateObject("adodb.recordset")
		  rs.open "select i.*,p.weight from ks_orderitem i left join ks_product p on i.proid=p.id where i.orderid='" & orderid & "'",conn,1,1
		  do while not rs.eof
		    totalrealprice=totalrealprice+Round(rs("totalprice"),2)
			if isnumeric(rs("weight")) then
		    totalweight=totalweight+Round(rs("weight")*rs("amount"),2)
			end if
		  rs.movenext
		  loop
		  rs.close
		  
		  if totalrealprice<>0 then
		    conn.execute("update ks_order set weight=" & totalweight & " where orderid='" & orderid & "'")
		    rs.open "select top 1 * from ks_order where orderid='" & orderid & "'",conn,1,3
			if not rs.eof then
			   rs("moneygoods")=totalrealprice
			   Dim TaxRate:TaxRate=KS.Setting(65)
			   Dim IncludeTax:IncludeTax=KS.Setting(64)
			   Dim TaxMoney,RealMoneyTotal,Freight
			   Freight=KS.GetFreight(RS("DeliverType"),RS("ToCity"),RS("weight"),"")
			   If IncludeTax=1 Or rs("NeedInvoice")=0 Then TaxMoney=1 Else TaxMoney=1+Taxrate/100
				'总金额 = (总价*付费方式折扣+运费)*(1+税率)
				RealMoneyTotal=Round((totalrealprice*KS.ReturnPayment(rs("PaymentType"),1)/100+Freight*TaxMoney),2)
				RS("Charge_Deliver")=Freight
			  rs("NoUseCouponMoney")=RealMoneyTotal
			  if rs("CouponUserID")<>0 then
			     'dim facevalue:facevalue=conn.execute("select facevalue from KS_ShopCoupon where id=" &rs("CouponUserID"))(0) 
			   	' If FaceValue>0 Then
				   RealMoneyTotal=Round(RealMoneyTotal-rs("usecouponmoney"),2)
				' End If
			  end if
			  rs("MoneyTotal")=RealMoneyTotal

  
			   rs.update
			end if
			rs.close
		  end if
		  set rs=nothing
  end sub		
		
  '修改指定价
  sub ModifyPrice()
           If KS.ReturnPowerResult(0, "M520013")=false Then
		    response.write "<script>alert('对不起，您没有权限修改订单价格!');parent.closeWindow();</script>"
			response.end
		   End If
		  dim id:id=ks.chkclng(request("id"))
		  dim price:price=request("price")
		  dim orderid:orderid=ks.g("orderid")
		  dim oprice:oprice=request("oprice")
		  if id=0 then
		    response.write "<script>alert('参数出错!');</script>"
			response.end
		  end if
		  if not isnumeric(price) then
		    response.write "<script>alert('输入的价格不对,请输入正确的数字!');</script>"
			response.end
		  end if
		  if oprice=price then
		    response.write "<script>alert('价格与修改前一样,没有更新!');</script>"
			response.end
		  end if
		  dim rs:set rs=server.createobject("adodb.recordset")
		  rs.open "select top 1 * from ks_orderitem where id=" &id,conn,1,3
		  if not rs.eof then
		     rs("realprice")=price
			 rs("totalprice")=price * rs("amount")
			 rs.update
		  end if
		  rs.close
		  set rs=nothing
		  call updateorderprice(orderid)
		  response.write "<script>alert('恭喜,指定价修改成功!');parent.location.replace(document.referrer);</script>"
		end sub
		
	
		
		Sub OrderList()
%>
  <div style="border:1px #cccccc dashed;overflow:hidden"></div>
  <table width="100%">
    <tr>
      <td align=left>您现在的位置：<a href="KS.ShopOrder.asp">订单处理</a>&nbsp;&gt;&gt;&nbsp;
	  <%
	     Dim SearchTypeStr,Keyword
		 Keyword=KS.G("Keyword")
	    Select Case SearchType
	    Case 0
		SearchTypeStr= "所有订单"
		Case 1
		SearchTypeStr= "24小时之内的新订单"
		Case 2
		SearchTypeStr= "最近10天内的新订单"
		Case 3
		SearchTypeStr= "最近一月内的新订单"
		Case 4
		SearchTypeStr="未确认的订单"
		Case 5
		SearchTypeStr="未付款的订单"
		Case 6
		SearchTypeStr="未付清的订单"
		Case 7
		SearchTypeStr="未送货的订单"
		Case 8
		SearchTypeStr="未签收的订单"
		Case 9
		SearchTypeStr="未开发票的订单"
		Case 10
		   Select Case  KS.ChkClng(KS.G("Field"))
		    Case 1:SearchTypeStr="订单编号含有<font color=red>""" & KeyWord & """</font>"
		    Case 2:SearchTypeStr="收货人含有<font color=red>""" & KeyWord & """</font>"
		    Case 3:SearchTypeStr="用户名含有<font color=red>""" & KeyWord & """</font>"
		    Case 4:SearchTypeStr="联系地址含有<font color=red>""" & KeyWord & """</font>"
		    Case 5:SearchTypeStr="联系电话含有<font color=red>""" & KeyWord & """</font>"
		    Case 6:SearchTypeStr="下单时间含有<font color=red>""" & KeyWord & """</font>"
		    Case 7:SearchTypeStr="推荐人为<font color=red>""" & KeyWord & """</font>"
		   End Select
		Case 11
		SearchTypeStr="未结清的订单"
		Case 12
		SearchTypeStr="已结清的订单"
		End Select
		Response.Write SearchTypeStr
	  %>
	  </td>
    </tr>
  </table>
  <div style="border:1px #cccccc dashed;overflow:hidden"></div>
  <table cellSpacing=0 cellPadding=0 width="100%" border=0>
    <tr>
<FORM name=myform onSubmit="return confirm('确定要删除选定的订单吗？');" action=KS.ShopOrder.asp method=post>
      <td>
        <table cellSpacing="0" cellPadding="0" width="100%" border=0>
          <tr class=sort align=middle>
            <td width=30>选中</td>
            <td width=110>订单编号</td>
            <td nowrap="nowrap">客户</td>
            <td>用户名</td>
            <td width=120>下单时间</td>
            <td width=60>总金额</td>
            <td width=60>应付金额</td>
            <td width=60>收款金额</td>
            <td width=30>需要<br>发票</td>
            <td width=30>已开<br>发票</td>
            <td width=60>订单状态</td>
            <td width=60>付款状态</td>
            <td width=60>物流状态</td>
          </tr>
		  <%
		  	MaxPerPage=20
			If KS.G("page") <> "" Then
				  CurrentPage = KS.ChkClng(KS.G("page"))
			Else
				  CurrentPage = 1
			End If
			
			SqlParam="1=1"
			If SearchType<>"0" Then
			  Select Case SearchType
			   Case 1 SqlParam=SqlParam &" And datediff(" & DataPart_H & ",inputtime," & SqlNowString & ")<25"
			   Case 2 SqlParam=SqlParam &" And datediff(" & DataPart_D & ",inputtime," & SqlNowString & ")<=10"
			   Case 3 SqlParam=SqlParam &" And datediff(" & DataPart_D & ",inputtime," & SqlNowString & ")<=30"
			   Case 4:SqlParam=SqlParam &" And Status=0"
			   Case 5:SqlParam=SqlParam &" And MoneyReceipt=0"
			   Case 6:SqlParam=SqlParam &" And MoneyReceipt<=MoneyTotal"
			   Case 7:SqlParam=SqlParam &" And DeliverStatus=0"
			   Case 8:SqlParam=SqlParam &" And DeliverStatus=1"
			   Case 9:SqlParam=SqlParam &" And NeedInvoice=1 And Invoiced=0"
			   Case 10
			      Select Case KS.ChkClng(KS.G("Field"))
				   Case 1 SqlParam=SqlParam &" And OrderID Like '%" & Keyword & "%'"
				   Case 2 SqlParam=SqlParam &" And ContactMan Like '%" & Keyword & "%'"
				   Case 3 SqlParam=SqlParam &" And UserName Like '%" & Keyword & "%'"
				   Case 4 SqlParam=SqlParam &" And Address Like '%" & Keyword & "%'"
				   Case 5 SqlParam=SqlParam &" And Phone Like '%" & Keyword & "%'"
				   Case 6 SqlParam=SqlParam &" And InputTime Like '%" & Keyword & "%'"
				   Case 7 SqlParam=SqlParam & " and UserName in(select username from ks_user where AllianceUser='" & KeyWord & "')"
				  End Select
			   Case 11:SqlParam=SqlParam &" And status=1"
			   Case 12:SqlParam=SqlParam &" And status=2"
			  End Select
			End If

		   Set RS=Server.CreateObject("ADODB.RECORDSET")
		   SqlStr="Select * From KS_Order where " & SqlParam & " order by inputtime desc"
		   RS.Open SqlStr ,Conn,1,1
		   If RS.Eof And RS.Bof Then
		    Response.Write "<tr class=list onmouseover=""this.className='listmouseover'"" onmouseout=""this.className='list'"" align=middle><td height='30' colspan=12>找不到" & SearchTypeStr & "!</td></tr>"
		  Else
		  	               totalPut = RS.RecordCount
							If CurrentPage < 1 Then	CurrentPage = 1
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (totalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
							If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
							End If
							Call showContent()
		  End If
			 RS.Close:Set RS=Nothing
		%>
        <table cellSpacing=0 cellPadding=0 width="100%" border=0>
          <tr>
            <td height=30>
              &nbsp;<Input id=chkAll onclick=CheckAll(this.form) type=checkbox value=checkbox name=chkAll> 选中本页显示的所有订单
  <Input id=Action type=hidden value=DelOrder name=Action> 
              <Input type=submit value="删除选定的订单" class="button" name=Submit>
		   </td>
		   <td>
		   <%
		   	  '显示分页信息
			  Call KS.ShowPageParamter(totalPut, MaxPerPage, "KS.ShopOrder.asp", True, "条订单", CurrentPage, KS.QueryParam("page"))
		   %>
		   </td>
          </tr>
        </table>
		</FORM>
		<div class="attention">
		<font color=red>说明：为便于销售统计已结清或已收到汇款(包括仅收到预付款)的订单不能删除。</font>
		</div>
      </td>
    </tr>
  </table>
</body>
<html>
		<%
		End Sub
		
		Sub ShowContent()
		      Dim I
			  Do While Not RS.Eof 
		   %>
			  <tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" align=middle>
				<td class='splittd' height='25'><Input id=ID type=checkbox value="<%=rs("id")%>" name="ID"></td>
				<td class='splittd'><a href="KS.ShopOrder.asp?Action=ShowOrder&ID=<%=RS("ID")%>"><%=RS("OrderID")%></a></td>
				<td class='splittd'><%=RS("ContactMan")%></td>
				<td class='splittd'><%=RS("UserName")%></td>
				<td class='splittd'><%=RS("InputTime")%></td>
				<td class='splittd' align=right>￥<%=RS("NoUseCouponMoney")%>元</td>
				<td  class='splittd'align=right>￥<%=RS("MoneyTotal")%>元</td>
				<td  class='splittd'align=right><font color=red><%=rs("MoneyReceipt")%></font></td>
				<td class='splittd'>
				<%If RS("NeedInvoice")=1 Then
				  Response.Write "<Font color=red>√</font>"
				  Else
				   Response.Write "&nbsp;"
				  End If
				  %>
				</td>
				<td class='splittd'>
				<%
				if RS("NeedInvoice")=1 Then
				  If RS("Invoiced")=1 Then
				   Response.Write "<font color=green>√</font>"
				  Else
				   Response.Write "<font color=red>×</font>"
				  End If
				Else
				  Response.Write "&nbsp;"
				End If
				 %>
				</td>
				<td class='splittd'>
				<%If RS("Status")=0 Then
				  Response.Write "<font color=red>等待确认</font>"
				  ElseIf RS("Status")=1 Then
				  Response.WRITE "<font color=green>已经确认</font>"
				  ElseIf RS("Status")=2 Then
				  Response.Write "<font color=#a7a7a7>已结清</font>"
				  ElseIf RS("Status")=3 Then
				  Response.Write "<font color=#a7a7a7>无效订单</font>"
				  End If
				%>
				  </td>
				<td class='splittd'>
				<%If RS("MoneyReceipt")<=0 Then
				   Response.Write "<font color=red>等待汇款</font>"
				  ElseIf RS("MoneyReceipt")<RS("MoneyTotal") Then
				   Response.WRITE "<font color=blue>已收定金</font>"
				  Else
				   Response.Write "<font color=green>已经付清</font>"
				  End If
				  %></td>
				<td class='splittd'>
				<% If RS("DeliverStatus")=0 Then
				 Response.Write "<font color=red>未发货</font>"
				 ElseIf RS("DeliverStatus")=1 Then
				  Response.Write "<font color=blue>已发货</font>"
				 ElseIf RS("DeliverStatus")=2 Then
				  Response.Write "<font color=green>已签收</font>"
				 ElseIf RS("DeliverStatus")=3 Then
				  Response.Write "<font color=#ff6600>退货</font>"
				 End If
				 %></td>
			  </tr>
			  <%
			    PageTotalMoney1=PageTotalMoney1+RS("MoneyTotal")
				PageTotalMoney2=PageTotalMoney2+RS("MoneyReceipt")
				I = I + 1
				RS.MoveNext
				If I >= MaxPerPage Then Exit Do
			  Loop
		  %>
          <tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" align=middle>
            <td  class='splittd' align=right colSpan=6><B>本页合计：</B></td>
            <td  class='splittd' align=right><%=PageTotalMoney1%></td>
            <td  class='splittd' align=right><%=PageTotalMoney2%></td>
            <td  class='splittd' colSpan=5>&nbsp;</td>
          </tr>
          <tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" align=middle>
            <td class='splittd' align=right colSpan=6><B>本次查询合计：</B></td>
            <td class='splittd' align=right><%=Conn.execute("Select Sum(MoneyTotal) From KS_Order where " & SqlParam)(0)%></td>
            <td class='splittd' align=right><%=Conn.execute("Select Sum(MoneyReceipt) From KS_Order where " & SqlParam)(0)%></td>
            <td class='splittd' colSpan=5>&nbsp;</td>
          </tr>
          <tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" align=middle>
            <td class='splittd' align=right colSpan=6><B>总计金额：</B></td>
            <td class='splittd' align=right><%=Conn.execute("Select Sum(MoneyTotal) From KS_Order")(0)%></td>
            <td class='splittd' align=right><%=Conn.execute("Select Sum(MoneyReceipt) From KS_Order")(0)%></td>
            <td class='splittd' colSpan=5>&nbsp;</td>
          </tr>
        </table>
		<%End Sub
		
		Sub ShowOrder()
		 Dim ID:ID=KS.ChkClng(KS.G("ID"))
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select top 1 * from ks_order where id=" & ID ,conn,1,1
		 IF RS.Eof And RS.Bof Then
		   RS.Close:Set RS=Nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   response.end
		 End If
		 
		response.write "<br>"
		response.write OrderDetailStr(RS,1)
		%>

        <br>
   <div align=center> 
           <% 
  			 '==================================三期开发代码=========================================
			 If RS("Status")=3 Then
			   response.write "本订单在指定时间内没有付款作废!"
			 Else
			 '=====================================================================================

		   If RS("Status")<>2 Then%>   
			 <% IF RS("Status")=0 Then%>
			 <input type='button' class='button' name='Submit' value='确认订单' onClick="javascript:if(confirm('请仔细检查此订单的所有信息，确认后将发送一封站内短信和邮件通知客户!')){window.location.href='KS.ShopOrder.asp?Action=OrderConfirm&ID=<%=RS("ID")%>';}">&nbsp;&nbsp;
			 <%ElseIf RS("Status")=1 And RS("MoneyReceipt")=0 Then%>
			 <input type='button' class='button' name='Submit' value='删除订单' onClick="javascript:if(confirm('确定删除此订单吗!')){window.location.href='KS.ShopOrder.asp?Action=DelOrder&ID=<%=RS("ID")%>';}">&nbsp;&nbsp;
			 <%End iF%>
			 <%
			 If RS("MoneyReceipt")<RS("MoneyTotal") Then%>
			 <input type='button'class='button'  name='Submit' value='银行汇款支付' onClick="window.location.href='KS.ShopOrder.asp?Action=BankPay&ID=<%=RS("id")%>'">&nbsp;
			 <%Else%>
			 <input type='button' class='button' name='Submit' value=' 退款 ' onClick="window.location.href='KS.ShopOrder.asp?Action=BankRefund&ID=<%=RS("id")%>'">&nbsp;
			 <%End IF%>
			 <%If RS("NeedInvoice")=1 And RS("Invoiced")=0 Then%>
			 <input type='button' class='button' name='Submit' value=' 开发票 ' onClick="window.location.href='KS.ShopOrder.asp?Action=Invoice&ID=<%=RS("ID")%>'">&nbsp;
			 <%End IF%>
			 <%If RS("Status")=1 Then%>
			 <input type='button' class='button' name='Submit' value='客户已签收' onClick="if(confirm('确定客户已收到货了吗?')){window.location.href='KS.ShopOrder.asp?Action=ClientSignUp&ID=<%=RS("ID")%>';}">&nbsp;
			 <%End If
			 If RS("MoneyReceipt")>=RS("MoneyTotal") And RS("Status")<>0 And RS("DeliverStatus")<>0 Then
			 %>
			 <input type='button' class='button' name='Submit' value='结清订单' onClick="if(confirm('订单一旦结算，该订单就不可进行任何操作，确定结清订单吗?')){window.location.href='KS.ShopOrder.asp?Action=FinishOrer&ID=<%=RS("ID")%>';}">&nbsp;

			 <%
			 End if
			 IF RS("DeliverStatus")=0 Then%>
			 <input type='button' class='button' name='Submit' value=' 发货 ' onClick="window.location.href='KS.ShopOrder.asp?Action=DeliverGoods&ID=<%=rs("id")%>'">&nbsp;
			 <%ElseIf RS("DeliverStatus")<>3 Then%>
			 <input type='button' class='button' name='Submit' value=' 客户退货 ' onClick="window.location.href='KS.ShopOrder.asp?Action=BackGoods&ID=<%=rs("id")%>'">&nbsp;
			 <%End If%>
			 <%End If%>
			 <input type='button' class='button' name='Submit' value=' 支付货款给卖方 ' onClick="window.location.href='KS.ShopOrder.asp?Action=PayMoney&ID=<%=rs("id")%>'">&nbsp;
			 <%
			End If

			 %>
			 <input type='button' class='button' name='Submit' value='打印订单' onClick="window.location.href='KS.ShopOrder.asp?Action=PrintOrder&ID=<%=RS("ID")%>'">
			 &nbsp;<input type='button' class='button' name='Submit' value='取消返回' onClick="javascript:history.back();">
			</div>
</body></html>
		<%
		 RS.Close:Set RS=Nothing
		End Sub
		
		'返回订单详细信息
		Function  OrderDetailStr(RS,flag)
		 OrderDetailStr="<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'> "&vbcrlf
		 OrderDetailStr=OrderDetailStr & "<tr align='center' class='title'>    <td height='22'><b>订 单 信 息</b>（订单编号：" & RS("ORDERID") & "）</td>  </tr>"&vbcrlf
		 OrderDetailStr=OrderDetailStr & "<tr>" & vbcrlf
		 OrderDetailStr=OrderDetailStr & " <td height='25'>" &vbcrlf
		 OrderDetailStr=OrderDetailStr & "  <table width='100%'  border='0' cellpadding='2' cellspacing='0'> "   & vbcrlf
		 OrderDetailStr=OrderDetailStr & "    <tr class='tdbg'>"
		 OrderDetailStr=OrderDetailStr & "	         <td width='18%'>客户姓名：<font color='red'>" & RS("Contactman") & "</td>      "
		 OrderDetailStr=OrderDetailStr & "			 <td width='20%'>用 户 名：<font color='red'>" & rs("username") & "</td> " &vbcrlf
		OrderDetailStr=OrderDetailStr & "			 <td width='20%'>代 理 商：</td>"
		OrderDetailStr=OrderDetailStr & "			 <td width='18%'>购买日期：<font color='red'>" & formatdatetime(rs("inputtime"),2) & "</font></td>" & vbcrlf
		OrderDetailStr=OrderDetailStr & "			 <td width='24%'>下单时间：<font color='red'>" & rs("inputtime") & "</font></td>" & vbcrlf
		OrderDetailStr=OrderDetailStr & "	</tr>"
		OrderDetailStr=OrderDetailStr & "	<tr class='tdbg'> "      
		OrderDetailStr=OrderDetailStr & "	  <td width='18%'>需要发票："
			    If RS("NeedInvoice")=1 Then
				  OrderDetailStr=OrderDetailStr & "<Font color=red>√</font>"
				  Else
				  OrderDetailStr=OrderDetailStr & "<font color=red>×</font>"
				  End If
		OrderDetailStr=OrderDetailStr & "		 </td> "       
		OrderDetailStr=OrderDetailStr & "	 <td width='20%'>已开发票："	
				  If RS("Invoiced")=1 Then
				   OrderDetailStr=OrderDetailStr & "<font color=green>√</font>"
				  Else
				   OrderDetailStr=OrderDetailStr & "<font color=red>×</font>"
				  End If
		OrderDetailStr=OrderDetailStr & "	</td> "
		OrderDetailStr=OrderDetailStr & "	<td width='20%'>订单状态："	
			if RS("Status")=0 Then
				 OrderDetailStr=OrderDetailStr & "<font color=red>等待确认</font>"
				  ElseIf RS("Status")=1 Then
				 OrderDetailStr=OrderDetailStr & "<font color=green>已经确认</font>"
				  ElseIf RS("Status")=2 Then
				 OrderDetailStr=OrderDetailStr & "<font color=#a7a7a7>已结清</font>"
				  End If
		OrderDetailStr=OrderDetailStr & "	</td>"
		OrderDetailStr=OrderDetailStr & "	  <td width='18%'>付款情况："	
			     If RS("MoneyReceipt")<=0 Then
				   OrderDetailStr=OrderDetailStr & "<font color=red>等待汇款</font>"
				  ElseIf RS("MoneyReceipt")<RS("MoneyTotal") Then
				   OrderDetailStr=OrderDetailStr & "<font color=blue>已收定金</font>"
				  Else
				  OrderDetailStr=OrderDetailStr & "<font color=green>已经付清</font>"
				  End If

       OrderDetailStr=OrderDetailStr & "</td>"
	   OrderDetailStr=OrderDetailStr & "        <td width='24%'>物流状态："
				if RS("DeliverStatus")=0 Then
				 OrderDetailStr=OrderDetailStr & "<font color=red>未发货</font>"
				 ElseIf RS("DeliverStatus")=1 Then
				  OrderDetailStr=OrderDetailStr & "<font color=blue>已发货</font>"
				 ElseIf RS("DeliverStatus")=2 Then
				  OrderDetailStr=OrderDetailStr & "<font color=blue>已签收</font>"
				 ElseIf RS("DeliverStatus")=3 Then
				  OrderDetailStr=OrderDetailStr & "<font color=#ff6600>退货</font>"
				 End If
	OrderDetailStr=OrderDetailStr & "		</td></tr>    </table> "
    OrderDetailStr=OrderDetailStr & " </td>  </tr> " 
	OrderDetailStr=OrderDetailStr & "   <tr align='center'>"
	OrderDetailStr=OrderDetailStr & "       <td height='25'>"
	OrderDetailStr=OrderDetailStr & "	   <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>"
	OrderDetailStr=OrderDetailStr & "	           <tr class='tdbg'>"
	OrderDetailStr=OrderDetailStr & "			             <td width='12%' align='right'>收货人姓名：</td>"
	OrderDetailStr=OrderDetailStr & "						 <td width='38%'>" & rs("contactman") & "</td>"
	OrderDetailStr=OrderDetailStr & "						 <td width='12%' align='right'>联系电话：</td> "      
	OrderDetailStr=OrderDetailStr & "						 <td width='38%'>" & rs("phone") & "</td>"
	OrderDetailStr=OrderDetailStr & "				</tr>"
	OrderDetailStr=OrderDetailStr & "				<tr class='tdbg' valign='top'>"
	OrderDetailStr=OrderDetailStr & "				          <td width='12%' align='right'>收货人地址：</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" & rs("address") & "</td>"          
	OrderDetailStr=OrderDetailStr & "						  <td width='12%' align='right'>邮政编码：</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" &rs("zipcode") & "</td>"
	OrderDetailStr=OrderDetailStr & "				</tr>  "      
	OrderDetailStr=OrderDetailStr & "				<tr class='tdbg'> "         
	OrderDetailStr=OrderDetailStr & "				          <td width='12%' align='right'>收货人邮箱：</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" & rs("email") & " 联系QQ: " & rs("qq") & "</td> "         
	OrderDetailStr=OrderDetailStr & "						  <td width='12%' align='right'>收货人手机：</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" & rs("mobile") & "</td>       "
	OrderDetailStr=OrderDetailStr & "			   </tr>"        
	OrderDetailStr=OrderDetailStr & "			   <tr class='tdbg'> "         
	OrderDetailStr=OrderDetailStr & "			              <td width='12%' align='right'>付款方式：</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" & KS.ReturnPayMent(rs("PaymentType"),0) & "</td>       "   
	OrderDetailStr=OrderDetailStr & "						  <td width='12%' align='right'>快递公司：</td>"          
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" 
	
	  dim rst,foundexpress
	  Set RST=Server.CreateObject("ADODB.RECORDSET")
	 RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and a.tocity like '%"&rs("tocity")&"%'",conn,1,1
	 If RST.Eof Then
	    foundexpress=false
	 Else
	    foundexpress=true
	OrderDetailStr=OrderDetailStr & "<span style='color:green'>" & rst("typename") & "</span> 首重<span style='color:#ff6600'>"&rst("fweight")&"kg/"&rst("carriage")&"元</span>  续重<span style='color:#ff6600'>"&rst("W_fee")&"kg/"&rst("C_fee")&"元</span>"
	 End If
	 RST.Close
	 If foundexpress=false Then
	  If DataBaseType=1 Then
	  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and (convert(varchar(200),tocity)='' or a.tocity is null)",conn,1,1
	  Else
	  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and (a.tocity='' or a.tocity is null)",conn,1,1
	  End If
	  if rst.eof then
	  else
	OrderDetailStr=OrderDetailStr & "<span style='color:green'>" & rst("typename") & "</span> 首重<span style='color:#ff6600'>"&rst("fweight")&"kg/"&rst("carriage")&"元</span>  续重<span style='color:#ff6600'>"&rst("W_fee")&"kg/"&rst("C_fee")&"元</span>"
	  end if
	  rst.close
	 End If
	 set rst=nothing
	
	
	OrderDetailStr=OrderDetailStr & " 发往<span style='color:red'>" & rs("tocity") & "</span></td>"
	OrderDetailStr=OrderDetailStr & "				</tr> "       
	OrderDetailStr=OrderDetailStr & "				<tr class='tdbg' valign='top'>  "        
	OrderDetailStr=OrderDetailStr & "				          <td width='12%' align='right'>发票信息：</td>"          
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>"
	 If RS("Invoiced")=1 Then OrderDetailStr=OrderDetailStr & rs("InvoiceContent") &"</td>"
    OrderDetailStr=OrderDetailStr & "						 <td width='12%' align='right'>备注/留言：</td>"          
	OrderDetailStr=OrderDetailStr & "							<td width='38%'>" & rs("Remark") & "</td>       "
	OrderDetailStr=OrderDetailStr & "				 </tr>  "  
	OrderDetailStr=OrderDetailStr & "				 </table>"
	if flag=1 And KS.ReturnPowerResult(0, "M520013") then
	 OrderDetailStr=OrderDetailStr & "<div style='text-align:left'><input type='button' onclick=""modifyInfo(" & rs("id") & ")"" class='button' value='修改配送信息'/> <input type='button' onclick=""modifyproduct('" & rs("orderid") & "')"" class='button' value='修改/添加商品'/> <input type='button' onclick=""modifytotalprice(" & rs("id") & "," & rs("moneytotal") &")"" class='button' value='修改订单总价'/></div>"
	End If
	OrderDetailStr=OrderDetailStr & "			</td>  "
	OrderDetailStr=OrderDetailStr & "		</tr>  "
	
	OrderDetailStr=OrderDetailStr & "		<tr><td>"
	OrderDetailStr=OrderDetailStr & "		<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'> "
	OrderDetailStr=OrderDetailStr & "		  <tr align='center' class='title' height='25'>  "  
	OrderDetailStr=OrderDetailStr & "		   <td><b>商 品 名 称</b></td> "   
	OrderDetailStr=OrderDetailStr & "		   <td width='45'><b>单位</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='55'><b>数量</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='65'><b>原价</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='65'><b>实价</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='65'><b>指定价</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='85'><b>金 额</b></td>   " 
	OrderDetailStr=OrderDetailStr & "		   <td width='65'><b>服务期限</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='45'><b>备注</b></td>  "
	OrderDetailStr=OrderDetailStr & "		  </tr> "
			 Dim attributecart,TotalPrice:totalprice=0
			 Dim RSI:Set RSI=Server.CreateObject("ADODB.RECORDSET")
			   RSI.Open "Select * From KS_OrderItem Where SaleType<>5 and SaleType<>6 and OrderID='" & RS("OrderID") & "' order by ischangedbuy,id",conn,1,1
			   If RSI.Eof Then
			     RSI.Close:Set RSI=Nothing
				' OrderDetailStr=OrderDetailStr & "<tr><td align='center' colspan='10'>记录已被删除</td></tr> "
			  Else
			   Do While Not RSI.Eof
			   If Conn.execute("select top 1 title from ks_product where id=" & rsi("proid")).eof Then
			   		OrderDetailStr=OrderDetailStr & "	  <tr valign='middle' class='tdbg' height='20'>"    
					OrderDetailStr=OrderDetailStr & "	  <td colspan='9'>该商品已被删除！</td>"   
					OrderDetailStr=OrderDetailStr & "	  </tr>"   
			   Else
			  attributecart=rsi("attributecart")
			  if not ks.isnul(attributecart) then attributecart="<br/><font color=#888888>" & attributecart & "</font>"
			  Dim RSP:Set RSP=Server.CreateObject("ADODB.RECORDSET")
			  RSP.Open "Select top 1 I.Title,I.Unit,I.IsLimitBuy,I.LimitBuyPrice,L.LimitBuyPayTime From KS_Product I Left Join KS_ShopLimitBuy L On I.LimitBuyTaskID=L.Id  Where I.ID=" & RSI("ProID"),conn,1,1
			  dim title,unit,LimitBuyPayTime
			  If Not RSP.Eof Then
				  title=rsp("title")
				  Unit=rsp("unit")
				  If RSI("IsChangedBuy")=1 Then 
				   title=title &"(换购)"
				  Else
				     If RSP("LimitBuyPayTime") Then
				  	   If LimitBuyPayTime="" Then
					   LimitBuyPayTime=RSP("LimitBuyPayTime")
					   ElseIf LimitBuyPayTime>RSP("LimitBuyPayTime") Then
						LimitBuyPayTime=RSP("LimitBuyPayTime")
					   End If
					 End If
				  End If
				  If RSI("IsLimitBuy")="1" Then  title=title & "<span style='color:green'>(限时抢购)</span>"
				  If RSI("IsLimitBuy")="2" Then title=title & "<span style='color:blue'>(限量抢购)</span>"
			  End If
			  RSP.Close:Set RSP=Nothing
			  
		OrderDetailStr=OrderDetailStr & "	  <tr valign='middle' class='tdbg' height='20'>"    
		OrderDetailStr=OrderDetailStr & "	   <td width='*'><a href='" & DomainStr & "item/show.asp?m=5&d=" & RSi("proid") & "' target='_blank'>" & title & "</a>" & attributecart & "</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td width='45' align=center>"& Unit & "</td>               <td width='55' align='center'>" & rsi("amount") &"</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td width='65' align='center'>" & formatnumber(rsi("price_original"),2) & "</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td width='65' align='center'>" & formatnumber(rsi("realprice"),2) & "</td>    "
		
		if flag=1 then
			OrderDetailStr=OrderDetailStr & "	   <td width='65' align='center'>" & formatnumber(rsi("realprice"),2) 
			If RSI("IsChangedBuy")<>1 And RSI("IsLimitBuy")<>"1" And RSI("IsLimitBuy")<>"2" Then
			OrderDetailStr=OrderDetailStr& " <a href=""javascript://"" onclick=""modifyPrice(event,'" & title & "','" & rs("orderid") & "'," & rsi("id")&"," & rsi("realprice") & ")""><font color=blue>改</font></a>"
			End If
			OrderDetailStr=OrderDetailStr & "</td>    "
		else
		OrderDetailStr=OrderDetailStr & "	   <td width='65' align='center'>" & formatnumber(rsi("realprice"),2) & "</td>    "
		end if
		OrderDetailStr=OrderDetailStr & "	   <td width='85' align='right'>" & formatnumber(rsi("realprice")*rsi("amount"),2) & "</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td width='65' align=center>" & rsi("ServiceTerm") & "</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td align=center width='45'>" & rsi("Remark") & "</td>  "
		OrderDetailStr=OrderDetailStr & "	   </tr> " 
		
		'==================================三期开发代码=========================================
		OrderDetailStr=OrderDetailStr & GetBundleSalePro(TotalPrice,RSI("ProID"),RSI("OrderID"))  '取得捆绑销售商品
		'=========================================================================================
		     end if
			    TotalPrice=TotalPrice+ rsi("realprice")*rsi("amount")
			    rsi.movenext
			  loop
			  rsi.close:set rsi=nothing
			End If
			
			
			OrderDetailStr=OrderDetailStr & GetPackage(TotalPrice,RS("OrderID"))         '超值礼包
			
			
		OrderDetailStr=OrderDetailStr & "	   <tr class='tdbg' height='30' > "   
		OrderDetailStr=OrderDetailStr & "	    <td colspan='6' align='right'><b>合计：</b></td> "   
		OrderDetailStr=OrderDetailStr & "		<td align='right'><b>" & formatnumber(totalprice,2) & "</b></td>    "
		OrderDetailStr=OrderDetailStr & "		<td colspan='3'> </td>  "
		OrderDetailStr=OrderDetailStr & "	  </tr>    "
		OrderDetailStr=OrderDetailStr & "	  <tr class='tdbg'>"
       OrderDetailStr=OrderDetailStr & "         <td colspan='4'>付款方式折扣率：" & rs("Discount_Payment") & "%&nbsp;&nbsp;" 
	   If RS("Weight")>0 Then
	   OrderDetailStr=OrderDetailStr & "重量：" & rs("weight") & " KG"
	   End If
	   OrderDetailStr=OrderDetailStr & "&nbsp;&nbsp;运费：" & rs("Charge_Deliver")&" 元&nbsp;&nbsp;&nbsp;&nbsp;税率：" & KS.Setting(65) &"%&nbsp;&nbsp;&nbsp;&nbsp;价格含税："
				IF KS.Setting(64)=1 Then 
				   OrderDetailStr=OrderDetailStr & "是"
				  Else
				   OrderDetailStr=OrderDetailStr & "不含税"
				  End If
				  Dim TaxMoney
				  Dim TaxRate:TaxRate=KS.Setting(65)
				 If KS.Setting(64)=1 Or rs("NeedInvoice")=0 Then TaxMoney=1 Else TaxMoney=1+TaxRate/100

				OrderDetailStr=OrderDetailStr & "<br>订单金额：(" & rs("MoneyGoods") & "×" & rs("Discount_Payment") & "%＋"&rs("Charge_Deliver") & ")×"
				if KS.Setting(64)=1 Or rs("NeedInvoice")=0 Then OrderDetailStr=OrderDetailStr & "100%" Else OrderDetailStr=OrderDetailStr & "(1＋" & TaxRate & "%)" 
				OrderDetailStr=OrderDetailStr & "＝" & formatnumber(rs("NoUseCouponMoney"),2) & "元  </td>"
    OrderDetailStr=OrderDetailStr & "<td  colspan='3' align=right><b>订单金额：</b> ￥" & formatnumber(rs("NoUseCouponMoney"),2) & " 元<br>"
	If KS.ChkClng(RS("CouponUserID"))<>0 And RS("UseCouponMoney")>0 Then
	OrderDetailStr=OrderDetailStr & "<b>使用优惠券：</b> <font color=#ff6600>￥" & formatnumber(RS("UseCouponMoney"),2,-1) & " 元</font><br>"
	End If
	OrderDetailStr=OrderDetailStr & "<b>应付金额：</b> ￥" & formatnumber(rs("MoneyTotal"),2) & "  元</td>"
    OrderDetailStr=OrderDetailStr & "<td colspan='3' align='left'><b>已付款：</b>￥<font color=red>" & formatnumber(rs("MoneyReceipt"),2) & "</font></b>"
	If RS("MoneyReceipt")<RS("MoneyTotal") Then
	OrderDetailStr=OrderDetailStr & "<br><B>尚欠款：￥<font color=blue>" & formatnumber(RS("MoneyTotal")-RS("MoneyReceipt"),2) &"</B>"
	End If
	OrderDetailStr=OrderDetailStr & "</td></tr></table></td>  "
	OrderDetailStr=OrderDetailStr & "</tr>"  
	OrderDetailStr=OrderDetailStr & "     <tr><td><br><b>注：</b>“<font color='blue'>原价</font>”指商品的原始零售价，“<font color='green'>实价</font>”指系统自动计算出来的商品最终价格，“<font color='red'>指定价</font>”指管理员根据不同会员组手动指定的最终价格。商品的最终销售价格以“指定价”为准。“订单金额”指系统自动算出来的价格，本订单的最终价格以“<font color=#ff6600>应付金额</font>”为准。<br>"
	OrderDetailStr=OrderDetailStr & "	 </td>"
	OrderDetailStr=OrderDetailStr & "	 </tr>"

	If not conn.execute("select top 1 * from ks_orderitem where orderid='" & RS("OrderID") &"' and islimitbuy<>0").eof Then
	OrderDetailStr=OrderDetailStr & "     <tr><td><div style='margin:10px;color:red;padding:10px;margin-top:3px;border:1px solid #f9c943;background:#FFFFF6'>温馨提示:本订单是限时/限量抢购订单,限制下单后" & LimitBuyPayTime & "小时之内必须付款,即在[" & DateAdd("h",LimitBuyPayTime,RS("InputTime")) & "]之前用户没有付款,本订单自动作废。</div>"
	OrderDetailStr=OrderDetailStr & "	 </td>"
	OrderDetailStr=OrderDetailStr & "	 </tr>"
	End If
	If RS("DeliverStatus")=1 Then
	 Dim RSD,DeliverStr
	 Set RSD=Conn.Execute("Select Top 1 * From KS_LogDeliver Where DeliverType=1 And OrderID='" & RS("OrderID") & "'")
	 If Not RSD.Eof Then
	  DeliverStr="快递公司:" & RSD("ExpressCompany") & " 物流单号:" & RSD("ExpressNumber") & " 发货日期:" & RSD("DeliverDate") & " 发货经手人:" & RSD("HandlerName")
	 End If
	 RSD.Close : Set RSD=Nothing
	OrderDetailStr=OrderDetailStr & "     <tr><td><div style='margin:10px;color:blue;padding:10px;margin-top:3px;border:1px solid #f9c943;background:#FFFFF6'>温馨提示:本订单已发货。" & DeliverStr & "</div>"
	OrderDetailStr=OrderDetailStr & "	 </td>"
	OrderDetailStr=OrderDetailStr & "	 </tr>"
	End If
	
	
	OrderDetailStr=OrderDetailStr & "	</table>"
 End Function

'==================================三期开发代码=========================================
'取得捆绑销售商品
Function GetBundleSalePro(ByRef TotalPrice,ProID,OrderID)
  Dim Str,RS,XML,Node
  Set RS=Server.CreateObject("adodb.recordset")
  RS.Open "Select I.Title,I.Unit,O.* From KS_OrderItem O inner join KS_Product I On O.ProID=I.ID Where O.SaleType=6 and BundleSaleProID=" & ProID & " and OrderID='" & OrderID & "' order by O.id",conn,1,1
  If Not RS.Eof Then
    Set XML=KS.RsToXml(rs,"row","")
  End If
  RS.Close:Set RS=Nothing
  If IsObject(XML) Then
	     str=str & "<tr height=""25"" align=""left""><td colspan=9 style=""color:green"">&nbsp;&nbsp;选购捆绑促销:</td></tr>"
       For Each Node In Xml.DocumentElement.SelectNodes("row")
         str=str & "<tr>"
		 str=str &" <td style='color:#999999'>&nbsp;" & Node.SelectSingleNode("@title").text &"</td>"
		 str=str &" <td align='center'>" & Node.SelectSingleNode("@unit").text &"</td>"
		 str=str &" <td align='center'>" & Node.SelectSingleNode("@amount").text &"</td>"
		 str=str &" <td align='center'>" & formatnumber(Node.SelectSingleNode("@price_original").text,2,-1) &"</td>"
		 str=str &" <td align='center'>" & formatnumber(Node.SelectSingleNode("@realprice").text,2,-1) &"</td>"
		 str=str &" <td align='center'>" & formatnumber(Node.SelectSingleNode("@realprice").text,2,-1) &"</td>"
		 str=str &" <td align='right'>" & formatnumber(Node.SelectSingleNode("@realprice").text*Node.SelectSingleNode("@amount").text,2,-1) &"</td>"
		 str=str &" <td align='center'>" & Node.SelectSingleNode("@serviceterm").text &"</td>"
		 str=str &" <td align='center'>" & Node.SelectSingleNode("@remark").text &"</td>"
		 str=str & "</tr>"
		 TotalPrice=TotalPrice +round(Node.SelectSingleNode("@realprice").text*Node.SelectSingleNode("@amount").text,2) 
       Next
  End If
  GetBundleSalePro=str
End Function
'============================================================================================

 '得到超值礼包
 Function GetPackage(ByRef TotalPrice,OrderID)
	    If KS.IsNul(OrderID) Then Exit Function
		Dim RS,RSB,GXML,GNode,str,n,Price
		Set RS=Conn.Execute("select packid,OrderID from KS_OrderItem Where SaleType=5 and OrderID='" & OrderID & "' group by packid,OrderID")
		If Not RS.Eof Then
		 Set GXML=KS.RsToXml(Rs,"row","")
		End If
		RS.Close : Set RS=Nothing
		If IsOBJECT(GXml) Then
		   FOR 	Each GNode In GXML.DocumentElement.SelectNodes("row")
		     Set RSB=Conn.Execute("Select top 1 * From KS_ShopPackAge Where ID=" & GNode.SelectSingleNode("@packid").text)
			 If Not RSB.Eof Then
					  
						Dim RSS:Set RSS=Server.CreateObject("adodb.recordset")
						RSS.Open "Select a.title,a.GroupPrice,a.Price_Member,a.Price,b.* From KS_Product A inner join KS_OrderItem b on a.id=b.proid Where b.SaleType=5 and b.packid=" & GNode.SelectSingleNode("@packid").text & " and  b.orderid='" & OrderID & "'",Conn,1,1
						  str=str & "<tr class='tdbg' height=""25"" align=""center""><td colspan=2><strong><a href='" & DomainStr & "shop/pack.asp?id=" & RSB("ID") & "' target='_blank'>" & RSB("PackName") & "</a></strong></td>"
						  n=1
						  Dim TotalPackPrice,tempstr,i
						  TotalPackPrice=0 : tempstr=""
						Do While Not RSS.Eof
						 
						  For I=1 To RSS("Amount") 
							  '得到单件品价格 
							  IF KS.C("UserName")<>"" Then
								  If RSS("GroupPrice")=0 Then
								   Price=RSS("Price_Member")
								  Else
								   Dim RSP:Set RSP=Conn.Execute("Select Price From KS_ProPrice Where GroupID=(select groupid from ks_user where username='" & KS.C("UserName") & "') And ProID=" & RSS("ID"))
								   If RSP.Eof Then
									 Price=RSS("Price_Member")
								   Else
									 Price=RSP(0)
								   End If
								   RSP.Close:Set RSP=Nothing
								  End If
							  Else
								  Price=RSS("Price")
							  End If
							
							   TotalPackPrice=TotalPackPrice+Price
							  tempstr=tempstr & n & "." & rss("title") & " " & rss("AttributeCart") & "<br/>"
							  n=n+1
						  Next
						  RSS.MoveNext
						Loop
						
						str=str &"<td>1</td><td>￥" & TotalPackPrice & "</td><td>" & rsb("discount") & "折</td><td>￥" & formatnumber((TotalPackPrice*rsb("discount")/10),2,-1) & "</td><td>￥" & formatnumber((TotalPackPrice*rsb("discount")/10),2,-1) & "</td><td>---</td><td>---</td>"
					   
						str=str & "</tr><tr><td align='left' colspan=9>选择的套装详细如下:<br/>" & tempstr & "</td></tr>" 
						
						TotalPrice=TotalPrice+round(formatnumber((TotalPackPrice*rsb("discount")/10),2,-1))   '将礼包金额加入总价
						
						RSS.Close
						Set RSS=Nothing
					
			End If
			RSB.Close
		   Next
			
	    End If
		GetPackage=str
		
End Function


	
 '删除订单
 Sub DelOrder()
		 Dim ID:ID=KS.G("ID")
		 If ID="" Then KS.echo "<script>history.back();</script>" : Exit Sub
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select OrderID,CouponUserID From KS_Order Where Status<>2 And MoneyReceipt=0 And ID In(" & ID  &")",Conn,1,1
		 If Not RS.Eof Then
		  Do While Not RS.Eof
		   Conn.execute("Update KS_ShopCouponUser Set UseFlag=0,OrderID='' Where ID=" & rs(1))
		   Conn.Execute("Delete From KS_OrderItem Where OrderID='" & RS(0) & "'")
		   RS.MoveNext
		  Loop
		 End If
		 RS.Close:Set RS=Nothing
		 Conn.Execute("Delete From KS_Order Where Status<>2 And MoneyReceipt=0 And ID In(" & ID  &")")
		  KS.AlertHintScript "恭喜,订单删除成功!"
End Sub
		
		'确认订单
		Sub  OrderConfirm()
		  Dim MailContent:MailContent=KS.Setting(73)
		  Dim ID:ID=KS.G("ID")
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select * From KS_Order Where ID=" & ID ,Conn,1,3
		  If Not RS.Eof Then
		    RS("Status")=1
			RS.Update
			Dim RSA:Set RSA=Server.CreateObject("ADODB.RECORDSET")
			RSA.Open "Select ProID,Amount From KS_OrderItem Where OrderID='" & RS("OrderID") & "'",conn,1,1
			do while not rsa.eof
			 Conn.Execute("update ks_product set TotalNum=TotalNum-" & RSA(1) & " Where ID=" & RSA(0))
			 RSA.MoveNext
			loop
			rsa.close:set rsa=nothing
		    If Trim(RS("UserName"))<>"游客" Then   '游客下的订单不允许发送站内信件
				'参数Incept--接收者,Sender-发送者,title--主题,Content--信件内容
				Call KS.SendInfo(rs("username"),KS.C("AdminName"),"订单确认通知",ReplaceOrderLabel(MailContent,RS))
			End If
			If RS("Email")<>"" Then
				Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "收到汇款通知", RS("Email"),RS("ContactMan"), ReplaceOrderLabel(MailContent,rs),KS.Setting(11))
			 End If
		 %> <br><br>
		     <br><table align=center width='50%' border='0' cellpadding='2' cellspacing='1' class='border'>    
		       <tr align='center' class='title'>     
			   <td height='22'><b>恭喜你！ </b></td>
			   </tr>
			  <tr class='tdbg'><td><br>订单确认成功！
			  <%If Trim(RS("UserName"))<>"游客" Then%>
			  <br><br>已经向<%=rs("username")%>会员发送了一条站内短信，通知他订单已经确认！
			  <%end if%><br><br>
			   <%IF ReturnInfo="OK" Then%>
			  <br><br>已经向<%=rs("Email")%>发送了一封邮件通知，通知他订单已确认！
			  <%end if%>
			  
			  </td></tr>
			<tr class='tdbg'><td height=25 align=center><a href='KS.ShopOrder.asp?Action=ShowOrder&ID=<%=ID%>'><<点此返回</a></td></tr>
			</table>
		 <%
		  Else
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		  End If
		  RS.Close:Set RS=Nothing
		End Sub
		
		'银行付款
		Sub BankPay()
		 Dim ID:ID=KS.G("ID")
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select * From KS_Order Where ID=" & ID ,Conn,1,1
		 If RS.Eof Then
		   Response.Write "<script>alert('参数错误！');history.back();</script>"
		 End IF
		  %>
		<form name='form4' method='post' action='KS.ShopOrder.asp' onSubmit="return confirm('确定所输入的信息都完全正确吗？一旦输入就不可更改哦！')">  
		<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>    <tr align='center' class='title'>      <td height='25' colspan='2'><b>添 加 订 单 汇 款 信 息</b></td>    </tr>    <tr class='tdbg'>      <td width='15%' align='right'>客户姓名：</td>      <td><%=rs("contactman")%></td>    </tr>    <tr class='tdbg'>      <td width='15%' align='right'>用户名：</td>      <td><%=rs("username")%></td>    </tr>    <tr class='tdbg'>      <td width='15%' align='right'>支付内容：</td>      <td><table  border='0' cellspacing='2' cellpadding='0'>        <tr class='tdbg'>          <td width='15%' align='right'>订单编号：</td>          <td><%=rs("orderid")%></td>          <td>&nbsp;</td>        </tr>        <tr class='tdbg'>          <td width='15%' align='right'>订单金额：</td>          <td><%=rs("MoneyTotal")%>元</td>          <td></td>        </tr>        <tr class='tdbg'>          <td width='15%' align='right'>已 付 款：</td>          <td><%=rs("MoneyReceipt")%>元</td>          <td>&nbsp;</td>        </tr>      </table>      </td>    </tr>    <tr class='tdbg'>      <td width='15%' align='right'>汇款日期：</td>      <td><input name='PayDate' type='text' id='PayDate' value='<%=formatdatetime(now,2)%>' size='15' maxlength='30'></td>    </tr>    <tr class='tdbg'>      <td width='15%' align='right'>汇款金额：</td>      <td><input name='Money' type='text' id='Money' value='<%=rs("MoneyTotal")-rs("MoneyReceipt")%>' size='10' maxlength='10'> 元</td>    </tr>       <tr class='tdbg'>      <td width='15%' align='right'>备注：</td>      <td><input name='Remark' type='text' id='Remark' size='50' maxlength='200' value="支付订单费用，订单号：<%=rs("orderid")%>"></td>    </tr>    <tr class='tdbg'>      <td width='15%' align='right'>通知会员：</td>      <td><input type='checkbox' name='SendMessageToUser' value='1' checked>同时使用站内短信通知会员已经收到汇款<br><input type='checkbox' name='SendMailToUser' value='1' checked>同时发送邮件通知会员已经收到汇款</td>    </tr>    <tr class='tdbg'>      <td height='30' colspan='2'><b><font color='#FF0000'>注意：汇款信息一旦录入，就不能再修改或删除！所以在保存之前确认输入无误！</font></b></td>    </tr>    <tr align='center' class='tdbg'>      <td height='30' colspan='2'><input name='Action' type='hidden' id='Action' value='DoBankPay'>      <input name='ID' type='hidden' id='ID' value='<%=rs("id")%>'>      <input  class='button' type='submit' name='Submit' value='保存汇款信息'>&nbsp;<input type='button' class='button' onclick='javascript:history.back();' name='Submit' value='取消返回'></td>    </tr>  </table></form>
		<%
		RS.Close:Set RS=Nothing
		End Sub
		
		'开始银行支付操作
		Sub DoBankPay()
		 Dim ID:ID=KS.G("ID")
		 Dim PayDate:PayDate=KS.G("PayDate")
		 Dim Money:Money=KS.G("Money")
		 Dim Remark:Remark=KS.G("Remark")
		 Dim SendMessageToUser:SendMessageToUser=KS.ChkClng(KS.G("SendMessageToUser"))
		 Dim SendMailToUser:SendMailToUser=KS.ChkClng(KS.G("SendMailToUser"))
		 If Not IsDate(PayDate) Then Response.Write "<script>alert('付款日期格式有误');history.back();</script>":response.end
		 If Not IsNumeric(Money) Then 
		  Response.Write "<script>alert('输入的汇款金额不合法!');history.back();</script>":response.end
		 else
		  If Money<=0 Then
		  Response.Write "<script>alert('汇款金额必须大于0!');history.back();</script>":response.end
		  End If
		 End If
		 
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID,Conn,1,3
		 If RS.Eof Then
		    rs.close:set rs=nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   Exit Sub :Response.end
		 end if
		  If Remark="" Then Remark="支付订单费用，订单号：" & rs("orderid")
          RS("MoneyReceipt")=RS("MoneyReceipt")+Money
		  Dim OrderStatus:OrderStatus=RS("Status")
		  RS("Status")=1
				'==================================三期开发代码=========================================
				RS("PayTime")=now   '记录付款时间
				'=========================================================================================
		  RS.Update
		  Dim Email:Email=RS("Email")
		  Dim ContactMan:ContactMan=RS("ContactMan")
		  Call KS.MoneyInOrOut(rs("UserName"),ContactMan,Money,2,1,now,rs("orderid"),KS.C("AdminName"),"银行汇款",0,0,0)
		  Call KS.MoneyInOrOut(rs("UserName"),ContactMan,Money,4,2,now,rs("orderid"),KS.C("AdminName"),Remark,0,0,0)
		 If SendMessageToUser=1 and Trim(RS("UserName"))<>"游客" Then
				'参数Incept--接收者,Sender-发送者,title--主题,Content--信件内容
				Call KS.SendInfo(rs("username"),KS.C("AdminName"),"收到汇款通知",ReplaceOrderLabel(KS.Setting(74),rs))
		 End If
		 If SendMailToUser=1 and Email<>"" Then
		    Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "收到汇款通知", Email,ContactMan, ReplaceOrderLabel(KS.Setting(74),rs),KS.Setting(11))
		 End If
		 %>
		 <br><br>
		     <br><table align=center width='50%' border='0' cellpadding='2' cellspacing='1' class='border'>    
		       <tr align='center' class='title'>     
			   <td height='22'><b>恭喜你！ </b></td>
			   </tr>
			  <tr class='tdbg'><td><br>保存汇款信息成功！
			  <%If Trim(RS("UserName"))<>"游客" Then%>
			  <br><br>已经向<%=rs("username")%>会员发送了一条站内短信通知，通知他已经收到汇款！
			  <%end if%>
			  <%IF ReturnInfo="OK" Then%>
			  <br><br>已经向<%=Email%>发送了一封邮件通知，通知他已经收到汇款！
			  <%end if%>
			  </td></tr>
			<tr class='tdbg'><td height=25 align=center><a href='KS.ShopOrder.asp?Action=ShowOrder&ID=<%=ID%>'><<点此返回</a></td></tr>
			</table>
		 <%
		 
					'====================为用户增加购物应得积分========================
					Dim rsp:set rsp=conn.execute("select point,id,title from ks_product where id in(select proid from KS_OrderItem where orderid='" & rs("orderid") & "')")
					do while not rsp.eof
					  dim amount:amount=conn.execute("select amount from ks_orderitem where orderid='" & rs("orderid") & "' and proid=" & rsp(1))(0)
					  if OrderStatus<>1 Then
					  conn.execute("update ks_product set totalnum=totalnum-" & amount &" where totalnum>=" & amount &" and id=" & rsp(1))         '扣库存量
					  Call KS.ScoreInOrOut(rs("username"),1,KS.ChkClng(rsp(0))*amount,"系统","购买商品<font color=red>" & rsp("title") & "</font>赠送!",0,0)
					  End If
					rsp.movenext
					loop
					rsp.close
					set rsp=nothing
					'================================================================
		  RS.Close:Set RS=Nothing
		End Sub
		
		'退款
		Sub BankRefund()
		 Dim ID:ID=KS.G("ID")
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select * From KS_Order Where ID=" & ID ,Conn,1,1
		 If RS.Eof Then
		   Response.Write "<script>alert('参数错误！');history.back();</script>"
		 End IF
		  %>
<form name='form4' method='post' action='KS.ShopOrder.asp' onSubmit="return confirm('确定所输入的信息都完全正确吗？一旦输入就不可更改哦！')">  
<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>    <tr align='center' class='title'>      <td height='25' colspan='2'><b>添 加 订 单 退 款 信 息</b></td>    </tr>    <tr class='tdbg'>      <td width='15%' class='tdbg' align='right'>客户姓名：</td>      <td><%=rs("contactman")%></td>    </tr>    <tr class='tdbg'>      <td width='15%' class='tdbg' align='right'>用户名：</td>      <td><%=rs("username")%></td>    </tr>    <tr class='tdbg'>      <td width='15%' class='tdbg' align='right'>支付内容：</td>      <td><table  border='0' cellspacing='2' cellpadding='0'>        <tr class='tdbg'>          <td width='15%' class='tdbg' align='right'>订单编号：</td>          <td><%=rs("orderid")%></td>          <td>&nbsp;</td>        </tr>        <tr class='tdbg'>          <td width='15%' class='tdbg' align='right'>订单金额：</td>          <td><%=rs("moneytotal")%>元</td>        </tr>        <tr class='tdbg'>          <td width='15%' class='tdbg' align='right'>已 付 款：</td>          <td><%=rs("MoneyReceipt")%>元</td>          <td>&nbsp;</td>        </tr>      </table>      </td>    </tr>    <tr class='tdbg'>      <td width='15%' class='tdbg' align='right'>退款日期：</td>      <td><input name='PayDate' type='text' id='PayDate' value='<%=FormatDateTime(Now,2)%>' size='15' maxlength='30'></td>    </tr>    <tr class='tdbg'>      <td width='15%' class='tdbg' align='right'>退款金额：</td>      <td><input name='Money' type='text' id='Money'  size='10' value='<%=rs("MoneyReceipt")%>' maxlength='10'> 元&nbsp;&nbsp;<font color='#0000FF'>退款金额将从已付款中扣除。</font></td>    </tr>    <tr class='tdbg'>      <td width='15%' class='tdbg' align='right'>退款方式：</td>      <td><input type='radio' name='RefundType' value='1' onClick="Remark.value='订单退款金额，订单号：<%=RS("orderid")%>'" <%if rs("username")<>"游客" then Response.Write " checked"%>>扣除的金额添加到会员资金余额中<br><input type='radio' name='RefundType' value='2' onClick="Remark.value='订单退款金额，退款方式采用其它方式，订单号：<%=rs("orderid")%>'"<%if rs("username")="游客" then Response.Write " checked"%>>采用其它方式：如银行转帐，现金交付等等</td>    </tr>    <tr class='tdbg'>      <td width='15%' class='tdbg' align='right'>备注：</td>      <td><input name='Remark' type='text' id='Remark' value=<%if rs("username")<>"游客" then Response.Write "'订单退款金额，订单号："&rs("orderid") &"'"  Else Response.Write "'订单退款金额，退款方式采用其它方式，订单号：" & rs("orderid") & "'"%> size='50' maxlength='200'></td>    </tr>    <tr class='tdbg'>      <td width='15%' class='tdbg' align='right'>通知会员：</td>      <td><input type='checkbox' name='SendMessageToUser' value='1' checked>同时使用站内短信通知会员已经退款<br><input type='checkbox' name='SendMailToUser' value='1' checked>同时发送Email通知会员已经退款</td>    </tr>    <tr class='tdbg'>      <td height='30' colspan='2'><b><font color='#FF0000'>注意：退款信息一旦录入，就不能再修改或删除！所以在保存之前确认输入无误！<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp退款方式采用其它方式时，为方便查找退款记录，在会员资金明细中也会有对应记录，但收入金额为0</b></td>    </tr>    <tr align='center' class='tdbg'>      <td height='30' colspan='2'><input name='Action' type='hidden' id='Action' value='DoRefundMoney'>      <input name='ID' type='hidden' id='ID' value='<%=rs("id")%>'>      <input class='button' type='submit' name='Submit' value=' 保存退款款信息 '></td>    </tr>  </table></form>
		<%
		RS.Close:Set RS=Nothing
		End Sub
		
		'开始退款相关操作
		Sub DoRefundMoney()
		 Dim ID:ID=KS.G("ID")
		 Dim PayDate:PayDate=KS.G("PayDate")
		 Dim Money:Money=KS.G("Money")
		 Dim Remark:Remark=KS.G("Remark")
		 Dim RefundType:RefundType=KS.G("RefundType")
		 Dim SendMessageToUser:SendMessageToUser=KS.ChkClng(KS.G("SendMessageToUser"))
		 Dim SendMailToUser:SendMailToUser=KS.ChkClng(KS.G("SendMailToUser"))
		 If Not IsDate(PayDate) Then Response.Write "<script>alert('退款日期格式有误');history.back();</script>":response.end
		 If KS.ChkClng(Money)=0 Then Response.Write "<script>alert('退款金额必须大于0!');history.back();</script>":response.end
		 
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID,Conn,1,3
		 If RS.Eof Then
		    rs.close:set rs=nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   Exit Sub :Response.end
		 end if
		   
		  If round(Money)>round(RS("MoneyReceipt")) Then Response.Write "<script>alert('退款金额必须小于已付款金额!');history.back();</script>":response.end
		  If Remark="" Then Remark="订单退款金额，订单号：" & rs("orderid")
          RS("MoneyReceipt")=RS("MoneyReceipt")-Money
		  RS.Update
		  Dim Email:Email=RS("Email")
		  Dim ContactMan:ContactMan=RS("ContactMan")
		  
		  Call KS.MoneyInOrOut(rs("UserName"),ContactMan,Money,4,1,now,rs("orderid"),KS.C("AdminName"),Remark,0,0,0)

		  
		 If SendMessageToUser=1 and Trim(RS("UserName"))<>"游客" Then
				'参数Incept--接收者,Sender-发送者,title--主题,Content--信件内容
				Call KS.SendInfo(rs("username"),KS.C("AdminName"),"退款通知",ReplaceOrderLabel(KS.Setting(75),rs))
		 End If
		 If SendMailToUser=1 and Email<>"" Then
		    Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "收到汇款通知", Email,ContactMan, ReplaceOrderLabel(KS.Setting(74),rs),KS.Setting(11))
		 End If
		 %>
		 <br><br>
		     <br><table align=center width='50%' border='0' cellpadding='2' cellspacing='1' class='border'>    
		       <tr align='center' class='title'>     
			   <td height='22'><b>恭喜你！ </b></td>
			   </tr>
			  <tr class='tdbg'><td><br>保存退款信息成功！
			  <%If Trim(RS("UserName"))<>"游客" Then%>
			  <br><br>已经向<%=rs("username")%>会员发送了一条站内短信，通知他已经退款！
			  <%end if%>
			  <%IF ReturnInfo="OK" Then%>
			  <br><br>已经向<%=Email%>发送了一封邮件通知，通知他已经退款！
			  <%end if%>
			  </td></tr>
			<tr class='tdbg'><td height=25 align=center><a href='KS.ShopOrder.asp?Action=ShowOrder&ID=<%=ID%>'><<点此返回</a></td></tr>
			</table>
		 <%
					'====================为用户增加购物应得积分========================
					Dim rsp:set rsp=conn.execute("select point,id from ks_product where id in(select proid from KS_OrderItem where orderid='" & rs("orderid") & "')")
					do while not rsp.eof
					  dim amount:amount=conn.execute("select amount from ks_orderitem where orderid='" & rs("orderid") & "' and proid=" & rsp(1))(0)
					  conn.execute("update ks_user set score=score-" & KS.ChkClng(rsp(0))*amount & " where username='" & rs("username") & "'")
					rsp.movenext
					loop
					rsp.close
					set rsp=nothing
					'================================================================
		  RS.Close:Set RS=Nothing
		End Sub
		
		'发货操作
		Sub DeliverGoods()
		 Dim ID:ID=KS.G("ID")
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select * From KS_Order Where ID=" & ID ,Conn,1,1
		 If RS.Eof Then
		   Response.Write "<script>alert('参数错误！');history.back();</script>"
		 End IF
        %><br>
<FORM name=form4 onSubmit="return confirm('确定录入的发货信息都正确无误了吗？');" action="KS.ShopOrder.asp" method=post>
  <table class=border cellSpacing=1 cellPadding=2 width="100%" border=0>
    <tr class=title align=middle>
      <td colSpan=2 height=25><B>录 入 发 货 信 息</B></td>
    </tr>
    <tr class=tdbg>
      <td  align=right width="15%">客户名称：</td>
      <td><%=rs("contactman")%></td>
    </tr>
    <tr class=tdbg>
      <td  align=right width="15%">用户名：</td>
      <td><%=rs("username")%></td>
    </tr>
    <tr class=tdbg>
      <td  align=right width="15%">收货人姓名：</td>
      <td><%=rs("contactman")%></td>
    </tr>
    <tr class=tdbg>
      <td  align=right width="15%">订单编号：</td>
      <td><%=rs("orderid")%></td>
    </tr>
    <tr class=tdbg>
      <td  align=right width="15%">订单金额：</td>
      <td><%=formatnumber(rs("MoneyTotal"),2,-1,-1)%>元</td>
    </tr>
    <tr class=tdbg>
      <td  align=right width="15%">已 付 款：</td>
      <td><%=formatnumber(rs("MoneyReceipt"),2,-1,-1)%>元</td>
    </tr>
    <tr class=tdbg>
      <td  align=right width="15%">客户指定：</td>
      <td>快递公司:<%
	  dim rst,foundexpress,companyname
	  Set RST=Server.CreateObject("ADODB.RECORDSET")
	 RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and a.tocity like '%"&rs("tocity")&"%'",conn,1,1
	 If RST.Eof Then
	    foundexpress=false
	 Else
	    foundexpress=true
		companyname=rst("typename")
	response.write "<span style='color:green'>" & companyname & "</span> 首重<span style='color:#ff6600'>"&rst("fweight")&"kg/"&rst("carriage")&"元</span>  续重<span style='color:#ff6600'>"&rst("W_fee")&"kg/"&rst("C_fee")&"元</span>"
	 End If
	 RST.Close
	 If foundexpress=false Then
	  If DataBaseType=1 Then
		  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and (convert(varchar(200),tocity)='' or a.tocity is null)",conn,1,1
	 Else
	  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and (a.tocity='' or a.tocity is null)",conn,1,1
	 End If
	  if rst.eof then
	    rst.close : set rst=nothing
	  else
	response.write "<span style='color:green'>" & rst("typename") & "</span> 首重<span style='color:#ff6600'>"&rst("fweight")&"kg/"&rst("carriage")&"元</span>  续重<span style='color:#ff6600'>"&rst("W_fee")&"kg/"&rst("C_fee")&"元</span>"
	   rst.close
	  end if
	 End If
	 set rst=nothing
	
	
	response.write " 发往<span style='color:red'>" & rs("tocity") & "</span>"
	  
	  %></td>
    </tr>
    <tr class=tdbg>
      <td  align=right width="15%">发货日期：</td>
      <td>
        <Input id="DeliverDate" maxLength=30 size=15 value="<%=formatdatetime(now,2)%>" name="DeliverDate"></td>
    </tr>
    <tr class=tdbg>
      <td  align=right width="15%">快递公司：</td>
      <td>
  <Input id="ExpressCompany" maxLength=30 size=15 name="ExpressCompany" value="<%=companyname%>"> <=
  <select id="Code" name="Code" onChange="document.getElementById('ExpressCompany').value=this.value"> 
          <option value=''>---快速选择快递公司---</option>
           <%
		    dim rss:set rss=conn.execute("select * from ks_deliverytype")
			do while not rss.eof
			  response.write "<option value='" & rss("typename") &"'>" & rss("typename") & "</option>"
			  rss.movenext
			loop
			rss.close
			set rss=nothing
		   %>

			 </select>
     </td>
    </tr>
    <tr class=tdbg>
      <td  align=right width="15%">快递单号：</td>
      <td>
        <Input id="ExpressNumber" maxLength=30 size=15 name="ExpressNumber"></td>
    </tr>
    <tr class=tdbg>
      <td  align=right width="15%">经 手 人：</td>
      <td>
        <Input id="HandlerName" maxLength=50 size=30 value="<%=KS.C("AdminName")%>" name="HandlerName"></td>
    </tr>
    <tr class=tdbg>
      <td  align=right width="15%">备&nbsp;&nbsp;&nbsp;&nbsp;注：</td>
      <td>
        <Input id=Remark maxLength=200 size=50 name="Remark" value="订单号：<%=rs("orderid")%>的货物已送出"></td>
    </tr>
    <tr class=tdbg>
      <td  align=right width="15%">通知会员：</td>
      <td>
  <Input type=checkbox CHECKED value="1" name="SendMessageToUser">同时使用站内短信通知会员已经发货<br>
  <input type="checkbox" checked value="1" name="SendMailToUser">同时发送Email通知会员已经发货</td>
    </tr>
    <tr class=tdbg align=middle>
      <td colSpan=2 height=30>
	  <Input id=Action type=hidden value="DoDeliverGoods" name="Action"> 
	  <Input id=OrderFormID type=hidden value="<%=rs("id")%>" name="ID"> 
      <Input class='button' type=submit value=" 保 存 发 货" name=Submit></td>
    </tr>
  </table>
</FORM>
		<% rs.close:set rs=nothing
		End Sub
		
		'发货操作
		Sub DoDeliverGoods()
		 Dim ID:ID=KS.G("ID")
		 Dim DeliverDate:DeliverDate=KS.G("DeliverDate")
		 Dim ExpressCompany:ExpressCompany=KS.G("ExpressCompany")
		 Dim ExpressNumber:ExpressNumber=KS.G("ExpressNumber")
		 Dim HandlerName:HandlerName=KS.G("HandlerName")
		 Dim Remark:Remark=KS.G("Remark")
		 Dim SendMessageToUser:SendMessageToUser=KS.ChkClng(KS.G("SendMessageToUser"))
		 Dim SendMailToUser:SendMailToUser=KS.ChkClng(KS.G("SendMailToUser"))
		 
		 If Not IsDate(DeliverDate) Then Response.Write "<script>alert('发货日期格式有误');history.back();</script>":response.end
		 If (HandlerName="") Then Response.Write "<script>alert('经手人必须填写');history.back();</script>":response.end
		 If (ExpressCompany="") Then Response.Write "<script>alert('快递公司必须填写');history.back();</script>":response.end
		 If (ExpressNumber="") Then Response.Write "<script>alert('快递单号必须填写');history.back();</script>":response.end
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID,Conn,1,3
		 If RS.Eof Then
		   rs.close:set rs=nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   Exit Sub :Response.end
		 end if
		 If rs("DeliverStatus")=1 Then  Response.Write "<script>alert('此订单已经发过货!');history.back();</script>":Response.end 
		  rs("DeliverStatus")=1
		  rs.update
		  Dim Email:Email=RS("Email")
		  Dim ContactMan:ContactMan=rs("Contactman")
		  Dim RSLog:Set RSLog=Server.CreateObject("ADODB.RECORDSET")
		  RSLog.Open "Select top 1 * From KS_LogDeliver",Conn,1,3
		   RSLog.AddNew
		    RSLog("OrderID")=RS("OrderID")
			RSLog("UserName")=RS("UserName")
			RSLog("ClientName")=RS("ContactMan")
			RSLog("Inputer")=KS.C("AdminName")
			RSLog("HandlerName")=HandlerName  
			RSLog("DeliverDate")=DeliverDate
			RSLog("DeliverType")=1  '发货
			RSLog("Remark")=Remark
			RSLog("ExpressCompany")=ExpressCompany
			RSLog("ExpressNumber")=ExpressNumber
			RSLog("Status")=0
		 RSLog.Update
		 RSLog.Close:Set RSLog=Nothing
		  If SendMessageToUser=1 and trim(rs("UserName"))<>"游客" Then
				'参数Incept--接收者,Sender-发送者,title--主题,Content--信件内容
				Call KS.SendInfo(rs("username"),KS.C("AdminName"),"发货通知",ReplaceOrderLabel(KS.Setting(77),rs))
		 End If
		 If SendMailToUser=1 and Email<>"" Then
		    Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "发货通知", Email,ContactMan, ReplaceOrderLabel(KS.Setting(77),rs),KS.Setting(11))
		 End If
%>
		 <br><br>
		     <br><table align=center width='50%' border='0' cellpadding='2' cellspacing='1' class='border'>    
		       <tr align='center' class='title'>     
			   <td height='22'><b>恭喜你！ </b></td>
			   </tr>
			  <tr class='tdbg'><td><br>保存发货信息成功！
			  <%If Trim(RS("UserName"))<>"游客" Then%>
			  <br><br>已经向<%=rs("username")%>会员发送了一条站内短信，通知他已经发货！
			  <%end if%>
			  <%IF ReturnInfo="OK" Then%>
			  <br><br>已经向<%=Email%>发送了一封邮件通知，通知他已经发货！
			  <%end if%>
			  </td></tr>
			<tr class='tdbg'><td height=25 align=center><a href='KS.ShopOrder.asp?Action=ShowOrder&ID=<%=ID%>'><<点此返回</a></td></tr>
			</table>
			<%
		 RS.Close:Set RS=Nothing
		End Sub
		
		'退货操作
		Sub BackGoods()
		 Dim ID:ID=KS.G("ID")
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID ,Conn,1,1
		 If RS.Eof Then
		   Response.Write "<script>alert('参数错误！');history.back();</script>"
		 End IF
		%>
				<br>
		<FORM name=form4 onSubmit="return confirm('确定录入的退货信息都正确无误了吗？');" action=KS.ShopOrder.asp method=post>
		<table class=border cellSpacing=1 cellPadding=2 width="100%" border=0>
		  <tr class=title align=middle>
			<td colSpan=2 height=22><B>录 入 退 货 信 息</B></td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 align=right width="15%">客户名称：</td>
			<td><%=rs("contactman")%></td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 align=right width="15%">用户名：</td>
			<td><%=rs("username")%></td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 align=right width="15%">收货人姓名：</td>
			<td><%=rs("contactman")%></td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 align=right width="15%">订单编号：</td>
			<td><%=rs("orderid")%></td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 align=right width="15%">订单金额：</td>
			<td><%=formatnumber(rs("MoneyTotal"),2,-1,-1)%>元</td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 align=right width="15%">已 付 款：</td>
			<td><%=formatnumber(rs("MoneyReceipt"),2,-1,-1)%>元</td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 align=right width="15%">送货方式：</td>
			<td>快递公司:<%
	  dim rst,foundexpress,companyname
	  Set RST=Server.CreateObject("ADODB.RECORDSET")
	 RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and a.tocity like '%"&rs("tocity")&"%'",conn,1,1
	 If RST.Eof Then
	    foundexpress=false
	 Else
	    foundexpress=true
		companyname=rst("typename")
	response.write "<span style='color:green'>" & companyname & "</span> 首重<span style='color:#ff6600'>"&rst("fweight")&"kg/"&rst("carriage")&"元</span>  续重<span style='color:#ff6600'>"&rst("W_fee")&"kg/"&rst("C_fee")&"元</span>"
	 End If
	 RST.Close
	 If foundexpress=false Then
		 If DataBaseType=1 Then
			  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and (convert(varchar(200),tocity)='' or a.tocity is null)",conn,1,1
		Else
		  RST.Open "select Top 1 a.fweight,carriage,c_fee,w_fee,b.typename from KS_Delivery a inner join KS_DeliveryType b on A.ExpressID=B.TypeID where a.ExpressID="& rs("delivertype") &" and (a.tocity='' or a.tocity is null)",conn,1,1
		 End If
	  if rst.eof then
	    rst.close : set rst=nothing
	  else
	response.write "<span style='color:green'>" & rst("typename") & "</span> 首重<span style='color:#ff6600'>"&rst("fweight")&"kg/"&rst("carriage")&"元</span>  续重<span style='color:#ff6600'>"&rst("W_fee")&"kg/"&rst("C_fee")&"元</span>"
	  rst.close
	  end if
	  
	 End If
	 set rst=nothing
	
	
	response.write " 发往<span style='color:red'>" & rs("tocity") & "</span>"
	  
	  %>&nbsp;&nbsp;&nbsp;&nbsp;<font color=blue>客户指定的送货方式</font></td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 align=right width="15%">退货日期：</td>
			<td>
			  <Input id=DeliverDate maxLength=30 size=15 value="<%=now%>" name=DeliverDate></td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 align=right width="15%">经 手 人：</td>
			<td>
			  <Input id=HandlerName maxLength=50 size=30 value="<%=KS.C("AdminName")%>" name=HandlerName></td>
		  </tr>
		  <tr class=tdbg>
			<td class=tdbg5 align=right width="15%">退货原因：</td>
			<td>
			  <Input id=Remark maxLength=200 size=50 name=Remark></td>
		  </tr>
		  <tr class=tdbg align=middle>
			<td colSpan=2 height=30>
		  <Input id=Action type=hidden value="SaveBack" name=Action> 
		  <Input id=ID type=hidden value=<%=rs("id")%> name=ID> 
			  <Input type=submit value=" 保 存 " class="button" name=Submit></td>
		  </tr>
		</table>
		</FORM>
		<%
		rs.close:set rs=nothing
		End Sub
		
		'退货操作
		Sub SaveBack()
		 Dim ID:ID=KS.G("ID")
		 Dim DeliverDate:DeliverDate=KS.G("DeliverDate")
		 Dim HandlerName:HandlerName=KS.G("HandlerName")
		 Dim Remark:Remark=KS.G("Remark")
		 
		 If Not IsDate(DeliverDate) Then Response.Write "<script>alert('退货日期格式有误');history.back();</script>":response.end
		 If (HandlerName="") Then Response.Write "<script>alert('经手人必须填写');history.back();</script>":response.end
		 If Remark="" Then Response.Write "<script>alert('请输入退货原因!');history.back();</script>":response.end
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID,Conn,1,3
		 If RS.Eof Then
		   rs.close:set rs=nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   Exit Sub :Response.end
		 end if
		  dim DeliverStatus:DeliverStatus=rs("DeliverStatus")
		  rs("DeliverStatus")=3
		  rs.update
		  
		  
		  if DeliverStatus<>3 then
		   '====================为用户减少购物应得积分========================
					Dim rsp:set rsp=conn.execute("select point,id,title from ks_product where id in(select proid from KS_OrderItem where orderid='" & ID & "')")
					do while not rsp.eof
					  dim amount:amount=conn.execute("select top 1 amount from ks_orderitem where orderid='" &ID & "' and proid=" & rsp(1))(0)
					  conn.execute("update ks_product set totalnum=totalnum+" & amount &" where id=" & rsp(1))         '扣库存量
					 ' response.write rs("orderid") & "=55<br>"
					 ' response.write amount & "<br>"
					 ' response.write username & "<br>"
					  
					  Call KS.ScoreInOrOut(UserName,2,KS.ChkClng(rsp(0))*amount,"系统","商品退货<font color=red>" & rsp("title") & "</font>扣除!",0,0)

					  
					rsp.movenext
					loop
					rsp.close
					set rsp=nothing
					'================================================================
		  end if
		  
		  
		  Dim RSLog:Set RSLog=Server.CreateObject("ADODB.RECORDSET")
		  RSLog.Open "Select top 1 * From KS_LogDeliver where DeliverType=2 and orderid='" & RS("OrderID") & "'",Conn,1,3
		  If RSLog.Eof Then
		   RSLog.AddNew
		  End If
		    RSLog("OrderID")=RS("OrderID")
			RSLog("UserName")=RS("UserName")
			RSLog("ClientName")=RS("ContactMan")
			RSLog("Inputer")=KS.C("AdminName")
			RSLog("HandlerName")=HandlerName  
			RSLog("DeliverDate")=DeliverDate
			RSLog("DeliverType")=2  '退货
			RSLog("Remark")=Remark
			RSLog("ExpressCompany")=""
			RSLog("ExpressNumber")=""
			RSLog("Status")=0
		 RSLog.Update
		 RSLog.Close:Set RSLog=Nothing
%>
		 <br><br>
		     <br><table align=center width='50%' border='0' cellpadding='2' cellspacing='1' class='border'>    
		       <tr align='center' class='title'>     
			   <td height='22'><b>恭喜你！ </b></td>
			   </tr>
			  <tr class='tdbg'><td><br>保存退货信息成功！
			 <br><br></td></tr>
			<tr class='tdbg'><td height=25 align=center><a href='KS.ShopOrder.asp?Action=ShowOrder&ID=<%=ID%>'><<点此返回</a></td></tr>
			</table>
			<%
		 RS.Close:Set RS=Nothing		
		 End Sub
		
		'开发票
		Sub Invoice()
		 Dim ID:ID=KS.G("ID")
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select * From KS_Order Where ID=" & ID ,Conn,1,1
		 If RS.Eof Then
		   Response.Write "<script>alert('参数错误！');history.back();</script>"
		 End IF
		%>
		<FORM name=form4 onSubmit="return confirm('确定录入的发票信息都正确无误了吗？');" action="KS.ShopOrder.asp" method=post>
  <table class=border cellSpacing=1 cellPadding=2 width="100%" border=0>
    <tr class=title align=middle>
      <td colSpan=2 height=22><B>录 入 开 发 票 信 息</B></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">客户名称：</td>
      <td><%=RS("ContactMan")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">用户名：</td>
      <td><%=RS("UserName")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">订单编号：</td>
      <td><%=RS("OrderID")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">订单金额：</td>
      <td><%=RS("MoneyTotal")%>元</td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">已 付 款：</td>
      <td><%=RS("MoneyReceipt")%>元</td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">发票信息：</td>
      <td><%=RS("InvoiceContent")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">开票日期：</td>
      <td>
        <Input id="InvoiceDate" maxLength=30 size=15 value="<%=FormatDateTime(Now,2)%>" name="InvoiceDate"></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">发票类型：</td>
      <td>
<Select name="InvoiceType">
  <Option value="地税普通发票" selected>地税普通发票</Option>
  <Option value="国税普通发票">国税普通发票</Option>
  <Option value="增值税发票">增值税发票</Option>
      </Select></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">发票号码：</td>
      <td>
        <Input id=InvoiceNum maxLength=30 size=15 name="InvoiceNum"></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">发票抬头：</td>
      <td>
        <Input id=InvoiceTitle maxLength=50 size=50 value="" name="InvoiceTitle"></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">发票内容：</td>
      <td><TEXTAREA name=InvoiceContent rows=4 cols=50><%=RS("InvoiceContent")%></TEXTAREA></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">发票金额：</td>
      <td>
        <Input id="MoneyTotal" maxLength=15 size=15 value="<%=RS("MoneyTotal")%>" name="MoneyTotal"></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">开 票 人：</td>
      <td>
        <Input id="HandlerName" maxLength=30 size=15 value="<%=KS.C("AdminName")%>" name="HandlerName"></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">通知会员：</td>
      <td>
  <Input type=checkbox CHECKED value="1" name="SendMessageToUser">同时使用站内短信通知会员已经开具发票<br>
  <Input type=checkbox CHECKED value="1" name="SendMailToUser">同时发送Email通知会员已经开具发票<br>
  </td>
    </tr>
    <tr class=tdbg align=middle>
      <td colSpan=2 height=30>
  <Input id=Action type=hidden value="DoSaveInvoice" name="Action"> 
  <Input id="ID" type=hidden value="<%=RS("ID")%>" name="ID"> 
        <Input type=submit class='button' value=" 保 存 " name=Submit></td>
    </tr>
  </table>
</FORM>
		<%
		RS.Close:Set RS=Nothing
		End Sub
		
		'保存发票
		Sub DoSaveInvoice()
		 Dim ID:ID=KS.G("ID")
		 Dim InvoiceDate:InvoiceDate=KS.G("InvoiceDate")
		 Dim InvoiceType:InvoiceType=KS.G("InvoiceType")
		 Dim InvoiceNum:InvoiceNum=KS.G("InvoiceNum")
		 Dim InvoiceTitle:InvoiceTitle=KS.G("InvoiceTitle")
		 Dim InvoiceContent:InvoiceContent=KS.G("InvoiceContent")
		 Dim MoneyTotal:MoneyTotal=KS.G("MoneyTotal")
		 Dim HandlerName:HandlerName=KS.G("HandlerName")
		 Dim SendMessageToUser:SendMessageToUser=KS.ChkClng(KS.G("SendMessageToUser"))
		 Dim SendMailToUser:SendMailToUser=KS.ChkClng(KS.G("SendMailToUser"))
		 If Not IsDate(InvoiceDate) Then Response.Write "<script>alert('开票日期格式有误');history.back();</script>":response.end
		 If (HandlerName="") Then Response.Write "<script>alert('开票人必须填写');history.back();</script>":response.end
		 If (InvoiceTitle="") Then Response.Write "<script>alert('发票抬头必须填写');history.back();</script>":response.end
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID,Conn,1,3
		 If RS.Eof Then
		   rs.close:set rs=nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   Exit Sub :Response.end
		 end if
           rs("Invoiced")=1
		  rs.update
		  Dim Email:Email=RS("Email")
		  Dim ContactMan:ContactMan=rs("ContactMan")
		  Dim RSLog:Set RSLog=Server.CreateObject("ADODB.RECORDSET")
		  RSLog.Open "Select top 1 * From KS_LogInvoice",Conn,1,3
		   RSLog.AddNew
			RSLog("UserName")=RS("UserName")
			RSLog("ClientName")=RS("ContactMan")		    
			RSLog("OrderID")=RS("OrderID")
            RSLog("InvoiceType")=InvoiceType
			RSLog("InvoiceNum")=InvoiceNum
			RSLog("InvoiceTitle")=InvoiceTitle
			RSLog("InvoiceContent")=InvoiceContent
			RSLog("InvoiceDate")=InvoiceDate
			RSLog("InputTime")=Now
			RSLog("MoneyTotal")=MoneyTotal
			RSLog("Inputer")=KS.C("AdminName")
			RSLog("HandlerName")=HandlerName  
		 RSLog.Update
		 RSLog.Close:Set RSLog=Nothing
		  If SendMessageToUser=1 and Trim(RS("UserName"))<>"游览" Then
				'参数Incept--接收者,Sender-发送者,title--主题,Content--信件内容
				Call KS.SendInfo(rs("username"),KS.C("AdminName"),"开发票通知",ReplaceOrderLabel(KS.Setting(76),rs))
		 End If
		 If SendMailToUser=1 and Email<>"" Then
		    Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "开发票通知", Email,ContactMan, ReplaceOrderLabel(KS.Setting(76),rs),KS.Setting(11))
		 End If
%>
		 <br><br>
		     <br><table align=center width='50%' border='0' cellpadding='2' cellspacing='1' class='border'>    
		       <tr align='center' class='title'>     
			   <td height='22'><b>恭喜你！ </b></td>
			   </tr>
			  <tr class='tdbg'><td><br>保存开发票信息成功！
			  <%If Trim(RS("UserName"))<>"游客" Then%>
			  <br><br>已经向<%=rs("username")%>会员发送了一条站内短信，通知他已经开发票！
			  <%end if%>
			  <%IF ReturnInfo="OK" Then%>
			  <br><br>已经向<%=Email%>发送了一封邮件通知，通知他已经开发票！
			  <%end if%></td></tr>
			<tr class='tdbg'><td height=25 align=center><a href='KS.ShopOrder.asp?Action=ShowOrder&ID=<%=ID%>'><<点此返回</a></td></tr>
			</table>
			<%
		 RS.Close:Set RS=Nothing
		End Sub
		
		'已签收商品
		Sub ClientSignUp()
		 Dim ID:ID=KS.G("ID")
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID,Conn,1,3
		 If RS.Eof Then
		   rs.close:set rs=nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   Exit Sub :Response.end
		 end if
		   rs("DeliverStatus")=2
		   rs.update
		   Conn.execute("update KS_LogDeliver Set Status=1 Where OrderID='" & RS("OrderID") & "'")
		 RS.Close:Set RS=Nothing
		 Response.Redirect "KS.ShopOrder.asp?Action=ShowOrder&ID=" & ID
		End Sub
		
		'打印清单
		Sub PrintOrder() 
		 Dim ID:ID=KS.G("ID")
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_Order Where ID=" & ID,Conn,1,1
		 If RS.Eof Then
		   rs.close:set rs=nothing
		   Response.Write "<script>alert('参数错误!');history.back();</script>"
		   Exit Sub :Response.end
		 end if
		 response.write "<br>" & OrderDetailStr(RS,0)
		 RS.Close:Set RS=Nothing
		 %>  <br>
		 <div id='Varea' align='center'>
		 	 <input type='button' class='button' name='Submit' value='开始打印' onClick="document.all.Varea.style.display='none';window.print();">&nbsp;<input type='button' class='button' name='Submit' value='取消打印' onClick="javascript:history.back();">
             </div>
		 <%
		End Sub
		
	 '支付货款给卖方
		Sub PayMoney()
		 Dim ID:ID=KS.G("ID")
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select * From KS_Order Where ID=" & ID ,Conn,1,1
		 If RS.Eof Then
		   Response.Write "<script>alert('参数错误！');history.back();</script>"
		 End IF
		  %>
		<form name='form4' method='post' action='KS.ShopOrder.asp' onSubmit="return confirm('确定所输入的信息都完全正确吗？一旦输入就不可更改哦！')">  
		<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>    <tr align='center' class='title'>      <td height='25' colspan='2'><b>支 付 货 款 给 卖 方</b></td>    </tr>    
		 <tr class='tdbg'>      <td width='15%' align='right'>订单信息：</td>      <td><table  border='0' cellspacing='2' cellpadding='0'>        <tr class='tdbg'>          <td width='15%' align='right'>订单编号：</td>          <td><%=rs("orderid")%></td>          <td>&nbsp;</td>        </tr>        <tr class='tdbg'>          <td width='15%' align='right'>订单金额：</td>          <td><%=rs("MoneyTotal")%>元</td>          <td></td>        </tr>        <tr class='tdbg'>          <td width='15%' align='right'>已 付 款：</td>          <td><%=rs("MoneyReceipt")%>元</td>          <td>&nbsp;</td>        </tr>      </table>      </td>    
		 </tr> 
		 <tr class='tdbg'>      
	  	  <td width='15%' align='right'>支付明细：</td>      
		  <td>
		  <%dim rso:set rso=server.createobject("adodb.recordset")
		  rso.open "select sum(a.TotalPrice),Inputer from ks_orderitem a inner join ks_product b on a.proid=b.id where a.orderid='" & rs("orderid") & "' group by inputer",conn,1,1
		  do while not rso.eof
		    response.write "卖方<font color=red>" & rso("inputer") & "</font>总价款:" & rso(0) & "元 本次应支付<font color=green>" & rso(0)-(rso(0) * ks.setting(79))/100 & "</font>元<br>"
		  rso.movenext
		  loop
		  rso.close
		  set rso=nothing
		  %>
		  </td>    
		 </tr>  
		
		  <tr class='tdbg'>      <td width='15%' align='right'>支付时间：</td>      <td><input name='PayDate' type='text' id='PayDate' value='<%=now%>' size='25' maxlength='30'></td>    </tr>  
		  
		   <tr class='tdbg'>      <td width='15%' align='right'>备注：</td>      <td><input name='Remark' type='text' id='Remark' size='50' maxlength='200' value="收到货款费用，订单号：<%=rs("orderid")%>"></td>    </tr>    <tr class='tdbg'>      <td width='15%' align='right'>通知会员：</td>      <td><input type='checkbox' name='SendMessageToUser' value='1' checked>同时使用站内短信通知卖主已经支付<br><input type='checkbox' name='SendMailToUser' value='1' checked>同时发送邮件通知卖主已经支付</td>    </tr>    <tr class='tdbg'>      <td height='30' colspan='2'><b><font color='#FF0000'>注意：一旦按确定支付，就不能再修改或删除！所以在保存之前确认输入无误！</font></b></td>    </tr>    <tr align='center' class='tdbg'>      <td height='30' colspan='2'><input name='Action' type='hidden' id='Action' value='DoPayMoney'>      <input name='OrderID' type='hidden' id='orderID' value='<%=rs("orderid")%>'>
		   <input name='ID' type='hidden' id='ID' value='<%=rs("id")%>'>
		   <input  class='button' type='submit' name='Submit' value='确定支付'>&nbsp;<input type='button' class='button' onclick='javascript:history.back();' name='Submit' value='取消返回'></td>    </tr>  </table></form>
		<%
		RS.Close:Set RS=Nothing
		End Sub
		
		'开始支付货款给卖家操作
		Sub DoPayMoney()
		 Dim OrderID:OrderID=KS.G("OrderID")
		 Dim PayDate:PayDate=KS.G("PayDate")
		 Dim Remark:Remark=KS.G("Remark")
		 If Remark="" Then Remark="收到货款费用，订单号：" & rs("orderid")

		 Dim SendMessageToUser:SendMessageToUser=KS.ChkClng(KS.G("SendMessageToUser"))
		 Dim SendMailToUser:SendMailToUser=KS.ChkClng(KS.G("SendMailToUser"))
		 If Not IsDate(PayDate) Then Response.Write "<script>alert('支付日期格式有误');history.back();</script>":response.end
		 If not Conn.Execute("Select PayToUser From ks_Order Where Paytouser=1 and OrderID='" & OrderID & "'").eof Then
		   response.write "<script>alert('对不起，该订单已支付过。不能重复支付!');history.back();</script>"
		   response.end
		 End If
		 
		 
		 dim rso,rsu
		 set rso=server.createobject("adodb.recordset")
		  rso.open "select sum(a.TotalPrice),Inputer from ks_orderitem a inner join ks_product b on a.proid=b.id where a.orderid='" & OrderID & "' group by inputer",conn,1,1
		  do while not rso.eof
		     set rsu=server.createobject("adodb.recordset")
			 rsu.open "select top 1 * from ks_user where username='" & rso(1) & "'",conn,1,1
			 if not rsu.eof then
			    Dim TotalMoney:TotalMoney=rso(0)
				Dim ServiceMoney:ServiceMoney=(TotalMoney * ks.setting(79))/100
				Dim MustPayMoney:MustPayMoney=(TotalMoney-ServiceMoney)
				
				Call KS.MoneyInOrOut(rsu("UserName"),rsu("RealName"),TotalMoney,4,1,PayDate,OrderID,KS.C("AdminName"),Remark,0,0,0)
				Call KS.MoneyInOrOut(rsu("UserName"),rsu("RealName"),ServiceMoney,4,2,PayDate,OrderID,KS.C("AdminName"),"支付订单:"& OrderID & "的服务费",0,0,0)

				 
				 Dim Email:Email=RSU("Email")
				 Dim ContactMan:ContactMan=RSU("RealName")
				 Dim SiteMessage,Mail,MailContent
				 If ContactMan="" or isnull(ContactMan) Then ContactMan=RSU("UserName")
				 
				 MailContent=KS.Setting(80)
				 MailContent=Replace(MailContent,"{$ContactMan}",ContactMan)
				 MailContent=Replace(MailContent,"{$OrderID}",orderid)
				 MailContent=Replace(MailContent,"{$TotalMoney}",TotalMoney)
				 MailContent=Replace(MailContent,"{$ServiceCharges}",ServiceMoney)
				 MailContent=Replace(MailContent,"{$RealMoney}",TotalMoney-ServiceMoney)
				 
		
				 If SendMessageToUser=1 Then
					'参数Incept--接收者,Sender-发送者,title--主题,Content--信件内容
					Call KS.SendInfo(rsu("username"),KS.C("AdminName"),"支付货款通知",MailContent)
					SiteMessage="已经向卖方" & rsu("username") & "发送了一条站内短信通知，通知他已经支付货款<br>"
				 End If
				 If SendMailToUser=1 and Email<>"" Then
					Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), KS.Setting(0)&"向您支付货款通知", Email,ContactMan, MailContent,KS.Setting(11))
					If ReturnInfo="OK" Then
					 Mail="已经向" & Email  &"发送了一封邮件通知，通知他已经支付货款！<br>"
					End If
				 End If
		 
			 end if
			 rsu.close
		  rso.movenext
		  loop
		  rso.close
		  set rso=nothing
		  set rsu=nothing

          '标志已支付
		  Conn.Execute("Update KS_Order Set PayToUser=1 where orderid='" & OrderID & "'")
		
		 %>
		 <br><br>
		     <br><table align=center width='50%' border='0' cellpadding='2' cellspacing='1' class='border'>    
		       <tr align='center' class='title'>     
			   <td height='22'><b>恭喜你！ </b></td>
			   </tr>
			  <tr class='tdbg'><td><br>已将货款支付给卖方！
			  <br><br>
			  <%=SiteMessage%>
			  <br>
			  <%=Mail%>
			  </td></tr>
			<tr class='tdbg'><td height=25 align=center><a href='KS.ShopOrder.asp?Action=ShowOrder&ID=<%=KS.G("ID")%>'><<点此返回</a></td></tr>
			</table>
		 <%
		End Sub		
		
		'替换可用订单标签
		Function ReplaceOrderLabel(MailContent,RS)
				 MailContent=Replace(MailContent,"{$ContactMan}",RS("ContactMan"))
				 MailContent=Replace(MailContent,"{$InputTime}",RS("InputTime"))
				 MailContent=Replace(MailContent,"{$OrderID}",RS("OrderID"))
				 MailContent=Replace(MailContent,"{$OrderInfo}",OrderDetailStr(RS,0))
				ReplaceOrderLabel=MailContent
		End Function
End Class
%> 
