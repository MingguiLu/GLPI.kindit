<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="payfunction.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Response.Buffer = true 
Response.Expires = 1 
Response.CacheControl = "no-cache"

Dim KSUser:Set KSUser=New UserCls
Dim KS:Set KS=New PublicCls
Dim PaymentPlat:PaymentPlat=KS.ChkClng(KS.S("PaymentPlat"))
If PaymentPlat=0 Then PaymentPlat=12  'paypal���ʰ�

Dim RSP:Set RSP=Server.CreateObject("ADODB.RECORDSET")
RSP.Open "Select top 1 * From KS_PaymentPlat where id=" & PaymentPlat,conn,1,1
If RSP.Eof Then
		 RSP.Close:Set RSP=Nothing
		 Response.Write "Error!"
		 Response.End()
End If
Dim AccountID:AccountID=RSP("AccountID")
Dim MD5Key:MD5Key=RSP("MD5Key")
Dim PayOnlineRate:PayOnlineRate=KS.ChkClng(RSP("Rate")) 
Dim RateByUser:RateByUser=KS.ChkClng(RSP("RateByUser")) 
RSP.Close:Set RSP=Nothing
Call PayPal()
'paypal ���ʰ�
Sub PayPal()

		Dim Item_name, Item_number, Payment_status, Payment_amount,Payment_currency
		Dim Txn_id, Receiver_email, Payer_email
		Dim objHttp, str,paypalurl,msg
		
		' read post from PayPal system and add 'cmd'
		str = Request.Form & "&cmd=_notify-validate"
		
	     'paypalurl="https://www.sandbox.paypal.com/cgi-bin/webscr"   '���Խӿ�ר�ã���ʽʹ��Ҫʹ�����½ӿ�
	     paypalurl="https://www.paypal.com/cgi-bin/webscr"            '��ʽ����ʹ�ô˽ӿ�
		 
		 
		set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
		' set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
		' set objHttp = Server.CreateObject("Microsoft.XMLHTTP")
		objHttp.open "POST", paypalurl, false
		objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
		objHttp.Send str
		
		' assign posted variables to local variables
		Item_name = Request.Form("item_name")
		Item_number = Request.Form("item_number")
		Payment_status = Request.Form("payment_status")
		Payment_amount = Request.Form("mc_gross")
		Payment_currency = Request.Form("mc_currency")
		Txn_id = Request.Form("txn_id")
		Receiver_email = Request.Form("receiver_email")
		Payer_email = Request.Form("payer_email")
		
		 dim v_oid,usdmoney,v_amount,remark2,v_pmode,i,aParts
         Dim CusTom:CusTom=Request.Form("CusTom")
		 If Instr(CusTom,"|")<>0 Then
		 SUserName=Split(Custom,"|")(0)
		 sPayFrom=split(Custom,"|")(1)
		 SUserCardID=split(Custom,"|")(2)
		 End if		
		 v_amount=round(Payment_amount*KS.Setting(81),2)
         v_oid=item_number

		' Check notification validation
		if (objHttp.status <> 200 ) then
		' HTTP error handling
		elseif (objHttp.responseText = "VERIFIED") then
		  Msg="��ϲ��֧���ɹ���"
		  Call UpdateOrder(v_amount,"֧��������" & v_oid &"����!",v_oid,"PayPal")
		' check that Payment_status=Completed
		' check that Txn_id has not been previously processed
		' check that Receiver_email is your Primary PayPal email
		' check that Payment_amount/Payment_currency are correct
		' process payment
		elseif (objHttp.responseText = "INVALID") then
		 Msg="�Բ���֧��ʧ�ܣ�����ϵ��վ����Ա��"
		' log for manual investigation
		else
		' error
		end if
		set objHttp = nothing

    ks.die msg 
	'Call ShowResult(Msg)

End Sub

function urldecodes(encodestr)   '��������Ƕ�paypal����ֵ��urldecode�����
	dim newstr:newstr="" 
	dim havechar:havechar=false 
	dim lastchar:lastchar="" 
	dim i,char_c,next_1_c,next_1_num
	for i=1 to len(encodestr) 
	char_c=mid(encodestr,i,1) 
	if char_c="+" then 
	newstr=newstr & " " 
	elseif char_c="%" then 
	next_1_c=mid(encodestr,i+1,2) 
	next_1_num=cint("&H" & next_1_c) 
	if havechar then 
	havechar=false 
	newstr=newstr & chr(cint("&H" & lastchar & next_1_c)) 
	else 
	if abs(next_1_num)<=127 then 
	newstr=newstr & chr(next_1_num) 
	else 
	havechar=true 
	lastchar=next_1_c 
	end if 
	end if 
	i=i+2 
	else 
	newstr=newstr & char_c 
	end if 
	next 
	urldecodes=newstr 
end Function

%>