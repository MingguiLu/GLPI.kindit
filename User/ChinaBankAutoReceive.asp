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
Dim PaymentPlat:PaymentPlat=7

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

 Call ChinaBank()
'�������߷���
Sub ChinaBank() 
 Dim v_oid,v_pmode,v_pstatus,v_pstring,v_string,v_amount,v_moneytype,remark2,v_md5str,text,md5text,zhuangtai
' ȡ�÷��ز���ֵ
	v_oid=request("v_oid")                               ' �̻����͵�v_oid�������
	v_pmode=request("v_pmode")                           ' ֧����ʽ���ַ����� 
	v_pstatus=request("v_pstatus")                       ' ֧��״̬ 20��֧���ɹ���;30��֧��ʧ�ܣ�
	v_pstring=request("v_pstring")                       ' ֧�������Ϣ ֧����ɣ���v_pstatus=20ʱ����ʧ��ԭ�򣨵�v_pstatus=30ʱ����
	v_amount=request("v_amount")                         ' ����ʵ��֧�����
	v_moneytype=request("v_moneytype")                   ' ����ʵ��֧������
	remark2=request("remark2")                           ' ��ע�ֶ�2
	v_md5str=request("v_md5str")                         ' ��������ƴ�յ�Md5У�鴮
	if request("v_md5str")="" then
		response.Write("v_md5str����ֵ")
		response.end
	end if
	text = v_oid&v_pstatus&v_amount&v_moneytype&MD5Key 'md5У��
	md5text = Ucase(trim(md5(text,32)))    '�̻�ƴ�յ�Md5У�鴮
	if md5text<>v_md5str then		' ��������ƴ�յ�Md5У�鴮 �� �̻�ƴ�յ�Md5У�鴮 ���жԱ�
	  	response.write("error") '���߷�������֤ʧ�ܣ�Ҫ���ط�
	    response.end '�жϳ���
	else
	  response.write("ok")
	  if v_pstatus=20 then '֧���ɹ�
		Call UpdateOrder(v_amount,remark2,v_oid,v_pmode)
	  else
	   	response.write("error") '���߷�������֤ʧ�ܣ�Ҫ���ط�
	    response.end '�жϳ���
	  end if
	end if
end Sub
%>