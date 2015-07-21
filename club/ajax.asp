<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="../KS_Cls/ClubFunction.asp"-->
<!--#include file="../Plus/Session.asp"-->
<%
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="gb2312" 
Dim master,BSetting
Dim KS:Set KS=New PublicCls
Dim id:id=KS.ChkClng(KS.S("ID"))
Dim BoardID:BoardID=KS.ChkClng(KS.S("BoardID"))
Dim TopicID:TopicID=KS.ChkClng(KS.S("TopicID"))
Dim Action:Action=KS.S("Action")
select case Action
     case "fav" call fav
	 case "delusertopic" call delusertopic
	 case "settop" Call SetTOP
	 case "setbest" Call SetBest
	 case "canceltop" Call CancelTop
	 case "cancelbest" Call CancelBest
	 case "delsubject" Call delsubject
	 case "delreply" Call delreply
	 case "verify" Call verify
	 case "locked" Call Locked
	 case "unlocked" call unlocked
	 case "replylock" call replylock
	 case "replyunlock" call replyunlock
	 case "movetopic" call movetopic
	 case "support" support
	 case "opposition" opposition
	 case "lockuser" lockuser
	 case "unlockuser" unlockuser
	 case "verifictopic" verifictopic
	 case "dovote" dovote
	 case "checkcomments" checkcomments
	 case "comments" comments
	 case "getcommentpage" getcommentpage
	 case "delcomment" delcomment
	 case else ks.die "error!"
End select	
Set KS=Nothing
CloseConn

function getPostTable()
   dim rs :set rs=conn.execute("select top 1 PostTable From KS_GuestBook Where ID=" & TopicID)
   If RS.Eof Then
      RS.Close :Set RS=Nothing
	  KS.Die "error"
   End If
   getPostTable=RS(0)
   RS.Close : Set RS=Nothing
end function

'����Ƿ��������
sub checkcomments()
	Dim KSUser:Set KSUser=New UserCls
	Dim LoginTF:LoginTF=KSUser.UserLoginChecked()
	If Cbool(LoginTF)=False Then
	  KS.Die ("err|�Բ��𣬵�¼��ſ��Ե���!")
	End If
  Call doCheckComments(KSUser)
  If Conn.Execute("Select top 1 ID From KS_GuestComment Where pid=" & ID &" and UserName='" & KSUser.UserName & "' and Comment like '%��<i>%'").eof Then
  KS.Die "success|" & BSetting(49)
  Else
  KS.Die "success|"
  End If
  Set KSUser=Nothing		 
end sub
Sub doCheckComments(KSUser)
  if Boardid=0 then KS.Die "err|��������!"
  KS.LoadClubBoard()
  Dim Node:Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
  BSetting=Node.SelectSingleNode("@settings").text
  If KS.IsNul(BSetting) Then KS.Die "err|��������!"
  BSetting=Split(BSetting&"$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$","$")
  if KS.ChkClng(BSetting(44))=0 Then
    KS.Die "err|������û�п�����������!"
  End If

  If KSUser.GetUserInfo("userId")=KS.S("UserId") And KS.ChkClng(BSetting(48))=0 Then
    KS.Die "err|�����治������Լ������ӽ��е���!"
  End If
  If KS.ChkClng(KSUser.GetUserInfo("prestige"))<KS.ChkClng(BSetting(45)) Then
    KS.Die "err|�Բ�����������ֵ�������޷������ӽ��е���!"
  End If
  if KS.ChkClng(BSetting(46))=0 And KS.ChkClng(KS.S("N"))=1 Then
    KS.Die "err|�����治�����������е���!"
  End If
  if KS.ChkClng(BSetting(47))=0 And KS.ChkClng(KS.S("N"))>1 Then
    KS.Die "err|�����治����Իظ����е���!"
  End If
 End Sub

'�������
Sub comments()
	Dim KSUser:Set KSUser=New UserCls
	Dim RS,LoginTF:LoginTF=KSUser.UserLoginChecked()
	If Cbool(LoginTF)=False Then
	  KS.Die Escape("err|�Բ��𣬵�¼��ſ��Ե���!")
	End If
	Call doCheckComments(KSUser)
	If ID=0 Then KS.Die Escape("err|����������!")
	if IsDate(Request.Cookies("clubcmts")) then
      If DateDiff("s",Request.Cookies("clubcmts"),now)<15 Then
	     KS.Die Escape("err|���η�����ʱ�䲻������15�룬���Ժ򷢱�!")
	  End If
    end if
	If KS.ChkClng(KS.S("Prestige"))>2 Or KS.ChkClng(KS.S("Prestige"))<-2 Then
	     KS.Die Escape("err|���ύ�ˣ����Ϸ�������ֵ!")
	End If

    Dim Comment:Comment=UnEscape(Replace(KS.LoseHtml(KS.S("Comment")),"'",""))
	If KS.IsNul(Comment) Then KS.Die Escape("err|û�������������!")
	Conn.Execute("Insert Into KS_GuestComment(tid,pid,username,userface,userid,comment,adddate,Prestige,OrderID) values(" & topicid &"," & id & ",'" & KSUser.UserName & "','" & KSUser.GetUserInfo("userface") &"'," & KS.ChkClng(KSUser.GetUserInfo("UserID")) & ",'" & comment & "'," & SQLNowString& "," & KS.ChkClng(KS.S("Prestige"))&",1)")
	If Instr(comment,"��<i>")<>0 Then
	  Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open "Select top 1 * From KS_GuestComment Where Pid=" & id & " and username='' and userid=0",conn,1,3
	  If RS.Eof And RS.Bof Then
	    RS.AddNew
		 RS("TID")=topicid
		 RS("Pid")=id
		 RS("UserID")=0
		 RS("UserName")=""
		 RS("AddDate")=Now
		 RS("Comment")=split(comment,"<br/>")(0)
		 RS("Prestige")=0
		 RS("OrderID")=0
		RS.Update
	  Else
		 Call UpdateCommentStar(RS)
	  End If
	  RS.Close :Set RS=Nothing
	End If
	If KS.ChkClng(KS.S("Prestige"))<>0 Then
	    Set RS=Conn.Execute("Select top 1 UserID From " & getPostTable & " Where ID=" & ID)
		If Not RS.Eof Then
		Conn.Execute("Update KS_User Set Prestige=prestige+" & KS.ChkClng(KS.S("Prestige")) &" where userid=" & KS.ChkClng(RS(0)))
		End If
		RS.Close
    End If
    Response.Cookies("clubcmts")=Now
	
	'���½���������
	If Bsetting(50)="1" Then
	 dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		Doc.async = false
		Doc.setProperty "ServerHTTPRequest", true 
		Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
		Dim XMLDate:XMLDate=doc.documentElement.attributes.getNamedItem("date").text
					If DateDiff("d",xmldate,now)=0 Then
					  doc.documentElement.attributes.getNamedItem("todaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text+1
					  If KS.ChkClng(doc.documentElement.attributes.getNamedItem("todaynum").text)>KS.ChkClng(doc.documentElement.attributes.getNamedItem("maxdaynum").text) then
					   doc.documentElement.attributes.getNamedItem("maxdaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text
					  end if
					  Conn.Execute("Update KS_GuestBoard set postnum=postnum+1 where id=" & BoardID)
					  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text=KS.ChkClng(Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text)+1
					Else
					  doc.documentElement.attributes.getNamedItem("date").text=now
					  doc.documentElement.attributes.getNamedItem("yesterdaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text
					  doc.documentElement.attributes.getNamedItem("todaynum").text=0
					  Conn.Execute("Update KS_GuestBoard set postnum=1 where id=" & BoardID)
					  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text=1
					End If
					  doc.documentElement.attributes.getNamedItem("postnum").text=doc.documentElement.attributes.getNamedItem("postnum").text+1
	  doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
	  Conn.Execute("Update KS_GuestBook Set LastReplayTime=" & SQLNowString &",LastReplayUser='" & KSUser.UserName&"',LastReplayUserID=" & KS.ChkClng(KSUser.GetUserInfo("userid")) & " Where id=" & topicid) '����������󷢱�ʱ��
	End If
	Set KSUser=Nothing
	KS.Echo "success|"
	Set RS=Conn.Execute("Select * From KS_GuestComment Where Pid=" & ID & " Order By orderid,Id Desc")
	If Not RS.Eof Then
	   Dim Xml:Set XML=KS.RsToXml(RS,"row","")
	   KS.echo GetComments(XML,BoardID,id,KS.ChkClng(BSetting(44)),check)
	End If
	Set RS=Nothing
End Sub
'���µ�����������
Sub UpdateCommentStar(RST)
  Dim Comment:Comment=RST("Comment")
  Dim CommentArr:CommentArr=Split(Comment,"��")
  Dim K,GD,TempStr
  For K=0 To Ubound(CommentArr)
    If K=0 Then
	  GD=CommentArr(k)
	  TempStr=GetGDStar(GD,RST("Pid"))
	ElseIf Instr(CommentArr(k),"</i> ")<>0 Then
	  GD=split(CommentArr(k),"</i> ")(1)
	  TempStr=TempStr &" " & GetGDStar(GD,RST("Pid"))
	End If
  Next
  If Not KS.IsNUL(BSetting(49)) Then
    Dim DefaultGDArr:DefaultGDArr=Split(BSetting(49),",")
    For K=0 To Ubound(DefaultGDArr)
	 If Instr(TempStr,DefaultGDArr(k)&"��")=0 Then
	  TempStr=TempStr &" " & GetGDStar(DefaultGDArr(k),RST("Pid"))
	 End If
    Next
  End If
  Conn.Execute("Update KS_GuestComment Set AddDate=" & SQLNowString& ",Comment='" & TempStr& "' Where ID=" & RST("id"))
End Sub
Function GetGDStar(GD,pid)
  Dim RS,N,Star
  Set RS=Conn.Execute("Select Comment From KS_GuestComment Where Comment Like '%" & GD&"��<i>%' and UserID<>0 And Pid=" & PID)
  N=0:Star=0
  Do While Not RS.Eof
     N=n+1
	 Star=Star+KS.CutFixContent(RS(0),GD&"��<i>","</i>",0)
  RS.MoveNext
  Loop
  If N<>0 Then
  GetGDStar=GD&"��<i>" & Star/N & "</i>"
  Else
  GetGDStar=""
  End If
End Function
'����ҳȡ��������
Sub GetCommentPage()
    Dim Pid:Pid=KS.ChkClng(KS.S("pid"))
	If Pid=0 Then KS.Die "���س���!"
	if Boardid=0 then KS.Die "err|��������!"
	KS.LoadClubBoard()
	Dim Node:Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
	BSetting=Node.SelectSingleNode("@settings").text
	If KS.IsNul(BSetting) Then KS.Die "err|��������!"
	BSetting=Split(BSetting&"$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$","$")

    Dim RS:Set RS=Conn.Execute("Select * From KS_GuestComment Where Pid=" & PID & " Order By orderid,Id Desc")
	If Not RS.Eof Then
	   Dim Xml:Set XML=KS.RsToXml(RS,"row","")
	   KS.echo GetComments(XML,BoardID,pid,KS.ChkClng(BSetting(44)),check)
	End If
	Set RS=Nothing
End Sub
'ɾ������
Sub delcomment()
    If BoardID=0 Then KS.Die "����������!"
	If cbool(check)=false Then
		KS.Die "�Բ�����û�����õ�Ȩ��!"
	End If
	Conn.Execute("Delete From KS_GuestComment Where ID=" & KS.ChkClng(KS.S("ID")))
	KS.Die "success"
End Sub


'ͶƱ����
Sub dovote()
           Dim ID:ID=KS.ChkClng(KS.S("voteid"))
		   If Id=0 Then KS.Die "error!"
		   Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		   RS.Open "Select Top 1 * From KS_Vote Where id=" & id,CONN,1,3
		   If RS.Eof And RS.Bof Then
		     RS.Close:Set RS=Nothing
			 KS.Die "error!"
		   End If
		   
		   Dim LoopStr,XML,Node,Str,LC,Xstr,TotalVote
		   
		   'ͶƱ����
		     If RS("Status")="0" Then RS.Close:Set RS=Nothing : KS.Die Escape("��ͶƱ�ѹر�!")
			 Set KSUser=New UserCls
			 Dim LoginTF:LoginTF=KSUser.UserLoginChecked()
			 Dim GroupIds:GroupIds=RS("GroupIDs")
			 Dim TopicID:TopicID=RS("TopicID")
			 If RS("nmtp")="1" and LoginTF=false Then RS.Close:Set RS=Nothing:KS.Die Escape("�Բ���ֻ���¼��Ա����ͶƱ!")
			 If Not KS.IsNul(GroupIDs) And KS.FoundInArr(GroupIDs, KSUser.GroupID, ",")=False Then
			 	RS.Close:Set RS=Nothing
				KS.Die Escape("�Բ��������ڵĻ�Ա�鲻����ͶƱ!'")
			 End If
			 If RS("TimeLimit")="1" Then
			 	if now<RS("TimeBegin") then RS.Close:Set RS=Nothing: KS.Die Escape("�Բ��𣬸�ͶƱ��" & RS("TimeBegin") & "����!")
		        if now>RS("TimeEnd") then RS.Close:Set RS=Nothing : KS.Die Escape("�Բ��𣬸�ͶƱ����" & RS("TimeBegin") & "ֹͣ��")
			 End If
			 
			 
		     Dim VoteOption,ItemArr,I,UserName
			 VoteOption=KS.FilterIds(KS.S("VoteOption"))
			 If KS.IsNul(VoteOption) Then KS.Die Escape("��û��ѡ��ͶƱ��Ŀ!")
			 
			 Dim IPNum:IPNum=KS.ChkClng(RS("IpNum"))
			 Dim IPNums:IPNums=RS("IPNums")
			 If IpNums<>0 Then
			 	If Conn.Execute("Select Count(ID) From KS_PhotoVote Where UserIp='" & KS.GetIP & "' and ChannelID=-1 And InfoID='" & ID & "'")(0)>=IPNums  Then
				 RS.Close:Set RS=Nothing
	             KS.Die Escape("�Բ���ÿ��IP���ֻ��Ͷ" & IPNums & "��!")
	             End If
			 End If
			 If IpNum<>0 Then
			 	If Conn.Execute("Select Count(ID) From KS_PhotoVote Where Year(VoteTime)=" & Year(Now) & " and Month(VoteTime)=" & Month(Now) & " and Day(VoteTime)=" & Day(Now) & " and UserIp='" & KS.GetIP & "' and ChannelID=-1 And InfoID='" & ID & "'")(0)>=IPNum  Then
				 RS.Close:Set RS=Nothing
	             KS.Die Escape("�Բ���ÿ��IPһ�����ֻ��Ͷ" & IPNum & "��!")
	             End If
			 End If
			 
			 ItemArr=Split(VoteOption,",")
		     set XML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			 XML.async = false
			 XML.setProperty "ServerHTTPRequest", true 
			 XML.load(Server.MapPath(KS.Setting(3)&"Config/voteitem/vote_" &id&".xml"))
			 For I=0 To Ubound(ItemArr)
				 Set Node=XML.DocumentElement.SelectSingleNode("voteitem[@id=" & KS.ChkClng(ItemArr(i)) & "]")
				 Node.childNodes(1).text=KS.ChkClng(Node.childNodes(1).text)+1
				 XML.Save(Server.MapPath(KS.Setting(3)&"Config/voteitem/vote_" &id&".xml"))
			 Next
			 If LoginTF=False Then UserName="�ο�" Else UserName=KSUser.UserName
			 Conn.Execute("Insert Into [KS_PhotoVote]([ChannelID],[ClassID],[InfoID],[VoteTime],[UserName],[UserIP],[VoteOptions]) Values(-1,'-1','" & ID & "'," & SqlNowString & ",'" & UserName & "','" & KS.GetIP() & "','" & VoteOption & "')")
		     RS("VoteNums")=RS("VoteNums")+1
			 Dim VoteUserList:VoteUserList=RS("VoteUserList")
			 If KS.FoundInArr(VoteUserList,UserName,",")=false Then
			   If Instr(VoteUserList,",")=0 Then
			    RS("VoteUserList")=UserName
			   Else
			    RS("VoteUserList")=VoteUserList&"," & UserName
			   End If
			 End If
			 RS.Update
			 RS.Close:Set RS=Nothing

			 Application(KS.SiteSN&"_Configvoteitem/vote_"&ID)=empty
			 KS.Die "success@@@"&escape(GetVote(TopicID,XML))
end sub

sub fav()
  Dim KSUser:Set KSUser=New UserCls
  If Cbool(KSUser.UserLoginChecked)=false Then 
    KS.Die "���ȵ�¼��"
	exit sub
  else
    dim rs:set rs=conn.execute("select top 1 id from ks_guestbook where ID=" & TopicID)
	if rs.eof and rs.bof then 
	 rs.close :set rs=noting
	 KS.Die "��������"
	end if
	rs.close
	rs.open "select top 1 * From  KS_AskFavorite where username='" & KSUser.UserName & "' and typeflag=1 and topicid=" & TopicID,conn,1,3
	if not rs.eof then
	  rs.close:set rs=nothing
	  KS.Die "���ղع���!"
	else
	  rs.addnew
	   rs("username")=KSUser.UserName
	   rs("topicid")=topicid
	   rs("typeflag")=1
	   rs("FavorTime")=now
	  rs.update
	end if 
	rs.close:set rs=nothing
	ks.die "success"
  end if
end sub

'ɾ��ָ���û���ȫ������,���ؼ�������
sub delusertopic()
 	Dim KSUser:Set KSUser=New UserCls
	If Cbool(KSUser.UserLoginChecked)=false Then 
	 KS.Die "err|�Բ���,��û�д˲���Ȩ��!"
	end if
	if KS.ChkClng(KSUser.GroupId)<>1 Then
  	 KS.Die "err|�Բ���,ֻ�й���Ա�д�Ȩ��!"
    end if
	Dim DelType:DelType=KS.ChkClng(KS.S("DelType"))
	Dim RZM:RZM=UnEscape(KS.S("RZM"))
	If DelType<>0 And RZM<>SiteManageCode Then
	 KS.Die "err|�Բ������������֤�벻��ȷ��"
	End If
	
	'Dim KSLoginCls:Set KSLoginCls = New LoginCheckCls1
    'If KSLoginCls.Check=false Then
  	' KS.echo "<script>alert('�Բ���,Ϊ�˰�ȫ���������Ա�����ȵ�¼��̨�ſ���ִ�д˲���!');<//script>"
	' Response.Redirect "../" & KS.Setting(89) &"login.asp"
	' Exit Sub
    'End If
	
	Dim RS,TopicIDs,UserName:UserName=UnEscape(KS.S("UserName"))
	If KS.IsNul(UserName) Then
  	 KS.Die "err|�Բ���,����������!"
	End If
	If DelType<>1 Then
	  Conn.Execute("Update KS_GuestBook Set DelTF=1 Where UserName='" & UserName &"'")
	Else
	  Set RS=Conn.Execute("Select ID From KS_GuestBook Where UserName='" & UserName & "'")
	  Do While Not RS.Eof 
	    If TopicIDs="" Then
		 TopicIDs=RS(0)
		Else
	    TopicIDs=TopicIDs & "," & RS(0)
		End If
	    Conn.Execute("Delete From KS_GuestComment Where Tid=" & RS(0))
	  RS.MoveNext
	  Loop
	  RS.Close : Set RS=Nothing
	  Conn.Execute("Delete From KS_GuestBook Where UserName='" & UserName &"'")
	End If
	
	set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	TableXML.async = false
	TableXML.setProperty "ServerHTTPRequest", true 
	TableXML.load(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))
	Dim Url,N:N=0
    For Each Node In TableXML.DocumentElement.SelectNodes("item")
	  Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open "Select TopicID,ID From " & Node.SelectSingleNode("tablename").text & " Where UserName='" & UserName & "'",conn,1,1
	  Do While Not RS.Eof
	   n=n+1
	   Conn.Execute("Update KS_GuestBook Set TotalReplay=TotalReplay-1 Where ID=" & RS(0) & " And TotalReplay>0")
	   If DelType=1 Then
	    Conn.Execute("Delete From KS_GuestComment Where Pid=" & RS(1)&" And Tid=" & RS(0))
	   End If
	   RS.MoveNext
	  Loop
	  RS.Close
	  Set RS=Nothing
	  Dim Param:Param=" Where UserName='" & UserName &"'"
	  If TopicIDs<>"" Then Param=Param & " Or TopicID in(" & TopicIDs &")"
	  If DelType<>1 Then
	   Conn.Execute("update " & Node.SelectSingleNode("tablename").text & " set deltf=1 "& Param)
	  Else
	   Conn.Execute("Delete From " & Node.SelectSingleNode("tablename").text & Param)
	  End If
	Next
	If KS.S("N")="1" Then Url=KS.GetClubListUrl(BoardID) Else Url=KS.GetClubShowUrl(TopicID)
	KS.die "succ|��ϲ,����ɾ���û�[" & UserName & "]������������,�ۼ�"&n&"��!|" & url
end sub

sub support()
 dim sql
 if Not KS.IsNul(Request.Cookies("clubsupport" &ID)) then
   ks.echo "error"
   exit sub
 end if
 dim rs
 sql="select top 1 Support from " & getPostTable & " where id=" & id
 set rs=server.createobject("adodb.recordset")
 rs.open sql,conn,1,3
 if not rs.eof then
  rs(0)=ks.chkclng(rs(0))+1
  rs.update
 end if
 Response.Cookies("clubsupport" &ID)="ok"
 ks.echo rs(0)
 rs.close
 set rs=nothing
end sub
sub opposition()
 dim sql
 if Not KS.IsNul(Request.Cookies("clubsupport" &ID)) then
   ks.echo "error"
   exit sub
 end if
 dim rs
   sql="select top 1 opposition from " & getPostTable & " where id=" & id
 set rs=server.createobject("adodb.recordset")
 rs.open sql,conn,1,3
 if not rs.eof then
  rs(0)=ks.chkclng(rs(0))+1
  rs.update
 end if
 Response.Cookies("clubsupport" &ID)="ok"
 ks.echo rs(0)
 rs.close
 set rs=nothing
end sub


Sub SetBest()
    	Dim TopicIds:TopicIds=KS.FilterIds(KS.S("ID"))
		If TopicIds="" Then
		  KS.Die "�Բ���,��û��ѡ��Ҫ���þ���������!"
		End If

		If cbool(check)=false Then
		 KS.Die "�Բ�����û�����õ�Ȩ��!"
		End If
		Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select UserName,isbest,boardid,subject,id From KS_GuestBook Where ID in(" & TopicIds &")",conn,1,3
		If Not RS.Eof Then
		 Do While Not RS.Eof
			  ID=rs("id")
			  rs(1)=1
			  rs.update
			  Conn.Execute("Update KS_User Set BestTopicNum=BestTopicNum+1 Where UserName='" & rs(0) & "'")
			  boardid=rs(2)
			  if boardid<>0 and not KS.ISNul(rs(0)) then
				 KS.LoadClubBoard()
				 Dim Node:Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
				 BSetting=Node.SelectSingleNode("@settings").text
				 If Not KS.IsNul(BSetting) Then
				  If KS.ChkClng(Split(BSetting,"$")(33))<>0 Then
				  Conn.Execute("Update KS_User Set Prestige=Prestige+" & KS.ChkClng(Split(BSetting,"$")(33)) & " Where UserName='" & rs(0) &"'")
				  End If
				   If KS.ChkClng(Split(BSetting,"$")(6))>0 Then
					Call KS.ScoreInOrOut(rs(0),1,KS.ChkClng(Split(BSetting,"$")(6)),"ϵͳ","����̳��������[" & rs(3) & "]�����óɾ���!",0,0)
				   End If
				 End If
			  end if
		   rs.movenext
		 loop
		End If
		rs.close:set rs=nothing
		KS.Die "success"
	 End Sub
	 Sub SetTop()
		Dim TopicIds:TopicIds=KS.FilterIds(KS.S("ID"))
		Dim V:V=KS.ChkClng(KS.S("v"))
		If V=0 Then V=1
		If TopicIds="" Then
		  KS.Die "�Բ���,��û��ѡ��Ҫ�ö�������!"
		End If
		If check=false Then
		  KS.Die "�Բ�����û�����õ�Ȩ��!"
		End If
		
		Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select UserName,istop,boardid,subject,id From KS_GuestBook Where ID in(" & TopicIds &")",conn,1,3
		If Not RS.Eof Then
		  Do While Not RS.Eof
			  ID=rs("id")
			  rs(1)=v
			  rs.update
			  boardid=rs(2)
			  if boardid<>0 and not KS.ISNul(rs(0)) then
				 KS.LoadClubBoard()
				 Dim Node:Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
				 BSetting=Node.SelectSingleNode("@settings").text
				 If Not KS.IsNul(BSetting) Then
				  If KS.ChkClng(Split(BSetting,"$")(32))<>0 Then
				  Conn.Execute("Update KS_User Set Prestige=Prestige+" & KS.ChkClng(Split(BSetting,"$")(32)) & " Where UserName='" & rs(0) &"'")
				  End If
				   If KS.ChkClng(Split(BSetting,"$")(5))>0 Then
					Call KS.ScoreInOrOut(rs(0),1,KS.ChkClng(Split(BSetting,"$")(5)),"ϵͳ","����̳��������[" & rs(3) & "]�����ó��ö�!",0,0)
				   End If
				 End If
			  end if
			 RS.MoveNext
		  Loop
		End If
		rs.close:set rs=nothing
		MustReLoadTopTopic
		KS.Die "success"
	 End Sub
	 Sub CancelBest()
	    Dim TopicIds:TopicIds=KS.FilterIds(KS.S("ID"))
		If TopicIds="" Then
		  KS.Die "�Բ���,��û��ѡ��Ҫȡ������������!"
		End If

		If cbool(check)=false Then
		   KS.Die "�Բ�����û�����õ�Ȩ��!"
		End If
        Conn.Execute("Update KS_Guestbook set isbest=0 where id in(" & TopicIds &")")
		KS.Die "success"
	 End Sub
	 Sub CancelTop()
		Dim TopicIds:TopicIds=KS.FilterIds(KS.S("ID"))
		If TopicIds="" Then
		  KS.Die "�Բ���,��û��ѡ��Ҫȡ���ö�������!"
		End If
		If check=false Then
		  KS.Die "�Բ�����û�����õ�Ȩ��!"
		End If
        Conn.Execute("Update KS_Guestbook set istop=0 where id in(" & TopicIds &")")
		MustReLoadTopTopic
		KS.Die "success"
	 End Sub
	 
	 Sub delsubject()
		Dim TopicIds:TopicIds=KS.FilterIds(KS.S("ID"))
		Dim DelType:DelType=KS.ChkClng(KS.S("DelType"))
		Dim RZM:RZM=UnEscape(KS.S("RZM"))
		If TopicIds="" Then
		  KS.Die "�Բ���,��û��ѡ��Ҫɾ��������!"
		End If
		If cbool(check)=false Then
		  KS.Die "�Բ�����û��ɾ����Ȩ��!"
		End If
		
		If DelType<>0 And RZM<>SiteManageCode Then   '����ɾ�������֤��
		  KS.Die "�Բ������������֤������"
		End If
		
		If DelType<>1 Then     '�������վ���������������
		   Conn.Execute("Update KS_GuestBook  Set DelTF=1 Where ID In(" & TopicIDs &")")
		Else
			Dim TodayNum:TodayNum=0
			dim boardid,postTable,userName
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select UserName,boardid,subject,AddTime,PostTable,ID From KS_GuestBook Where ID in(" & TopicIds &")",conn,1,1
			If Not RS.Eof Then
			 Do While Not RS.Eof
				  id=RS("ID")
				  boardid=rs(1)
				  postTable=rs(4)
				  userName=rs(0)
				  If DateDiff("d",rs(3),Now)=0 Then
				   TodayNum=TodayNum+1
				  End If
				  If boardid<>0 then 
					 KS.LoadClubBoard()
					 Dim Node:Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
					 Dim LastPost,LastPostArr:LastPostArr=Split(Node.SelectSingleNode("@lastpost").text,"$")
					 
					 '������ҳ����������
					 If KS.ChkClng(LastPostArr(0))=ID Then
					   Dim RSNew:Set RSNew=Conn.Execute("Select top 1 ID,BoardID,Subject,AddTime From KS_GuestBook Where BoardID=" & boardid & " and verific=1 and id<>" & id & " order by id desc")
					   If Not RSNew.Eof Then
						 LastPost=RSNew(0) & "$" & RSNew(3) & "$" & Replace(left(RSNew(2),200),"$","") & "$$$$$$$$"
					   Else
						 LastPost="��$��$��$$$$$$$$"
					   End If
					   Conn.Execute("Update KS_GuestBoard Set LastPost='" & LastPost & "' Where ID=" & BoardID)
					   Node.SelectSingleNode("@lastpost").text=LastPost
					 End If
				  end if
				  
				  if not KS.ISNul(rs(0)) then
					 BSetting=Node.SelectSingleNode("@settings").text
					 If Not KS.IsNul(BSetting) Then
						 If KS.ChkClng(Split(BSetting,"$")(34))<>0 Then
						  Conn.Execute("Update KS_User Set Prestige=Prestige-" & KS.ChkClng(Split(BSetting,"$")(34)) & " Where UserName='" & rs(0) &"' and Prestige>0")
						 End If
					 
					   If KS.ChkClng(Split(BSetting,"$")(7))>0 Then
						Call KS.ScoreInOrOut(rs(0),2,KS.ChkClng(Split(BSetting,"$")(7)),"ϵͳ","����̳�����������[" & rs(2) & "]��ɾ��!",0,0)
					   End If
					 End If
				  end if
				  
				  Dim Num,replyNum:replyNum=Conn.Execute("Select count(id) from " & PostTable & " where topicid=" & id)(0)
				  TodayNum=TodayNum+Conn.Execute("Select count(id) from " & PostTable & " where topicid=" & id &" and datediff(" & DataPart_D & ",ReplayTime," & SqlNowString&")=0")(0)
				  Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				  Doc.async = false
				  Doc.setProperty "ServerHTTPRequest", true 
				  Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
				  Dim XMLDate:XMLDate=doc.documentElement.attributes.getNamedItem("date").text
				  Num=KS.ChkClng(doc.documentElement.attributes.getNamedItem("todaynum").text)-TodayNum
				  If Num<0 Then Num=0
				  doc.documentElement.attributes.getNamedItem("todaynum").text=Num
				  Num=KS.ChkClng(doc.documentElement.attributes.getNamedItem("postnum").text)-replyNum
				  If Num<0 Then Num=0
				  doc.documentElement.attributes.getNamedItem("postnum").text=Num
				  Num=KS.ChkClng(doc.documentElement.attributes.getNamedItem("topicnum").text)-1
				  If Num<0 Then Num=0
				  doc.documentElement.attributes.getNamedItem("topicnum").text=Num
				  
				  Conn.Execute("Update KS_GuestBoard Set TodayNum=TodayNum-" & TodayNum & " where id=" &boardid &" and todaynum>=" & TodayNum)
				  Conn.Execute("Update KS_GuestBoard Set PostNum=PostNum-" & replyNum -1& " where id=" &boardid &" and PostNum>=" & replyNum-1)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.selectSingleNode("row[@id=" & boardid & "]/@postnum").text=Conn.Execute("Select PostNum From KS_GuestBoard Where id=" & boardid)(0)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.selectSingleNode("row[@id=" & boardid & "]/@todaynum").text=Conn.Execute("Select TodayNum From KS_GuestBoard Where id=" & boardid)(0)
		
				  doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
					
					Conn.Execute("update KS_User set postNum=postNum-1 where userName='" & UserName & "' and postNum>0")
					Conn.Execute("delete from KS_Guestbook where id=" & ID)
					Conn.Execute("delete from " & PostTable & " where TopicID=" & ID)
					Conn.Execute("delete from KS_GuestComment Where Tid=" & ID)
					Conn.Execute("delete from KS_UploadFiles where ID=" & ID & " and channelid=9994")
			  RS.MoveNext
			Loop 
			End If
			rs.close:set rs=nothing
		End If
		KS.Echo "success"
	 End Sub
	 
	 Sub delreply()
		If cbool(check)=false Then
		  KS.Die "�Բ�����û�����õ�Ȩ��!"
		  Exit Sub
		End If
		Dim DelType:DelType=KS.ChkClng(KS.S("DelType"))
		Dim RZM:RZM=UnEscape(KS.S("RZM"))
		If ID=0 or KS.ChkClng(KS.S("ReplyID"))=0 Then
		  KS.Die "�Բ���,��û��ѡ��Ҫɾ���Ļظ�!"
		End If

		If DelType<>0 And RZM<>SiteManageCode Then   '����ɾ�������֤��
		  KS.Die "�Բ������������֤������"
		End If
		
		dim boardid,postTable,userName
		Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select top 1 UserName,boardid,subject,TotalReplay,postTable From KS_GuestBook Where ID=" & ID,conn,1,1
		If Not RS.Eof Then
		  if rs(3)>0 then 
		    conn.execute("update ks_guestbook set TotalReplay=TotalReplay-1 where id=" & id & " and TotalReplay>=1")
		  end if
		  boardid=rs(1)
		  postTable=rs(4)
		  userName=rs(0)
		  
		  Dim ReplayTime:ReplayTime=Conn.Execute("Select top 1 ReplayTime From " & postTable &" where ID=" & KS.ChkClng(KS.S("ReplyID")))(0)
		  '����������
		  Dim Num
		  Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		  Doc.async = false
		  Doc.setProperty "ServerHTTPRequest", true 
		  Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
		  Dim XMLDate:XMLDate=doc.documentElement.attributes.getNamedItem("date").text
		  If DateDiff("d",xmldate,ReplayTime)=0 Then
		    Conn.Execute("Update KS_GuestBoard Set TodayNum=TodayNum-1 where id=" &boardid &" and todaynum>0")
		    Num=KS.ChkClng(doc.documentElement.attributes.getNamedItem("todaynum").text)-1
			If Num<0 Then Num=0
		    doc.documentElement.attributes.getNamedItem("todaynum").text=Num
			
			Application(KS.SiteSN&"_ClubBoard").DocumentElement.selectSingleNode("row[@id=" & boardid & "]/@todaynum").text=Conn.Execute("Select TodayNum From KS_GuestBoard Where id=" & boardid)(0)
          End If
		    Conn.Execute("Update KS_GuestBoard Set PostNum=PostNum-1 where id=" &boardid &" and PostNum>0")
		    Num=KS.ChkClng(doc.documentElement.attributes.getNamedItem("postnum").text)-1
			If Num<0 Then Num=0
		    doc.documentElement.attributes.getNamedItem("postnum").text=Num
			doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
			Application(KS.SiteSN&"_ClubBoard").DocumentElement.selectSingleNode("row[@id=" & boardid & "]/@postnum").text=Conn.Execute("Select PostNum From KS_GuestBoard Where id=" & boardid)(0)

		  if boardid<>0 and not KS.ISNul(rs(0)) then
		     KS.LoadClubBoard()
			 Dim Node:Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
			 BSetting=Node.SelectSingleNode("@settings").text
			 If Not KS.IsNul(BSetting) Then
			     Dim ReplyUser:ReplyUser=Conn.Execute("Select top 1 UserName From " & postTable &" where ID=" & KS.ChkClng(KS.S("ReplyID")))(0)
			     If KS.ChkClng(Split(BSetting,"$")(35))<>0 Then
					  Conn.Execute("Update KS_User Set Prestige=Prestige-" & KS.ChkClng(Split(BSetting,"$")(35)) & " Where UserName='" & ReplyUser &"' and Prestige>0")
				 End If
			   If KS.ChkClng(Split(BSetting,"$")(8))>0 Then
			    Call KS.ScoreInOrOut(ReplyUser,2,KS.ChkClng(Split(BSetting,"$")(8)),"ϵͳ","����̳������[" & rs(2) & "]�Ļظ���ɾ��!",0,0)
			   End If
			 End If
		  end if

		End If
		rs.close:set rs=nothing
		If DelType=1  Then  '����ɾ��
		 Conn.Execute("delete from " & postTable & " where topicid=" & id & " and ID=" & KS.ChkClng(KS.S("ReplyID")))
		 Conn.Execute("delete from KS_GuestComment Where tid=" & id & " and PID=" & KS.ChkClng(KS.S("ReplyID")))
		Else
		 Conn.Execute("update " & postTable & " set deltf=1 where topicid=" & id & " and ID=" & KS.ChkClng(KS.S("ReplyID")))
		End If
		KS.Echo "success"
	 End Sub
	 
	 sub verify()
		If check=false Then
		  Response.Write"<script>alert('�Բ�����û�����õ�Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
		Conn.Execute("update " & getPostTable &" set verific=1 where ID=" & KS.ChkClng(KS.S("ReplyID")))
		Response.Redirect request.servervariables("http_referer")
	 end sub

sub Locked()
		If check=false Then
		  KS.Echo "�Բ�����û��������Ȩ��!"
		  Exit Sub
		End If
		Conn.Execute("update ks_guestbook set verific=2 where ID=" & KS.ChkClng(KS.S("ID")))
		Conn.Execute("update " & getPostTable &" set verific=2 where parentid=0 and topicid=" & KS.ChkClng(KS.S("TopicID")))
		KS.Echo "success"
end sub
sub unlocked()
	If check=false Then
		  KS.Echo "�Բ�����û��������Ȩ��!"
		  Exit Sub
	End If
		Conn.Execute("update ks_guestbook set verific=1 where ID=" & KS.ChkClng(KS.S("ID")))
		Conn.Execute("update " & getPostTable &" set verific=1 where parentid=0 and topicid=" & KS.ChkClng(KS.S("TopicID")))
		KS.Echo "success"
end sub
 sub replyLock()
		If check=false Then
		  Response.Write"<script>alert('�Բ�����û��������Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
		dim rs:set rs=server.CreateObject("adodb.recordset")
		rs.open "select top 1 * from " & getPostTable & " Where ID="& KS.ChkClng(KS.S("replyID")),CONN,1,3
		If Not RS.Eof then
		rs("verific")=2
		rs.update
		if rs("parentid")=0 then
		conn.execute("update ks_guestbook set verific=2 where id=" & rs("topicid"))
		end if
		end if
		rs.close : set rs=nothing

		'Conn.Execute("update " & getPostTable &" set verific=2 where ID=" & KS.ChkClng(KS.S("replyID")))
		Response.Redirect request.servervariables("http_referer")
end sub
 sub replyunlock()
		If check=false Then
		  Response.Write"<script>alert('�Բ�����û��������Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
		dim rs:set rs=server.CreateObject("adodb.recordset")
		rs.open "select top 1 * from " & getPostTable & " Where ID="& KS.ChkClng(KS.S("replyID")),CONN,1,3
		If Not RS.Eof then
		rs("verific")=1
		rs.update
		if rs("parentid")=0 then
		conn.execute("update ks_guestbook set verific=1 where id=" & rs("topicid"))
		end if
		end if
		rs.close : set rs=nothing
		
		'Conn.Execute("update " & getPostTable &" set verific=1 where ID=" & KS.ChkClng(KS.S("replyID")))
		Response.Redirect request.servervariables("http_referer")
end sub
sub lockuser()
		If check=false Then
		  Response.Write"<script>alert('�Բ�����û�������û���Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
		Conn.Execute("update KS_User set lockonclub=1 where UserID=" & KS.ChkClng(KS.S("UserId")))
		Response.Redirect request.servervariables("http_referer")
end sub
sub UNlockuser()
		If check=false Then
		  Response.Write"<script>alert('�Բ�����û�������û���Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
		Conn.Execute("update KS_User set lockonclub=0 where UserID=" & KS.ChkClng(KS.S("UserId")))
		Response.Redirect request.servervariables("http_referer")
end sub

sub movetopic()
		Dim TopicIds:TopicIds=KS.FilterIds(KS.S("ID"))
		If TopicIds="" Then
		  KS.Die "�Բ���,��û��ѡ��Ҫ�ƶ�������!"
		End If

		If check=false Then
		  Response.Write"<script>alert('�Բ�����û���ƶ����ӵ�Ŀ������Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
		dim rs,oldboardid,replynum,boardid
		boardid=KS.ChkClng(KS.S("Boardid"))
		if boardid=0 then
		 KS.AlertHintScript "�����������!"
		end if
		set rs=server.createobject("adodb.recordset")
		rs.open "select top 100 * from ks_guestbook where id in(" & TopicIds & ")",conn,1,1
		if not rs.eof then
		 Do While Not RS.Eof
			 oldboardid=rs("boardid")
			 if oldboardid=boardid then
			  rs.close
			  set rs=nothing
			   Response.Redirect request.servervariables("http_referer")
			 end if
			 replynum=conn.execute("select count(id) from " & rs("postTable") & " where topicid=" & rs("id"))(0)
			 Conn.Execute("Update KS_GuestBoard set PostNum=PostNum-" & replynum &",TopicNum=TopicNum-1 where PostNum>" & replynum & " and id=" & oldboardid)
			 Conn.Execute("Update KS_GuestBoard set PostNum=PostNum+" & replynum &",TopicNum=TopicNum+1 where id=" & boardid)
			 Conn.Execute("update ks_guestbook set BoardID=" & Boardid & " where ID=" & rs("id"))
		 RS.MoveNext
		 Loop
			 rs.close
			 set rs=nothing
		  KS.AlertHintscript "��ϲ�������ƶ��ɹ�!"
		end if
		rs.close
		set rs=nothing
		Response.Redirect request.servervariables("http_referer")
	 end sub


'�������
sub verifictopic()
    dim id:id=KS.FilterIds(KS.S("ID"))
	If Id="" Then
	   KS.Die "û��ѡ��Ҫ��˵�����!"
	End If
	If check=false Then
	   KS.Die "�Բ�����û���������Ȩ��!"
	End If
	Conn.Execute("Update KS_GuestBook Set Verific=" & KS.ChkClng(KS.S("V")) &" Where Id in(" & Id & ")")
	KS.Die "success"
End Sub


'�����������Ա	 
function check()
	 	Dim KSLoginCls
		Set KSLoginCls = New LoginCheckCls1
		If KSLoginCls.Check=true Then
		  check=true
		  Exit function
		else
		    master=LFCls.GetSingleFieldValue("select top 1 master from ks_guestboard where id=" & BoardID)
			Dim KSUser:Set KSUser=New UserCls
			If Cbool(KSUser.UserLoginChecked)=false Then 
			  check=false
			  exit function
			elseif KSUser.GetUserInfo("ClubSpecialPower")=2 Or KSUser.GetUserInfo("ClubSpecialPower")=1 Then
			  check=true
			  exit function
			else
			   check=KS.FoundInArr(master, KSUser.UserName, ",")
			End If
		end if
End function

			
			

%>