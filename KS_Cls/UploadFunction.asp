<%
'KesionCMS V7 ��������ദ���༰���� �޸���2010-6-28 by xiaolin
Const NoAllowExt = "asa|asax|ascs|ashx|asmx|asp|aspx|axd|cdx|cer|cfm|config|cs|csproj|idc|licx|rem|resources|resx|shtm|shtml|soap|stm|vb|vbproj|vsdisco|webinfo"    '�������ϴ�����
Const NeedCheckFileMimeExt = "gif|jpg|jpeg|jpe|bmp|png|swf|mid|mp3|wmv|asf|avi|mpg|ram|rm|ra|rar|exe|doc|zip" '������Ҫ����Ƿ�α����ļ�����

Dim KS:Set KS=New PublicCls

'��Ӹ���
Sub AddAnnexToDB(ChannelID,Username,TempFileStr,FileTitles,ClassID)
	'д��KS_UploadFiles���ݿ�
	Dim FileArr,n,FileIDS,MaxID,TitleArr
	FileArr=Split(TempFileStr,"|")
	TitleArr=Split(Replace(FileTitles,"'",""),"|")
	For N=0 To Ubound(FileArr)
	  If Not KS.IsNul(FileArr(n)) Then
	       If Right(lcase(FileArr(n)),3)<>"gif" and Right(lcase(FileArr(n)),3)<>"bmp" and Right(lcase(FileArr(n)),3)<>"jpg" and Right(lcase(FileArr(n)),3)<>"png" and Right(lcase(FileArr(n)),4)<>"jpeg" Then
								 Conn.Execute("Insert Into [KS_UploadFiles](ChannelID,InfoID,Title,FileName,IsAnnex,UserName,Hits,AddDate,ClassID) values(" &ChannelID &",0,'" & TitleArr(n) & "','" & FileArr(n) & "',1,'" & UserName & "',0," & SQLNowString&"," & ClassID & ")")
								 MaxID=Conn.Execute("Select Max(ID) From  [KS_UploadFiles]")(0)
								 If FileIds="" Then
								   FileIds=MaxID
								 Else
								   FileIds=FileIds & "," & MaxID
								 End If
			Else
			   MaxID=0
			End If
		 Response.Write("parent.InsertFileFromUp('" & FileArr(n) &"'," & KS.GetFieSize(Server.MapPath(Replace(FileArr(n),KS.Setting(2),""))) & "," & MaxID & ",'" & TitleArr(n) & "');")
	  End If
	Next
	If Session("UploadFileIDs")="" Then
	  Session("UploadFileIDs")=FileIds
	Else
	  Session("UploadFileIDs")=Session("UploadFileIDs") & "," & FileIds
	End If
End Sub

Function CheckUpFile(KSUser,MustCheckSpaceSize,UpFileObj,FormPath,Path,FileSize,AllowExtStr,ByRef U_FileSize,ByRef TempFileStr,ByRef FileTitles,ByRef CurrNum,CreateThumbsFlag,DefaultThumb,ThumbPathFileName)
			Dim ErrStr,NoUpFileTF,FsoObj,FileName,FileExtName,FileContent,SameFileExistTF,FormName,AutoReName,BasicType
			AutoReName = KS.ChkClng(UpFileObj.Form("AutoRename"))
			BasicType=KS.ChkClng(UpFileObj.Form("BasicType")) 
			NoUpFileTF = True
			ErrStr = ""
			Set FsoObj = KS.InitialObject(KS.Setting(99))
			For Each FormName in UpFileObj.File
				SameFileExistTF = False
				FileName = UpFileObj.File(FormName).FileName
				If NoIllegalStr(FileName)=False Then ErrStr=ErrStr&"�ļ����ϴ�����ֹ��\n"
				FileExtName = UpFileObj.File(FormName).FileExt
				If Instr(FileExtName,"?")<>0 Then FileExtName=right("00" & FileExtName,4) '��ֹswfupload���������봦��
				If Instr(FileExtName,"?")<>0 Then FileExtName=right("00" & FileExtName,3)
				FileContent = UpFileObj.File(FormName).FileData
				U_FileSize=UpFileObj.File(FormName).FileSize
				Dim FileType:FileType=UpFileObj.File(FormName).FileType
				'�Ƿ���������ļ�
				if U_FileSize > 1 then
					NoUpFileTF = False
					ErrStr = ""
					if U_FileSize > CLng(FileSize)*1024 then
						ErrStr = ErrStr & FileName & "�ļ��ϴ�ʧ��\n���������ƣ����ֻ���ϴ�" & FileSize & "K���ļ�\n"
					end if
					
					If MustCheckSpaceSize=true Then
						If BasicType<>9994 Then
						 IF KS.ChkClng(KS.GetFolderSize(KSUser.GetUserFolder(ksuser.username))/1024+UpFileObj.File(FormName).FileSize/1024)>=KS.ChkClng(KSUser.GetUserInfo("SpaceSize")) Then
						   CheckUpFile="�ϴ�ʧ��1�����Ŀ��ÿռ䲻����"
						   Exit Function
						 End If
						End If
					End If
					
					if AutoRename = "0" then
						If FsoObj.FileExists(Path & FileName) = True  then
							ErrStr = ErrStr & FileName & "�ļ��ϴ�ʧ��,����ͬ���ļ�\n"
						else
							SameFileExistTF = True
						end if
					else
						SameFileExistTF = True
					End If
					if CheckFileType(AllowExtStr,FileExtName) = False then
						ErrStr = ErrStr & FileName & "�ļ��ϴ�ʧ��,�ļ����Ͳ�����\n�����������" + AllowExtStr + "\n"
					end if
					If Left(LCase(FileType), 5) = "text/" and KS.FoundInArr(NeedCheckFileMimeExt,FileExtName,"|")=true Then
					 ErrStr = ErrStr & FileName & "�ļ��ϴ�ʧ��\nΪ��ϵͳ��ȫ���������ϴ����ı��ļ�α���ͼƬ�ļ���\n"
					End If
					If instr(FileName,";")>0 or instr(lcase(FileName),".asp")>0 or instr(lcase(FileName),".php")>0 or instr(lcase(FileName),".cdx")>0 or instr(lcase(FileName),".asa")>0 or instr(lcase(FileName),".cer")>0 or instr(lcase(FileName),".cfm")>0 or instr(lcase(FileName),".jsp")>0 then
						ErrStr = ErrStr & FileName & "�ļ��ϴ�ʧ��,�ļ������Ϸ�\n"
					end if
					
					if ErrStr = "" then
						if SameFileExistTF = True then
							CheckUpFile = CheckUpFile & SaveFile(KSUser,UpFileObj,FormPath,Path,FormName,AutoReName,TempFileStr,FileTitles,CurrNum,CreateThumbsFlag,DefaultThumb,ThumbPathFileName)
						else
							CheckUpFile = CheckUpFile &SaveFile(KSUser,UpFileObj,FormPath,Path,FormName,"",TempFileStr,FileTitles,CurrNum,CreateThumbsFlag,DefaultThumb,ThumbPathFileName)

						end if
					else
						CheckUpFile = CheckUpFile & ErrStr
					end if
				end if
			Next
			Set FsoObj = Nothing
			if NoUpFileTF = True then
				CheckUpFile = "û���ϴ��ļ�"
			end if
End Function

Function NoIllegalStr(Byval FileNameStr)
			Dim Str_Len,Str_Pos
			Str_Len=Len(FileNameStr)
			Str_Pos=InStr(FileNameStr,Chr(0))
			If Str_Pos=0 or Str_Pos=Str_Len then
				NoIllegalStr=True
			Else
				NoIllegalStr=False
			End If
End function
Function DealExtName(Byval UpFileExt)
			If IsEmpty(UpFileExt) Then Exit Function
			DealExtName = Lcase(UpFileExt)
			DealExtName = Replace(DealExtName,Chr(0),"")
			DealExtName = Replace(DealExtName,".","")
			DealExtName = Replace(DealExtName,"'","")
			DealExtName = Replace(DealExtName,"asp","")
			DealExtName = Replace(DealExtName,"asa","")
			DealExtName = Replace(DealExtName,"aspx","")
			DealExtName = Replace(DealExtName,"cer","")
			DealExtName = Replace(DealExtName,"cdx","")
			DealExtName = Replace(DealExtName,"htr","")
			DealExtName = Replace(DealExtName,"php","")
End Function
Function CheckFileType(AllowExtStr,FileExtName)
	 Dim i,AllowArray
	 AllowArray = Split(AllowExtStr,"|")
	 FileExtName = LCase(FileExtName)
	 CheckFileType = False
	 For i = LBound(AllowArray) to UBound(AllowArray)
				if LCase(AllowArray(i)) = LCase(FileExtName) then
					CheckFileType = True
				end if
	 Next
	 If KS.FoundInArr(LCase(NoAllowExt),FileExtName,"|")=true Then
		CheckFileType = False
	 end if
End Function

Function SaveFile(KSUser,UpFileObj,FormPath,FilePath,FormNameItem,AutoNameType,TempFileStr,FileTitles,CurrNum,CreateThumbsFlag,DefaultThumb,ThumbPathFileName)
			Dim FileName,FileExtName,FileContent,FormName,RandomFigure,n,RndStr,Title,BasicType,ChannelID,UpType,ThumbFileName,AddWaterFlag
		    BasicType=KS.ChkClng(UpFileObj.Form("BasicType"))      
		    ChannelID=KS.ChkClng(UpFileObj.Form("ChannelID")) 
			UpType=UpFileObj.Form("UpType")
			AddWaterFlag = UpFileObj.Form("AddWaterFlag")
			If AddWaterFlag <> "1" Then	'�����Ƿ�Ҫ���ˮӡ���
				AddWaterFlag = "0"
			End if

			Randomize 
			n=2* Rnd+10
			RndStr=KS.MakeRandom(n)
			RandomFigure = CStr(Int((99999 * Rnd) + 1))
			FileName = UpFileObj.File(FormNameItem).FileName
			FileExtName = UpFileObj.File(FormNameItem).FileExt
			FileExtName = DealExtName(FileExtName)
			If Instr(FileExtName,"?")<>0 Then FileExtName=right("00" & FileExtName,4) '��ֹswfupload���������봦��
			If Instr(FileExtName,"?")<>0 Then FileExtName=right("00" & FileExtName,3)
			FileContent = UpFileObj.File(FormNameItem).FileData
			
			Title=replace(FileName,"." & FileExtName,"") 'ԭ����
			If BasicType=9999 Then   'ͷ��
			   FileName=KSUser.GetUserInfo("UserID") & ".jpg"
			Else
				select case AutoNameType 
				  case "1"
					FileName= "����" & FileName
				  case "2"
					FileName= RndStr&"."&FileExtName
				  Case "3"
					FileName= RndStr & FileName
				  case "4"
					FileName= Year(Now())&Right("0"&Month(Now()),2)&Right("0"&Day(Now()),2)&Right("0"&Hour(Now()),2)&Right("0"&Minute(Now()),2)&Right("0"&Second(Now()),2)&RandomFigure&"."&FileExtName
				  case else
					FileName=FileName
				End Select
			End If
		   UpFileObj.File(FormNameItem).SaveToFile FilePath  &FileName
		   
		   
		   
		   '======================���Ӽ���ļ������Ƿ�Ϸ�===================================
		   Dim CheckContent:CheckContent=CheckFileContent(FormPath  &FileName,UpFileObj.File(FormNameItem).FileSize /1024)
		   If KS.IsNul(CheckContent) Then
			'==================================================================================
			
		   
		   TempFileStr=TempFileStr & FormPath & FileName & "|"
		   FileTitles=FileTitles & Title & "|"
		  Dim T:Set T=New Thumb
		  CurrNum=CurrNum+1
		  IF CreateThumbsFlag=true and  (cint(CurrNum)=cint(DefaultThumb) or BasicType=2 or (Channelid=5 and UpType="ProImage")) Then
		  	  If KS.TBSetting(0)=0 then
			   if ThumbPathFileName="" then
			   ThumbPathFileName=FormPath &FileName
			   Else
			   ThumbPathFileName=ThumbPathFileName & "|" & FormPath & FileName
			   End If
			  Else
				ThumbFileName=split(FileName,".")(0)&"_S."&FileExtName
				Dim CreateTF:CreateTF=T.CreateThumbs(FilePath & FileName,FilePath & ThumbFileName)
				if CreateTF=true Then
				 'ȡ������ͼ��ַ
				 if ThumbPathFileName="" then
				 ThumbPathFileName=FormPath & ThumbFileName
				 else
				 ThumbPathFileName=ThumbPathFileName & "|" & FormPath & ThumbFileName
				end if
			   Else
				 'ȡ������ͼ��ַ
				 if ThumbPathFileName="" then
				 ThumbPathFileName=FormPath & FileName
				 else
				 ThumbPathFileName=ThumbPathFileName & "|" & FormPath & FileName
				 end if
			   End If
			  End If
		  End if
		  If AddWaterFlag = "1" Then   '�ڱ���õ�ͼƬ�����ˮӡ
				call T.AddWaterMark(FilePath  & FileName)
		  End if
		  Set T=Nothing
		  
		'======================���Ӽ���ļ������Ƿ�Ϸ�===================================
	     Else
		  SaveFile=CheckContent
		 End If
		'==================================================================================
End Function

'����ļ����ݵ��Ƿ�Ϸ�
Function  CheckFileContent(byval path,byval filesize)
		     dim kk,NoAllowExtArr
			 NoAllowExtArr=split(NoAllowExt,"|")
			 for kk=0 to ubound(NoAllowExtArr)
					   if instr(replace(lcase(path),lcase(KS.Setting(2)),""),"." & NoAllowExtArr(kk))<>0 then
					    call KS.DeleteFile(path)
					    CheckFileContent= "�ļ��ϴ�ʧ��,�ļ������Ϸ�"
						Exit Function
					   end if
			 Next

		    if filesize>50 then exit function  '����1000K�������
		    on error resume next
		    Dim findcontent,regEx,foundtf
			findcontent=KS.ReadFromFile(Replace(path,KS.Setting(2),""))
			if err then exit function:err.clear
			foundtf=false
			Set regEx = New RegExp
			regEx.IgnoreCase = True
			regEx.Global = True
			regEx.Pattern = "@\s*LANGUAGE\s*=\s*[""]?\s*(vbscript|jscript|javascript).encode\b"
			If regEx.Test(findcontent) Then
				foundtf=true
			end if	
			
			regEx.Pattern = "execute\s*request"
			If regEx.Test(findcontent) Then
				foundtf=true
			end if
			
			regEx.Pattern = "executeglobal\s*request"
			If regEx.Test(findcontent) Then
				foundtf=true
			end if
			regEx.Pattern = "<script.*runat.*server(\n|.)*execute(\n|.)*<\/script>"
			If regEx.Test(findcontent) Then
				foundtf=true
			end if
			regEx.Pattern = "\<%(.|\n)*%\>"
			If regEx.Test(findcontent) Then
				foundtf=true
			end if
			If Instr(lcase(findcontent),"scripting.filesystemobject")<>0 or instr(lcase(findcontent),"adodb.stream")<>0 Then
			foundtf=true
			End If
			
			set regEx=nothing
			
			if foundtf then
			   KS.DeleteFile(path)
			   CheckFileContent="ϵͳ��鵽���ϴ����ļ����ܴ���Σ�մ��룬�������ϴ���"
			end if
			
End Function


Dim UpFileStream
Class UpFileClass
	Dim Form,File,Err 
	Private Sub Class_Initialize
		Err = -1
	End Sub
	Private Sub Class_Terminate  
		'�������������
		If Err < 0 Then
			Form.RemoveAll
			Set Form = Nothing
			File.RemoveAll
			Set File = Nothing
			UpFileStream.Close
			Set UpFileStream = Nothing
		End If
	End Sub
	
	Public Property Get ErrNum()
		ErrNum = Err
	End Property
	
	Public Sub GetData ()
		'�������
		Dim RequestBinData,sSpace,bCrLf,sObj,iObjStart,iObjEnd,tStream,iStart,oFileObj
		Dim iFileSize,sFilePath,sFileType,sFormValue,sFileName
		Dim iFindStart,iFindEnd
		Dim iFormStart,iFormEnd,sFormName
		Dim KS:Set KS=New PublicCls
		'���뿪ʼ
		If Request.TotalBytes < 1 Then  '���û�������ϴ�
			Err = 1
			Exit Sub
		End If
		Set Form = KS.InitialObject ("Scripting.Dictionary")
		Form.CompareMode = 1
		Set File = KS.InitialObject ("Scripting.Dictionary")
		File.CompareMode = 1
		Set tStream = KS.InitialObject ("ADODB.Stream")
		Set UpFileStream = KS.InitialObject ("ADODB.Stream")
		UpFileStream.Type = 1
		UpFileStream.Mode = 3
		UpFileStream.Open
		dim ReadedBytes,ChunkBytes
		ReadedBytes=0
		ChunkBytes=1024*100 '100K�ֿ��ϴ����� 
		Do   While   ReadedBytes   <   Request.TotalBytes   
		UpFileStream.Write   Request.BinaryRead(ChunkBytes)    
		ReadedBytes   =   ReadedBytes   +   ChunkBytes   
		If   ReadedBytes   >   Request.TotalBytes   Then   ReadedBytes   =   Request.TotalBytes   
		Loop
			
		'UpFileStream.Write (Request.BinaryRead(Request.TotalBytes))
		UpFileStream.Position = 0
		RequestBinData=UpFileStream.Read 
		iFormEnd = UpFileStream.Size
		bCrLf = ChrB (13) & ChrB (10)
		'ȡ��ÿ����Ŀ֮��ķָ���
		sSpace=MidB (RequestBinData,1, InStrB (1,RequestBinData,bCrLf)-1)
		iStart=LenB (sSpace)
		iFormStart = iStart+2
		'�ֽ���Ŀ
		Do
			iObjEnd=InStrB(iFormStart,RequestBinData,bCrLf & bCrLf)+3
			tStream.Type = 1
			tStream.Mode = 3
			tStream.Open
			UpFileStream.Position = iFormStart
			UpFileStream.CopyTo tStream,iObjEnd-iFormStart
			tStream.Position = 0
			tStream.Type = 2
			tStream.CharSet = "gb2312"
			sObj = tStream.ReadText      
			'ȡ�ñ���Ŀ����
			iFormStart = InStrB (iObjEnd,RequestBinData,sSpace)-1
			iFindStart = InStr (22,sObj,"name=""",1)+6
			iFindEnd = InStr (iFindStart,sObj,"""",1)
			sFormName = Mid  (sObj,iFindStart,iFindEnd-iFindStart)
			'������ļ�
			If InStr  (45,sObj,"filename=""",1) > 0 Then
				Set oFileObj = new FileObj_Class
				'ȡ���ļ�����
				iFindStart = InStr (iFindEnd,sObj,"filename=""",1)+10
				iFindEnd = InStr (iFindStart,sObj,"""",1)
				sFileName = Mid (sObj,iFindStart,iFindEnd-iFindStart)
				oFileObj.FileName = Mid (sFileName,InStrRev (sFileName, "\")+1)
				oFileObj.FilePath = Left (sFileName,InStrRev (sFileName, "\"))
				oFileObj.FileExt = Mid (sFileName,InStrRev (sFileName, ".")+1)
				iFindStart = InStr (iFindEnd,sObj,"Content-Type: ",1)+14
				iFindEnd = InStr (iFindStart,sObj,vbCr)
				oFileObj.FileType = Mid  (sObj,iFindStart,iFindEnd-iFindStart)
				oFileObj.FileStart = iObjEnd
				oFileObj.FileSize = iFormStart -iObjEnd -2
				oFileObj.FormName = sFormName
				File.add sFormName,oFileObj
			else
				'����Ǳ���Ŀ
				tStream.Close
				tStream.Type = 1
				tStream.Mode = 3
				tStream.Open
				UpFileStream.Position = iObjEnd 
				UpFileStream.CopyTo tStream,iFormStart-iObjEnd-2
				tStream.Position = 0
				tStream.Type = 2
				tStream.CharSet = "gb2312"
				sFormValue = tStream.ReadText
				If Form.Exists(sFormName)Then
					Form (sFormName) = Form (sFormName) & ", " & sFormValue
				else
					form.Add sFormName,sFormValue
				End If
			End If
			tStream.Close
			iFormStart = iFormStart+iStart+2
			'������ļ�β�˾��˳�
		Loop Until  (iFormStart+2) >= iFormEnd 
		RequestBinData = ""
		Set tStream = Nothing
		Set KS=Nothing
	End Sub
End Class

'----------------------------------------------------------------------------------------------------
'�ļ�������
Class FileObj_Class
	Dim FormName,FileName,FilePath,FileSize,FileType,FileStart,FileExt
	'�����ļ�����
	Public Function SaveToFile (Path)
		On Error Resume Next
		Dim KS:Set KS=New PublicCls
		Dim oFileStream
		Set oFileStream = KS.InitialObject ("ADODB.Stream")
		oFileStream.Type = 1
		oFileStream.Mode = 3
		oFileStream.Open
		UpFileStream.Position = FileStart
		UpFileStream.CopyTo oFileStream,FileSize
		oFileStream.SaveToFile Path,2
		oFileStream.Close
		Set oFileStream = Nothing 
		Set KS=Nothing
	End Function
	'ȡ���ļ�����
	Public Function FileData
		UpFileStream.Position = FileStart

		FileData = UpFileStream.Read (FileSize)
	End Function
End Class

%>