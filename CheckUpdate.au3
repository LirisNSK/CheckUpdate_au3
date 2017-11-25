 ; ****************************************************************************
 ;
 ;	�������� ��������� ������������ ���� 1�:������������������� 10.3
 ;	� 2017 Liris 
 ;	mailto:liris@ngs.ru
 ;
 ;*****************************************************************************

#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#NoTrayIcon

Global	$sLogfileName	; ��� log-�����
Global	$sIBConn		; ������ ����������� � ��
Global	$IBConnectionParam	; ��������� ����������� � ���� ������
Global	$sV8exePath		; ���� � ������������ ����� 1cv8.exe
Global	$sArcLogfileName; ��� ������ log-�����
Global	$iLogMaxSize	; ������������ ������ log-�����
Global	$cPIDFileExt	; ���������� pid-�����
Global	$sComConnectorObj	; ��� COM-������� 
Global	$v8ComConnector, $connDB
Global	$g_IServerAgentConnection	; ���������� � ������� �������
Global	$g_ClusterInfo	; ������� ��������
Global	$g_InfoBaseInfo	; ���������� �� ������� �� � �������� ��������
Global	$sEmailTo, $sEmailCc ; ������ ��� �������� ���������
Global	$sHTMLBodyForEmail	; HTML-�������� ��� �������� ���������
Global	$g_ForceUpdate	; ������. ������ - ��������� �������������

; ������ �������� �� ini-����� � ������� ���������� ��������� ������
Func ReadParamsFromIni()
	
	$sIniFileName	= StringRegExpReplace(@ScriptFullPath, '^(?:.*\\)([^\\]*?)(?:\.[^.]+)?$', '\1')
	$sIniFileName	= @ScriptDir & "\" & $sIniFileName & ".ini"
	; ������ �� INI-����� �������� 'Key' � ������ 'Section'.
	$sLogfileName		= IniRead($sIniFileName,	"EXCHANGE", "LogfileName", "exchange_log.log")
	$sIBConn			= IniRead($sIniFileName,	"EXCHANGE", "IBConn", "File=D:\Retail;Usr=Admin;Pwd=Admin;")
	$sV8exePath			= IniRead($sIniFileName,	"EXCHANGE", "V8exePath", """C:\Program Files\1cv83\common\1cestart.exe""" )
	$sComConnectorObj	= IniRead($sIniFileName,	"EXCHANGE", "ComConnectorObj", "V83.COMConnector")
	$sArcLogfileName	= IniRead($sIniFileName,	"EXCHANGE", "ArcLogfileName", "exchange_DDMMYYYY_log.old")
	$cPIDFileExt		= IniRead($sIniFileName,	"EXCHANGE", "PIDFileExt", "pid")	
	$iLogMaxSize		= Int(IniRead($sIniFileName,	"EXCHANGE", "LogMaxSize", 512000))
	$sEmailTo			= IniRead($sIniFileName,	"EXCHANGE", "EmailTo", "")
	$sEmailCc			= IniRead($sIniFileName,	"EXCHANGE", "EmailCc", "")
	$g_ForceUpdate		= False

	; ���� ������� ��������� ������ ����������, ����� �� ��������� �� ��������� �����
	If StringLen($sIBConn) > 0 Then
		$IBConnectionParam = SplitConnectionString($sIBConn)
	EndIf
	; ������ ���������� ��������� ������
	; ���������, ���������� � ��������� ������ ����� ��������� ��� ����������� ini-�����
	$iParamCount	= $CmdLine[0]
	For $iCurrParam = 1 To $iParamCount
		Select
			Case StringLower($CmdLine[$iCurrParam]) = StringLower("ForceUpdate")
				$g_ForceUpdate	= True
				AddToLog("������������ ���� ForceUpdate")
		EndSelect
	Next

EndFunc

; ������� ��� ������ � ��� ������
; *****************************************************************************

Func GetTimestampString()
	Local $sDT
	$sDT	= "[" & @MDAY & "." & @MON & "." & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC & "." & @MSEC & "] ";
	Return $sDT
EndFunc

; ���������� ������� ���� � ���� 20171121
Func GetCurrentDateString()
	Local $sDT
	$sDT	= "" & @YEAR & "" & @MON & "" & @MDAY & "" ;
	Return $sDT
EndFunc

Func AddToLog($sMsg)
	
	PrepareLogFile()
	$mCurrFolder	=	@ScriptDir
	$mLogfilePath	=	$mCurrFolder & "\" & $sLogfileName
	$fFileLog		=	FileOpen($mLogfilePath, 1)
	$sWriteToFile	=	GetTimestampString() & $sMsg
	FileWriteLine($fFileLog, $sWriteToFile)
	FileClose($fFileLog)
	
EndFunc

Func PrepareLogFile()
	$mCurrFolder	=	@ScriptDir
	$mLogfilePath	=	$mCurrFolder & "\" & $sLogfileName
	If (FileExists($mLogfilePath)) Then
	
		$fLogfileSize	=	FileGetSize($mLogfilePath)
		
		If ($fLogfileSize >= $iLogMaxSize) Then
		
			;	��������� ������ ��� ���������
			$sDT	= @YEAR & "_" & @MON & "_" & @MDAY
			;	�������� ������ � ����� �����
			$sNewFileName	=	StringReplace($sArcLogfileName,"DDMMYYYY", $sDT)
			$sNewFileName	=	$mCurrFolder & "\" & $sNewFileName
			;	���������� ���� � �����
			FileMove($mLogfilePath, $sNewFileName)
			;	������� ����� ���� ����
			$fFileLog		=	FileOpen($mLogfilePath, 1)
			$sWriteToFile	=	GetTimestampString() & "���������� ��� ��������� � �����: " & $sNewFileName
			FileWriteLine($fFileLog, $sWriteToFile)
			FileClose($fFileLog)
		EndIf

	Else
		$fFileLog		=	FileOpen($mLogfilePath, 1)
		$sWriteToFile	=	GetTimestampString() & "����������� ����� ���-���� "
		FileWriteLine($fFileLog, $sWriteToFile)
		FileClose($fFileLog)
	EndIf

EndFunc

;
; ������� ��� ������ � ����������
; *****************************************************************************

; ���������� PID ����������� ��������. ���� � ������� ������ �������� �������� �� ��������, ���������� 0
Func GetLastPID()
	
	$mReturn		=	0;
	$mCurrFolder	=	@ScriptDir;
	$hSearch		=	FileFindFirstFile($mCurrFolder & "\*." & $cPIDFileExt)
	; ��������, �������� �� ����� ��������
	If $hSearch = -1 Then
		; ������ ��� ������ ������
		; ����� ������� 0
		Return 0
		Exit
	EndIf	
	
	While 1
		$sFile = FileFindNextFile($hSearch) ; ���������� ��� ���������� �����, ������� �� ������� �� ����������
		If @error Then ExitLoop
		
		AddToLog("������ PID-����: " & $sFile)
		$sPID	=	StringReplace($sFile, "." & $cPIDFileExt, "")
		;AddToLog("������������� ��������: " & $sPID)
		$mReturn=	Int($sPID)
	WEnd
	
	FileClose($hSearch)
	
	return $mReturn
	
EndFunc

; ������� ��� PID-����� � ����� �������
Func DeletePIDFile()
	
	AddToLog("������� ������� PID-����� � �����: " & @ScriptDir)
	$mCurrFolder	=	@ScriptDir;
	$iResult		=	FileDelete($mCurrFolder & "\*." & $cPIDFileExt)
	If $iResult > 0 Then
		AddToLog("PID-����� ������� �������")
	Else
		AddToLog("��� �������� ������ ��������� ������, ���� ��� ������ ��� ��������")
	EndIf
	
EndFunc

; ������� ����� PID-����
Func CreatePIDFile()
	
	$mCurrFolder	=	@ScriptDir;
	$iCurrentPID	=	@AutoItPID;
	$sPIDFileName	=	$mCurrFolder & "\" & $iCurrentPID & "." & $cPIDFileExt;
	
	$fFileOut		=	FileOpen($sPIDFileName, 1)
	FileWrite($fFileOut, String($iCurrentPID))
	FileClose($fFileOut)
	
	$sMsg = "������ ����� PID-����: " & String($iCurrentPID)
	AddToLog($sMsg)

EndFunc	

; ��������� ������ ����������� �� ������������
Func SplitConnectionString($sIBConn)
	
	; 0: ��� ���� 0 - ��������, 1 - ������-������
	; 1: ��� ������������
	; 2: ������
	; 3: ���� � ���� ������/��� ���� ������
	; 4: ��� �������
	
	Local $aResult[5]
	$aIBParams = StringSplit($sIBConn, ";")
	$iParamCount	= $aIBParams[0]

	For $iCurrParam = 1 To $iParamCount

		$iLength	= StringLen($aIBParams[$iCurrParam])
		If $iLength = 0 Then ContinueLoop

		$sParamName	= StringLeft($aIBParams[$iCurrParam], StringInStr($aIBParams[$iCurrParam], "=") -1)
		
		Select
			Case StringLower($sParamName) = StringLower("File")
				$aResult[3]	= StringReplace($aIBParams[$iCurrParam], $sParamName & "=", "")
				$aResult[0]	= 0
			Case StringLower($sParamName) = StringLower("Srvr")
				$aResult[4]	= StringReplace($aIBParams[$iCurrParam], $sParamName & "=", "")
				$aResult[0]	= 1
			Case StringLower($sParamName) = StringLower("Ref")
				$aResult[3]	= StringReplace($aIBParams[$iCurrParam], $sParamName & "=", "")
			Case StringLower($sParamName) = StringLower("Usr")
				$aResult[1]	= StringReplace($aIBParams[$iCurrParam], $sParamName & "=", "")
			Case StringLower($sParamName) = StringLower("Pwd")
				$aResult[2]= StringReplace($aIBParams[$iCurrParam], $sParamName & "=", "")
		EndSelect

	Next
	
	Return $aResult
	
EndFunc

;
; ������� ��� ������ � ���������
; *****************************************************************************

; ������� ServerAgentConnection
Func ConnectToServerAgent()

	If IsObj($g_IServerAgentConnection) = 1 Then
		;AddToLog("���������� � ������� ������� ����������� �����")
		Return True
	Else
		
		if Not CreateCOMConnector() Then
			Return False
		EndIf
		
		AddToLog("����������� � ������ ������� " & $IBConnectionParam[4])
		$g_IServerAgentConnection = $v8ComConnector.ConnectAgent($IBConnectionParam[4])

		If IsObj($g_IServerAgentConnection) = 0 Then
			AddToLog("������ ��� ����������� � ������ �������" )
			Return False
		Else
			AddToLog("����������� � ������ ������� �����������")
			Return True
		EndIf
	EndIf
	
EndFunc

; ��������� ���������� � ������� �������
Func DisconnectFromServerAgent()

	AddToLog("����������� ���������� � ������� �������")
	While IsObj($g_IServerAgentConnection)
		$g_IServerAgentConnection	= 0
		Sleep(1000)
		AddToLog("�������� �������� ���������� � ������� �������")
	WEnd
	
	Return True

EndFunc

Func FindIBInCluster($pClsr)
	
	$lInfoBases = $g_IServerAgentConnection.GetInfoBases($pClsr)
	$pIB	= ""
	For $vIB In $lInfoBases
		AddToLog("�������������� �� " & $vIB.Name)
		If $vIB.Name = $IBConnectionParam[3] Then
			$pIB	= $vIB
			AddToLog("� �������� " & $pClsr.ClusterName & " ������� ���� ������ " & $IBConnectionParam[3])
			ExitLoop
		EndIf
	Next
	If IsObj($pIB) = 0 Then
		AddToLog("� �������� " & $pClsr.ClusterName & " ���� ������ " & $IBConnectionParam[3] & " �� �������")
		Return $pIB
	Else
		Return $pIB
	EndIf

EndFunc

; �������� ������� ������ ��
Func CheckIBSessionsAndTryTerminate()

	Local $lClusters
	Local $aWorkProcs, $lWorkProc
	Local $lIBSessions, $vIBSession
	
	If Not ConnectToServerAgent() Then
		Return False
	EndIf

	AddToLog("��������� ������ ���������")
	$lClusters	= $g_IServerAgentConnection.GetClusters()
	For $vCurrentClr In $lClusters
		
		$g_IServerAgentConnection.Authenticate($vCurrentClr, "", "")
		
		AddToLog("��������� ������� " & $vCurrentClr.ClusterName)

		$mIB	= FindIBInCluster($vCurrentClr)
		If IsObj($mIB) = 0 Then
			Return False
		EndIf

		$lIBSessions	= $g_IServerAgentConnection.GetInfoBaseSessions($vCurrentClr, $mIB)
		AddToLog("���������� ��������� �������� ������� ��")
		For $vIBSession In $lIBSessions
			
			$lSessState	= ($vIBSession.Hibernate) ? " (������)" : " (��������)"
			If $vIBSession.AppID = "Designer" Then
				AddToLog("��������� �������� ����� �������������. ����������� ������ ������� ����������")
				AddHTMLBodyForEmail("��������� �������� ����� �������������. ���������� ��������� ���������")
				Return False
				ExitLoop
			ElseIf $vIBSession.AppID = "BackgroundJob" Then
				; ��������� �������� �������
				AddToLog("����� �������� ������� " & $vIBSession.SessionID & " ��-��: " & $vIBSession.AppID & " �����-��: " & $vIBSession.UserName & $lSessState)
			ElseIf $vIBSession.AppID = "COMConnection" Then
				; ��������� COMConnection 
				AddToLog("����� COMConnection " & $vIBSession.SessionID & " �����-��: " & $vIBSession.UserName & $lSessState)
			ElseIf ($vIBSession.AppID = "1CV8") Or ($vIBSession.AppID = "1CV8C") Then
				; ��������� ����������������� ������
				AddToLog("����������� ����� " & $vIBSession.SessionID & " ��-��: " & $vIBSession.AppID & " �����-��: " & $vIBSession.UserName & $lSessState)
				AddHTMLBodyForEmail("����������� ����� " & $vIBSession.SessionID & " ��-��: " & $vIBSession.AppID & " �����-��: " & $vIBSession.UserName & $lSessState)
				$g_IServerAgentConnection.TerminateSession($vCurrentClr, $vIBSession)
			EndIf
		Next
	Next
	
	DisconnectFromServerAgent()
	Return True

EndFunc

;
; ������� ��� ������ � ����� ������
; *****************************************************************************

; ������� COMConnector
Func CreateCOMConnector()

	If IsObj($v8ComConnector) = 1 Then
		Return True
	Else
		AddToLog("��������� ����� ������ COMConnector")
		
		$v8ComConnector = ObjCreate($sComConnectorObj)

		If IsObj($v8ComConnector) = 0 Then
			AddToLog("������ ��� �������� COM-������� " & $sComConnectorObj)
			Return False
		Else
			Return True
		EndIf
	EndIf

EndFunc

; ���������� COMConnector
Func DestroyCOMConnector()

	AddToLog("������������ ������ COMConnector")
	While IsObj($v8ComConnector)
		$v8ComConnector	= 0
		Sleep(1000)
		AddToLog("�������� ������������ ������ �� ������� COMConnector")
	WEnd

	Return True

EndFunc

; ������������ ����������� � ���� ������
Func ConnectToDataBase()
	
	Local $mv8exe

	AddToLog("������� ���������� ����������� � ����� ������")
	
	If IsObj($connDB) = 1 Then
		Return True
	Else

		If Not CreateCOMConnector() Then
			Return False
		EndIf

		AddToLog("����������� � ��")
		$connDB	=	$v8ComConnector.Connect($sIBConn)

		If IsObj($connDB) = 0 Then
			AddToLog("��� ����������� � �� ��������� ������")
			Return False
		Else
			AddToLog("����������� � �� �����������")
			; �������� ���� � ������������ ����� ������� ������ ���������
			$mv8exe	= $connDB.BinDir() & "1cv8.exe"
			AddToLog("���� � v8exe: " & $mv8exe)

			Return True
		EndIf
	EndIf

EndFunc

; ��������� ���������� � ����� ������ � ���������� COM-������
Func DisconnectFromDatabase()

	AddToLog("����������� ���������� � ����� ������")
	While IsObj($connDB)
		$connDB	= 0
		Sleep(1000)
		AddToLog("�������� �������� ���������� � ����� ������")
	WEnd
	
	Return True

EndFunc

; ������� ��������� ������ ��� ������� ������������� � ��������� ���
Func RunDesignerForUpdate()

	;v8exe & " DESIGNER /F" & $sIBPath  & " /N" & IBAdminName & " /P" & IBAdminPwd & " /WA- /UpdateDBCfg /Out" & $ServiceFileName & " -NoTruncate /DisableStartupMessages"
	Local $sUpdCmdLine, $sRunClientCmdLine
	Local $sIBPath, $sIBAdmin, $sIBAdminPwd, $ServiceFileName
	Local $ProcResult

	$sServiceFileName	=	@ScriptDir & "\" & GetCurrentDateString() & "_upd_result.txt" 
	
	; ������� ��������� ����������� �� ������ �����������
	$sIBAdmin	=	$IBConnectionParam[1]
	$sIBAdminPwd=	$IBConnectionParam[2]
	$sIBPath	=	$IBConnectionParam[3]
	$sIBSrvr	=	$IBConnectionParam[4]
	
	; ������������ ������ ��� ������� �������������
	If $IBConnectionParam[0] = 0 Then
		$sUpdCmdLine	=	$sV8exePath & " DESIGNER /F" & $sIBPath  & " /N" & $sIBAdminPwd & " /P" & $sIBAdminPwd 
	Else
		$sUpdCmdLine	=	$sV8exePath & " DESIGNER /S" & $sIBSrvr & "\" & $sIBPath  & " /N" & $sIBAdminPwd & " /P" & $sIBAdminPwd 
	EndIf
	
	$sUpdCmdLine	=	$sUpdCmdLine & " /WA- /UpdateDBCfg /Out""" & $sServiceFileName & """ -NoTruncate /DisableStartupMessages"
	
	; �������� ���������
	AddToLog("������ ������������� ��� �������� ���������")
	
	$PIDUpdCfg	= Run($sUpdCmdLine)
	
	If $PIDUpdCfg = 0 Then
		AddToLog("������ ��� ���������� �������")
		AddToLog("��������� ������: " & $sUpdCmdLine)
		Return False
	Else
		AddToLog("������� ������� " & $PIDUpdCfg)
	EndIf

	$ProcResult = ProcessWaitClose($PIDUpdCfg)
	
	; ���������� 1 � ������������� �������� @extended ������ ���� ������ ��������
	If $ProcResult = 1 Then
		$iCfgErrCode	= @extended
		AddToLog("������������ ������� �������� ������ � ����� = " & $iCfgErrCode)
		; ���� ������������ ������ ��������� = 1, ������ ������� �� ���������
		If $iCfgErrCode = 1 Then
			AddToLog("��� " & $iCfgErrCode & " �������� ������ �������� ���������")
			Return False
		Else
			Return True
		EndIf
		
	Else
		AddToLog("������� ����������� �������.")
		Return False
	EndIf
		
EndFunc

; ������� �������� ��������� ��������� �����������
Func RunDynamicUpdate()
	
	if Not RunDesignerForUpdate() Then
		; ���� �� ������� ���������, �������� ��������� � ������������� ��������� ���������
		; ����������� ��������� �� ������, ������� ����� ���������� �� Email
		AddHTMLBodyForEmail("�� ������� ��������� ��������� �����������.")
		AddHTMLBodyForEmail("���� ������ " & $IBConnectionParam[3] & " (" & $IBConnectionParam[4] & ")" )

		If Not PrepareAndSendEmail($sHTMLBodyForEmail) Then
			AddToLog("���������� ������������ ��������� ����������� �������")
		EndIf
		
		Return False

	Else
		; ������� ��������� ���������. ������� �� ���� ������
		; ����������� ��������� �� �������� �������� ���������, ������� ����� ���������� �� Email
		AddHTMLBodyForEmail("��������� ������������ ���������� ���� ������ " & $IBConnectionParam[3] & " (" & $IBConnectionParam[4] & ")." )
	
		If Not PrepareAndSendEmail($sHTMLBodyForEmail) Then
			AddToLog("���������� ������������ ��������� ����������� �������")
		EndIf

		Return True

	EndIf
	
EndFunc

; ������� �������� ��������� ��������� � �����������������
Func RunNonDynamicUpdate()
	
	; ���������� ��������� ���������

	if Not RunDesignerForUpdate() Then
		; ���� �� ������� ���������, �������� ��������� � ������������� ��������� ���������
		AddToLog("�� ������� ��������� ��������� �����������.")
		AddToLog("��� �������� ��������� ��������� ���������������� ��� ����������� ������")
		If $g_ForceUpdate Then
			AddToLog("����������� ������� ��������� ������")
			If CheckIBSessionsAndTryTerminate() Then
				Return True
			Else
				AddHTMLBodyForEmail("�� ������� ��������� ��������� �����������.")
				AddHTMLBodyForEmail("���� ������ " & $IBConnectionParam[3] & " (" & $IBConnectionParam[4] & ")" )
				AddHTMLBodyForEmail("��� �������� ��������� ��������� ����������� ������ � ����")
			EndIf
		Else
			AddHTMLBodyForEmail("�� ������� ��������� ��������� �����������.")
			AddHTMLBodyForEmail("���� ������ " & $IBConnectionParam[3] & " (" & $IBConnectionParam[4] & ")" )
			AddHTMLBodyForEmail("��� �������� ��������� ��������� ����������� ������ � ����")
		EndIf
		
		If Not PrepareAndSendEmail($sHTMLBodyForEmail) Then
			AddToLog("���������� ������������ ��������� ����������� �������")
		EndIf
		
		Return False

	Else
		; ������� ��������� ���������. ������� �� ���� ������
		; ����������� ��������� �� �������� �������� ���������, ������� ����� ���������� �� Email
		AddHTMLBodyForEmail("��������� ���������� ���� ������ " & $IBConnectionParam[3] & " (" & $IBConnectionParam[4] & ")." )
		AddHTMLBodyForEmail("�������� ��������� �������")
		If Not PrepareAndSendEmail($sHTMLBodyForEmail) Then
			AddToLog("���������� ������������ ��������� ����������� �������")
		EndIf

		Return True

	EndIf
EndFunc

; ������� ��������� ������������� ���������� ���������
Func CheckUpdate()
	
	Local $TryCount
	
	; ����������� � ���� ������
	ConnectToDataBase()

	; ����������� ������������� ���������� ������������ �� ��������� ������
	AddToLog("�������� ������������� ������� ���������")
	$TryCount = 0 

	While ($connDB.ConfigurationChanged() = True) AND $TryCount <= 3 
		
		$TryCount = $TryCount + 1

		if $connDB.DataBaseConfigurationChangedDynamically() Then
			AddToLog("� ������� ����������� ������������ �� ���� �������� �����������")
			AddToLog("��������� ��������������� � ���� ������ (���������� ������)")
			DisconnectFromDatabase()
		Else
			AddToLog("������������ �� ��������, ��������� ���������� ���������")
			AddToLog("������� ���������� ���������")
			DisconnectFromDatabase()
			
			If Not RunNonDynamicUpdate() Then
				ExitLoop
			EndIf
		EndIf

		if not ConnectToDataBase() Then
			AddToLog("����� �� ����� ��-�� ������ ����������� � ���� ������")
			ExitLoop
		EndIf

	WEnd

	DisconnectFromDatabase()
	DestroyCOMConnector()

EndFunc

; ���������� ����, ���� ������ ���������� ���������� �������. ����������� ������ ������� ���������
Func CanIContinue()
	$mResult	=	False
	$iLastPID	=	GetLastPID()
	
	If ($iLastPID > 0) Then
		If (ProcessExists($iLastPID)) Then
			AddToLog("������� " & String($iLastPID) & " ��������")
			$mResult	= False
		Else
			DeletePIDFile()
			$mResult	= True
		EndIf
	Else
		$mResult = True
	EndIf
	
	Return $mResult
	
EndFunc

;
; ������� ��� ������ � ����������� ������
; *****************************************************************************

; EmailParam - ������
; 0 - ���� (������)
; 1 - ����� (������)
; 2 - ���� (������)
; 3 - ���� ��������� (������)
Func SendEmail($EmailParam)

	Local $EmailMsg

	AddToLog("���������� �������� ��� ������������ ������������ ���������")
	$EmailMsg	= ObjCreate("CDO.Message")

	If @error Then
		AddToLog("������ ��� �������� CDOMessage ")
		Return False
	EndIf
	
	$EmailMsg.To		= $EmailParam[0]
	$EmailMsg.Cc		= $EmailParam[1]
	$EmailMsg.From		= "notifier@pelican.local"
	$EmailMsg.Subject	= $EmailParam[2]
	$EmailMsg.BodyPart.Charset = "Windows-1251"

	$EmailMsg.HTMLBody	= $EmailParam[3]
	$EmailMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/usemessageresponsetext").value		= true;
	$EmailMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/languagecode").value 				= 1049;
	$EmailMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing").value 					= 2;
	$EmailMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate").value 			= 1;
	$EmailMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver").value 					= "mx.pelican.local";
	$EmailMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout").value 		= "10";
	$EmailMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername").value 				= "notifier";
	$EmailMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword").value 				= "notifier";
	$EmailMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport").value 				= 251;
	$EmailMsg.Configuration.Fields.Update();
	
	$EmailMsg.Send();

	If @error Then
		AddToLog("������ ��� �������� ������������ ���������")
		Return False
	Else
		Return True
	EndIf

EndFunc

; ������� ���� ��������� � ���������� ���
; $EmailMsgTxt - ����� ��������� 
Func PrepareAndSendEmail($EmailMsgTxt)
	Local $mEmlParam[4]

	If $sEmailTo = "" Then
		Return False
	EndIf

	$mEmlParam[0] = $sEmailTo
	$mEmlParam[1] = $sEmailCc
	$mEmlParam[2] = "�������������� ����������� �� �������� ��� ����� ������"
	$mEmlParam[3] = $EmailMsgTxt

	SendEmail($mEmlParam)

	Return True
EndFunc

Func AddHTMLBodyForEmail($PlainText)
	Local $mHTMLbr	

	$mHTMLbr	= "<br>"
	If StringLen($sHTMLBodyForEmail) = 0 Then
		$sHTMLBodyForEmail	= $PlainText & $mHTMLbr
	Else
		$sHTMLBodyForEmail	= $sHTMLBodyForEmail & $PlainText & $mHTMLbr
	EndIf
	
EndFunc

; �������� ���������
; *****************************************************************************

; ������ ��������� � ���������� ����������
ReadParamsFromIni()

AddToLog("=> ����� ���������")

If ( CanIContinue() ) Then
	
	CreatePIDFile()	
	CheckUpdate()
	DeletePIDFile()	
	AddToLog("<= ��������� ���������");
	
else

	AddToLog("������ �� ����� ���������� ������. ������� CanIContinue �� �����������");
	AddToLog("<= ������ ��������� ������");
	
EndIf