 ; ****************************************************************************
 ;
 ;	Проверка изменения конфигурации узла 1С:УправлениеТорговлей 10.3
 ;	© 2017 Liris 
 ;	mailto:liris@ngs.ru
 ;
 ;*****************************************************************************

#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#NoTrayIcon

Global	$sLogfileName	; Имя log-файла
Global	$sIBConn		; Строка подключения к ИБ
Global	$IBConnectionParam	; Параметры подключения к базе данных
Global	$sV8exePath		; Путь к исполняемому файлу 1cv8.exe
Global	$sArcLogfileName; Имя архива log-файла
Global	$iLogMaxSize	; Максимальный размер log-файла
Global	$cPIDFileExt	; Расширение pid-файла
Global	$sComConnectorObj	; Имя COM-объекта 
Global	$v8ComConnector, $connDB
Global	$g_IServerAgentConnection	; Соединение с агентом сервера
Global	$g_ClusterInfo	; Кластер серверов
Global	$g_InfoBaseInfo	; Информация об объекте ИБ в кластере серверов
Global	$sEmailTo, $sEmailCc ; Адреса для доставки сообщений
Global	$sHTMLBodyForEmail	; HTML-документ для будущего сообщения
Global	$g_ForceUpdate	; Булево. Истина - отключить пользователей

; Чтение настроек из ini-файла и разбора параметров командной строки
Func ReadParamsFromIni()
	
	$sIniFileName	= StringRegExpReplace(@ScriptFullPath, '^(?:.*\\)([^\\]*?)(?:\.[^.]+)?$', '\1')
	$sIniFileName	= @ScriptDir & "\" & $sIniFileName & ".ini"
	; Читает из INI-файла параметр 'Key' в секции 'Section'.
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

	; Если успешно прочитали строку соединения, можно ее разобрать на составные части
	If StringLen($sIBConn) > 0 Then
		$IBConnectionParam = SplitConnectionString($sIBConn)
	EndIf
	; Чтение параметров командной строки
	; Параметры, переданные в командной строке имеют приоритет над параметрами ini-файла
	$iParamCount	= $CmdLine[0]
	For $iCurrParam = 1 To $iParamCount
		Select
			Case StringLower($CmdLine[$iCurrParam]) = StringLower("ForceUpdate")
				$g_ForceUpdate	= True
				AddToLog("Задействован ключ ForceUpdate")
		EndSelect
	Next

EndFunc

; Функции для работы с лог файлом
; *****************************************************************************

Func GetTimestampString()
	Local $sDT
	$sDT	= "[" & @MDAY & "." & @MON & "." & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC & "." & @MSEC & "] ";
	Return $sDT
EndFunc

; Возвращает текущую дату в виде 20171121
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
		
			;	Формируем строку для замещения
			$sDT	= @YEAR & "_" & @MON & "_" & @MDAY
			;	Замещаем строку в имени файла
			$sNewFileName	=	StringReplace($sArcLogfileName,"DDMMYYYY", $sDT)
			$sNewFileName	=	$mCurrFolder & "\" & $sNewFileName
			;	Перемещаем файл в архив
			FileMove($mLogfilePath, $sNewFileName)
			;	Создаем новый файл лога
			$fFileLog		=	FileOpen($mLogfilePath, 1)
			$sWriteToFile	=	GetTimestampString() & "Предыдущий лог перемещен в архив: " & $sNewFileName
			FileWriteLine($fFileLog, $sWriteToFile)
			FileClose($fFileLog)
		EndIf

	Else
		$fFileLog		=	FileOpen($mLogfilePath, 1)
		$sWriteToFile	=	GetTimestampString() & "Сформирован новый лог-файл "
		FileWriteLine($fFileLog, $sWriteToFile)
		FileClose($fFileLog)
	EndIf

EndFunc

;
; Функции для работы с процессами
; *****************************************************************************

; Возвращает PID работающего процесса. Если в текущий момент никакого процесса не запущено, возвращает 0
Func GetLastPID()
	
	$mReturn		=	0;
	$mCurrFolder	=	@ScriptDir;
	$hSearch		=	FileFindFirstFile($mCurrFolder & "\*." & $cPIDFileExt)
	; Проверка, является ли поиск успешным
	If $hSearch = -1 Then
		; Ошибка при поиске файлов
		; Нужно вернуть 0
		Return 0
		Exit
	EndIf	
	
	While 1
		$sFile = FileFindNextFile($hSearch) ; возвращает имя следующего файла, начиная от первого до последнего
		If @error Then ExitLoop
		
		AddToLog("Найден PID-файл: " & $sFile)
		$sPID	=	StringReplace($sFile, "." & $cPIDFileExt, "")
		;AddToLog("Идентификатор процесса: " & $sPID)
		$mReturn=	Int($sPID)
	WEnd
	
	FileClose($hSearch)
	
	return $mReturn
	
EndFunc

; Удаляет все PID-файлы в папке скрипта
Func DeletePIDFile()
	
	AddToLog("Попытка удалить PID-файлы в папке: " & @ScriptDir)
	$mCurrFolder	=	@ScriptDir;
	$iResult		=	FileDelete($mCurrFolder & "\*." & $cPIDFileExt)
	If $iResult > 0 Then
		AddToLog("PID-файлы удалены успешно")
	Else
		AddToLog("При удалении файлов произошла ошибка, либо нет файлов для удаления")
	EndIf
	
EndFunc

; Создает новый PID-файл
Func CreatePIDFile()
	
	$mCurrFolder	=	@ScriptDir;
	$iCurrentPID	=	@AutoItPID;
	$sPIDFileName	=	$mCurrFolder & "\" & $iCurrentPID & "." & $cPIDFileExt;
	
	$fFileOut		=	FileOpen($sPIDFileName, 1)
	FileWrite($fFileOut, String($iCurrentPID))
	FileClose($fFileOut)
	
	$sMsg = "Создан новый PID-файл: " & String($iCurrentPID)
	AddToLog($sMsg)

EndFunc	

; Разбирает строку подключения на составляющие
Func SplitConnectionString($sIBConn)
	
	; 0: Тип базы 0 - Файловая, 1 - Клиент-сервер
	; 1: Имя пользователя
	; 2: Пароль
	; 3: Путь к базе данных/Имя базы данных
	; 4: Имя сервера
	
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
; Функции для работы с кластером
; *****************************************************************************

; Создает ServerAgentConnection
Func ConnectToServerAgent()

	If IsObj($g_IServerAgentConnection) = 1 Then
		;AddToLog("Соединение с агентом сервера установлено ранее")
		Return True
	Else
		
		if Not CreateCOMConnector() Then
			Return False
		EndIf
		
		AddToLog("Подключение к агенту сервера " & $IBConnectionParam[4])
		$g_IServerAgentConnection = $v8ComConnector.ConnectAgent($IBConnectionParam[4])

		If IsObj($g_IServerAgentConnection) = 0 Then
			AddToLog("Ошибка при подключении к агенту сервера" )
			Return False
		Else
			AddToLog("Подключение к агенту сервера установлено")
			Return True
		EndIf
	EndIf
	
EndFunc

; Закрывает соединение с агентом сервера
Func DisconnectFromServerAgent()

	AddToLog("Закрывается соединение с агентом сервера")
	While IsObj($g_IServerAgentConnection)
		$g_IServerAgentConnection	= 0
		Sleep(1000)
		AddToLog("Ожидание закрытия соединения с агентом сервера")
	WEnd
	
	Return True

EndFunc

Func FindIBInCluster($pClsr)
	
	$lInfoBases = $g_IServerAgentConnection.GetInfoBases($pClsr)
	$pIB	= ""
	For $vIB In $lInfoBases
		AddToLog("Обрабатывается ИБ " & $vIB.Name)
		If $vIB.Name = $IBConnectionParam[3] Then
			$pIB	= $vIB
			AddToLog("В кластере " & $pClsr.ClusterName & " найдена база данных " & $IBConnectionParam[3])
			ExitLoop
		EndIf
	Next
	If IsObj($pIB) = 0 Then
		AddToLog("В кластере " & $pClsr.ClusterName & " база данных " & $IBConnectionParam[3] & " не найдена")
		Return $pIB
	Else
		Return $pIB
	EndIf

EndFunc

; Пытается закрыть сеансы ИБ
Func CheckIBSessionsAndTryTerminate()

	Local $lClusters
	Local $aWorkProcs, $lWorkProc
	Local $lIBSessions, $vIBSession
	
	If Not ConnectToServerAgent() Then
		Return False
	EndIf

	AddToLog("Получение списка кластеров")
	$lClusters	= $g_IServerAgentConnection.GetClusters()
	For $vCurrentClr In $lClusters
		
		$g_IServerAgentConnection.Authenticate($vCurrentClr, "", "")
		
		AddToLog("Обнаружен кластер " & $vCurrentClr.ClusterName)

		$mIB	= FindIBInCluster($vCurrentClr)
		If IsObj($mIB) = 0 Then
			Return False
		EndIf

		$lIBSessions	= $g_IServerAgentConnection.GetInfoBaseSessions($vCurrentClr, $mIB)
		AddToLog("Начинается обработка активных сеансов ИБ")
		For $vIBSession In $lIBSessions
			
			$lSessState	= ($vIBSession.Hibernate) ? " (Спящий)" : " (Активный)"
			If $vIBSession.AppID = "Designer" Then
				AddToLog("Обнаружен активный сеанс Конфигуратора. Продолжение работы скрипта невозможно")
				AddHTMLBodyForEmail("Обнаружен активный сеанс Конфигуратора. Невозможно применить изменения")
				Return False
				ExitLoop
			ElseIf $vIBSession.AppID = "BackgroundJob" Then
				; Обработка Фонового задания
				AddToLog("Сеанс фонового задания " & $vIBSession.SessionID & " пр-ие: " & $vIBSession.AppID & " польз-ль: " & $vIBSession.UserName & $lSessState)
			ElseIf $vIBSession.AppID = "COMConnection" Then
				; Обработка COMConnection 
				AddToLog("Сеанс COMConnection " & $vIBSession.SessionID & " польз-ль: " & $vIBSession.UserName & $lSessState)
			ElseIf ($vIBSession.AppID = "1CV8") Or ($vIBSession.AppID = "1CV8C") Then
				; Обработка пользовательского сеанса
				AddToLog("Закрывается сеанс " & $vIBSession.SessionID & " пр-ие: " & $vIBSession.AppID & " польз-ль: " & $vIBSession.UserName & $lSessState)
				AddHTMLBodyForEmail("Закрывается сеанс " & $vIBSession.SessionID & " пр-ие: " & $vIBSession.AppID & " польз-ль: " & $vIBSession.UserName & $lSessState)
				$g_IServerAgentConnection.TerminateSession($vCurrentClr, $vIBSession)
			EndIf
		Next
	Next
	
	DisconnectFromServerAgent()
	Return True

EndFunc

;
; Функции для работы с базой данных
; *****************************************************************************

; Создает COMConnector
Func CreateCOMConnector()

	If IsObj($v8ComConnector) = 1 Then
		Return True
	Else
		AddToLog("Создается новый объект COMConnector")
		
		$v8ComConnector = ObjCreate($sComConnectorObj)

		If IsObj($v8ComConnector) = 0 Then
			AddToLog("Ошибка при создании COM-объекта " & $sComConnectorObj)
			Return False
		Else
			Return True
		EndIf
	EndIf

EndFunc

; Уничтожает COMConnector
Func DestroyCOMConnector()

	AddToLog("Уничтожается объект COMConnector")
	While IsObj($v8ComConnector)
		$v8ComConnector	= 0
		Sleep(1000)
		AddToLog("Ожидание освобождения памяти от объекта COMConnector")
	WEnd

	Return True

EndFunc

; Устанавлиает подключение к базе данных
Func ConnectToDataBase()
	
	Local $mv8exe

	AddToLog("Попытка установить подключение с базой данных")
	
	If IsObj($connDB) = 1 Then
		Return True
	Else

		If Not CreateCOMConnector() Then
			Return False
		EndIf

		AddToLog("Подключение к ИБ")
		$connDB	=	$v8ComConnector.Connect($sIBConn)

		If IsObj($connDB) = 0 Then
			AddToLog("При подключении к ИБ произошла ошибка")
			Return False
		Else
			AddToLog("Подключение к ИБ установлено")
			; Выяснить путь к исполняемому файлу текущей версии Платформы
			$mv8exe	= $connDB.BinDir() & "1cv8.exe"
			AddToLog("Путь к v8exe: " & $mv8exe)

			Return True
		EndIf
	EndIf

EndFunc

; Закрывает соединение с базой данных и уничтожает COM-объект
Func DisconnectFromDatabase()

	AddToLog("Закрывается соединение с базой данных")
	While IsObj($connDB)
		$connDB	= 0
		Sleep(1000)
		AddToLog("Ожидание закрытия соединения с базой данных")
	WEnd
	
	Return True

EndFunc

; Функция формирует строку для запуска Конфигуратора и запускает его
Func RunDesignerForUpdate()

	;v8exe & " DESIGNER /F" & $sIBPath  & " /N" & IBAdminName & " /P" & IBAdminPwd & " /WA- /UpdateDBCfg /Out" & $ServiceFileName & " -NoTruncate /DisableStartupMessages"
	Local $sUpdCmdLine, $sRunClientCmdLine
	Local $sIBPath, $sIBAdmin, $sIBAdminPwd, $ServiceFileName
	Local $ProcResult

	$sServiceFileName	=	@ScriptDir & "\" & GetCurrentDateString() & "_upd_result.txt" 
	
	; Получаю параметры подключения из строки подключения
	$sIBAdmin	=	$IBConnectionParam[1]
	$sIBAdminPwd=	$IBConnectionParam[2]
	$sIBPath	=	$IBConnectionParam[3]
	$sIBSrvr	=	$IBConnectionParam[4]
	
	; Формирование строки для запуска Конфигуратора
	If $IBConnectionParam[0] = 0 Then
		$sUpdCmdLine	=	$sV8exePath & " DESIGNER /F" & $sIBPath  & " /N" & $sIBAdminPwd & " /P" & $sIBAdminPwd 
	Else
		$sUpdCmdLine	=	$sV8exePath & " DESIGNER /S" & $sIBSrvr & "\" & $sIBPath  & " /N" & $sIBAdminPwd & " /P" & $sIBAdminPwd 
	EndIf
	
	$sUpdCmdLine	=	$sUpdCmdLine & " /WA- /UpdateDBCfg /Out""" & $sServiceFileName & """ -NoTruncate /DisableStartupMessages"
	
	; Принятие изменений
	AddToLog("Запуск Конфигуратора для принятия изменений")
	
	$PIDUpdCfg	= Run($sUpdCmdLine)
	
	If $PIDUpdCfg = 0 Then
		AddToLog("Ошибка при выполнении команды")
		AddToLog("Командная строка: " & $sUpdCmdLine)
		Return False
	Else
		AddToLog("Запущен процесс " & $PIDUpdCfg)
	EndIf

	$ProcResult = ProcessWaitClose($PIDUpdCfg)
	
	; Возвращает 1 и устанавливает значение @extended равным коду выхода процесса
	If $ProcResult = 1 Then
		$iCfgErrCode	= @extended
		AddToLog("Конфигуратор успешно завершил работу с кодом = " & $iCfgErrCode)
		; Если Конфигуратор вернул результат = 1, значит команда не выполнена
		If $iCfgErrCode = 1 Then
			AddToLog("Код " & $iCfgErrCode & " означает ошибку принятия изменений")
			Return False
		Else
			Return True
		EndIf
		
	Else
		AddToLog("Команда завершилась ошибкой.")
		Return False
	EndIf
		
EndFunc

; Функция пытается применить изменения динамически
Func RunDynamicUpdate()
	
	if Not RunDesignerForUpdate() Then
		; Если не удалось применить, Написать сообщение о необходимости применить изменения
		; Формируется сообщение об ошибке, которое будет отправлено на Email
		AddHTMLBodyForEmail("Не удалось применить изменения динамически.")
		AddHTMLBodyForEmail("База данных " & $IBConnectionParam[3] & " (" & $IBConnectionParam[4] & ")" )

		If Not PrepareAndSendEmail($sHTMLBodyForEmail) Then
			AddToLog("Подготовка электронного сообщения завершилась ошибкой")
		EndIf
		
		Return False

	Else
		; Успешно применили изменения. Напишем об этом письмо
		; Формируется сообщение об успешном принятии изменений, которое будет отправлено на Email
		AddHTMLBodyForEmail("Выполнено динамическое обновление базы данных " & $IBConnectionParam[3] & " (" & $IBConnectionParam[4] & ")." )
	
		If Not PrepareAndSendEmail($sHTMLBodyForEmail) Then
			AddToLog("Подготовка электронного сообщения завершилась ошибкой")
		EndIf

		Return True

	EndIf
	
EndFunc

; Функция пытается применить изменения с реструктуризацией
Func RunNonDynamicUpdate()
	
	; Попытаться применить изменения

	if Not RunDesignerForUpdate() Then
		; Если не удалось применить, Написать сообщение о необходимости применить изменения
		AddToLog("Не удалось применить изменения динамически.")
		AddToLog("Для принятия изменений требуется реструктуризация или монопольный доступ")
		If $g_ForceUpdate Then
			AddToLog("Выполняется попытка отключить сеансы")
			If CheckIBSessionsAndTryTerminate() Then
				Return True
			Else
				AddHTMLBodyForEmail("Не удалось применить изменения динамически.")
				AddHTMLBodyForEmail("База данных " & $IBConnectionParam[3] & " (" & $IBConnectionParam[4] & ")" )
				AddHTMLBodyForEmail("Для принятия изменений требуется монопольный доступ к базе")
			EndIf
		Else
			AddHTMLBodyForEmail("Не удалось применить изменения динамически.")
			AddHTMLBodyForEmail("База данных " & $IBConnectionParam[3] & " (" & $IBConnectionParam[4] & ")" )
			AddHTMLBodyForEmail("Для принятия изменений требуется монопольный доступ к базе")
		EndIf
		
		If Not PrepareAndSendEmail($sHTMLBodyForEmail) Then
			AddToLog("Подготовка электронного сообщения завершилась ошибкой")
		EndIf
		
		Return False

	Else
		; Успешно применили изменения. Напишем об этом письмо
		; Формируется сообщение об успешном принятии изменений, которое будет отправлено на Email
		AddHTMLBodyForEmail("Выполнено обновление базы данных " & $IBConnectionParam[3] & " (" & $IBConnectionParam[4] & ")." )
		AddHTMLBodyForEmail("Операция выполнена успешно")
		If Not PrepareAndSendEmail($sHTMLBodyForEmail) Then
			AddToLog("Подготовка электронного сообщения завершилась ошибкой")
		EndIf

		Return True

	EndIf
EndFunc

; Функция проверяет необходимость применения изменений
Func CheckUpdate()
	
	Local $TryCount
	
	; Подключение к базе данных
	ConnectToDataBase()

	; Проверяется необходимость обновления конфигурации до процедуры обмена
	AddToLog("Проверка необходимости принять изменения")
	$TryCount = 0 

	While ($connDB.ConfigurationChanged() = True) AND $TryCount <= 3 
		
		$TryCount = $TryCount + 1

		if $connDB.DataBaseConfigurationChangedDynamically() Then
			AddToLog("С момента подключения Конфигурация ИБ была изменена динамически")
			AddToLog("Требуется переподключение к базе данных (перезапуск сеанса)")
			DisconnectFromDatabase()
		Else
			AddToLog("Конфигурация ИБ изменена, требуется применений изменений")
			AddToLog("Попытка применения изменений")
			DisconnectFromDatabase()
			
			If Not RunNonDynamicUpdate() Then
				ExitLoop
			EndIf
		EndIf

		if not ConnectToDataBase() Then
			AddToLog("Выход из цикла из-за ошибки подключения к базе данных")
			ExitLoop
		EndIf

	WEnd

	DisconnectFromDatabase()
	DestroyCOMConnector()

EndFunc

; Возвращает ЛОЖЬ, если нельзя продолжать выполнение скрипта. Запускается всегда вначале программы
Func CanIContinue()
	$mResult	=	False
	$iLastPID	=	GetLastPID()
	
	If ($iLastPID > 0) Then
		If (ProcessExists($iLastPID)) Then
			AddToLog("Процесс " & String($iLastPID) & " работает")
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
; Функции для работы с электронной почтой
; *****************************************************************************

; EmailParam - массив
; 0 - Кому (адреса)
; 1 - Копия (адреса)
; 2 - Тема (строка)
; 3 - Тело сообщения (строка)
Func SendEmail($EmailParam)

	Local $EmailMsg

	AddToLog("Подготовка объектов для формирования электронного сообщения")
	$EmailMsg	= ObjCreate("CDO.Message")

	If @error Then
		AddToLog("Ошибка при создании CDOMessage ")
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
		AddToLog("Ошибка при отправке электронного сообщения")
		Return False
	Else
		Return True
	EndIf

EndFunc

; Готовит тело сообщения и отправляет его
; $EmailMsgTxt - Текст сообщения 
Func PrepareAndSendEmail($EmailMsgTxt)
	Local $mEmlParam[4]

	If $sEmailTo = "" Then
		Return False
	EndIf

	$mEmlParam[0] = $sEmailTo
	$mEmlParam[1] = $sEmailCc
	$mEmlParam[2] = "Автоматическое уведомление об операции над базой данных"
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

; Основная программа
; *****************************************************************************

; Читаем настройки в глобальную переменную
ReadParamsFromIni()

AddToLog("=> Старт обработки")

If ( CanIContinue() ) Then
	
	CreatePIDFile()	
	CheckUpdate()
	DeletePIDFile()	
	AddToLog("<= Обработка завершена");
	
else

	AddToLog("Скрипт не может продолжить работу. Условие CanIContinue не выполнилось");
	AddToLog("<= Скрипт завершает работу");
	
EndIf