
#AutoIt3Wrapper_Icon="..\res\icon\17.ico"
#AutoIt3Wrapper_OutFile="..\build\Очистка кэша 1С.exe"

#include <Array.au3>
#include <.\ibases.au3>

Opt('MustDeclareVars', 1)
Opt('TrayIconHide', 1)

Local $sGTitle = 'Очистка кэша 1С'
Local $sGMask = '????????-????-????-????-????????????'

ClearCache(GetDelDir())
If @error Then Exit

AlarmSuccess()



Func ClearCache($aDir)
	Local $iRows = UBound($aDir)
	If $iRows = 0 Then Return

	Local $error = 0
	Local $sErrorDelDir = ''
	Local $sPath
	Local $sFile
	Local $oIbases = ParseIbases()
	Local $aDelDir[0]

	ProgressOn($sGTitle, '')
	For $i = 0 To $iRows - 1
		ProgressSet(($i + 1) * 100 / $iRows, 'Удаление каталогов...')
		$sPath = $aDir[$i].Item('Path')
		$sFile = $aDir[$i].Item('File')
		If Not DirRemove($sPath & $sFile, 1) Then
			$error = 1

			If IsObj($oIbases) And $oIbases.Exists($sFile) Then $sFile = $oIbases.Item($sFile).Item('Name')

			_ArraySearch($aDelDir, $sFile)
			If @error Then _ArrayAdd($aDelDir, $sFile)
		EndIf
	Next
	ProgressOff()

	If $error Then
		For $sFile In $aDelDir
			$sErrorDelDir &= '   ' & $sFile & @CRLF
		Next
		AlarmError('Закройте информационные базы' & @CRLF & $sErrorDelDir & 'и попробуйте ещё раз')
	EndIf

	SetError($error)
EndFunc   ;==>ClearCache

Func GetDelDir()
	Local $sPath
	Local $aDir[0]
	Local $iTimeout = 5

	$sPath = GetCacheDir(GetNameLocalAppData())
	If @error Then Exit
	FindDir($aDir, $sPath, $sGMask)
	FindDir($aDir, $sPath, 'Srvr__*__Ref__*__')

	Switch MsgBox(BitOR(4, 32, 256), $sGTitle, 'Очищать перемещаемый профиль?', $iTimeout)
		Case 6
			;Да
			$sPath = GetCacheDir(GetNameAppData())
			If @error Then Exit
			FindDir($aDir, $sPath, $sGMask)
		Case Else
			;Нет, или по таймауту
	EndSwitch

	Return $aDir
EndFunc   ;==>GetDelDir

Func GetNameAppData()
	Return EnvGet("AppData")
EndFunc   ;==>GetNameAppData

Func GetNameLocalAppData()
	Switch @OSVersion
		Case "WIN_XP"
			Return StringReplace(EnvGet("AppData"), "Application Data", "Local Settings\Application Data")
		Case Else
			Return EnvGet("LocalAppData")
	EndSwitch
EndFunc   ;==>GetNameLocalAppData

Func GetCacheDir($sPathLocal)
	Local $sPath = ''

	$sPath = $sPathLocal & "\1C\1cv8\"
	If Not FileExists($sPath) Then
		Local $textErr = 'Не найден каталог ' & $sPath
		$sPath = $sPathLocal & "\1C\1cv82\"
		If Not FileExists($sPath) Then
			AlarmError($textErr)
			SetError(1)
			Return ''
		EndIf
	EndIf

	Local $sLocFile = $sPath & "location.cfg"
	If FileExists($sLocFile) Then
		Local $hFile = FileOpen($sLocFile)
		Local $sPathLoc = FileRead($hFile)
		FileClose($hFile)

		$sPathLoc = StringReplace($sPathLoc, 'location=', '')
		$sPathLoc = StringReplace($sPathLoc, '/', '\')
		$sPathLoc &= '\'

		If FileExists($sPathLoc) Then $sPath = $sPathLoc
	EndIf

	Return $sPath
EndFunc   ;==>GetCacheDir

Func FindDir(ByRef $aDir, $sPath, $sMask)
	Local $sFile
	Local $oDict
	Local $hSearch = FileFindFirstFile($sPath & $sMask)
	While 1
		$sFile = FileFindNextFile($hSearch)
		If @error Then ExitLoop
		If @extended = 0 Then ContinueLoop ; каталоги обрабатываем, файлы пропускаем

		$oDict = CreateDict(1)
		$oDict.Item('Path') = $sPath
		$oDict.Item('File') = $sFile
		_ArrayAdd($aDir, $oDict)
	WEnd
	FileClose($hSearch)
EndFunc   ;==>FindDir



Func AlarmError($text = '')
	MsgBox(16, '', 'Не удалось очистить кэш 1С' & @CRLF & @CRLF & $text)
EndFunc   ;==>AlarmError

Func AlarmSuccess()
	MsgBox(64, '', 'Кэш 1С успешно очищен')
EndFunc   ;==>AlarmSuccess
