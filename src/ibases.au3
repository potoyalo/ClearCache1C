
#include-once

#include <Array.au3>

Func ParseIbases()
	Local $oIbases = CreateDict(1)
	Local $sPath = EnvGet("AppData") & "\1C\1CEStart\"
	If Not FileExists($sPath) Then Return SetError(1, 0, 0)

	Local $oCurrentIB = CreateDict(1)
	Local $aFiles = FindFiles($sPath, '*.v8i')
	For $sFile In $aFiles
		Local $hFile = FileOpen($sPath & $sFile)
		If $hFile = -1 Then ContinueLoop

		While 1
			Local $sLine = FileReadLine($hFile)
			If @error = -1 Then
				SaveCurrentIB($oIbases, $oCurrentIB)
				ExitLoop
			EndIf

			$sLine = StringStripWS($sLine, 3)

			Local $aRes = StringRegExp($sLine, '\[(.+?)\]', 1)
			If Not @error Then
				SaveCurrentIB($oIbases, $oCurrentIB)

				$oCurrentIB = CreateDict(1)
				$oCurrentIB.Item('Name') = StringStripWS($aRes[0], 3)

				ContinueLoop
			EndIf

			$aRes = StringRegExp($sLine, '.+=.+', 1)
			If Not @error Then
				Local $aArr = StringSplit($aRes[0], '=')
				If $aArr[0] = 2 Then
					$oCurrentIB.Item(StringStripWS($aArr[1], 3)) = StringStripWS($aArr[2], 3)
				EndIf
				ContinueLoop
			EndIf
		WEnd

		FileClose($hFile)
	Next

	Return $oIbases
EndFunc   ;==>ParseIbases

Func SaveCurrentIB($oIbases, $oCurrentIB)
	If $oCurrentIB.exists('Name') And $oCurrentIB.exists('ID') Then
		$oIbases.Item($oCurrentIB.Item('ID')) = $oCurrentIB
	EndIf
EndFunc   ;==>SaveCurrentIB


Func FindFiles($sPath, $sMask)
	Local $aDir[0]
	Local $sFile
	Local $hSearch = FileFindFirstFile($sPath & $sMask)
	While 1
		$sFile = FileFindNextFile($hSearch)
		If @error Then ExitLoop
		If @extended = 1 Then ContinueLoop ; файлы обрабатываем, каталоги пропускаем

		_ArrayAdd($aDir, $sFile)
	WEnd
	FileClose($hSearch)

	Return $aDir
EndFunc   ;==>FindFiles

Func CreateDict($CompareMode = 0)
	Local $oDict = ObjCreate('Scripting.Dictionary')
	$oDict.CompareMode = $CompareMode
	Return $oDict
EndFunc   ;==>CreateDict

