#include <File.au3>
#include <Excel.au3>

Local $sFileName = "people.csv"
Local $sFilePath = @ScriptDir & "\"
Local $sResources = "Resources\"

Local $aColumnToRead[] = [6]
Local $aUniquePlans
Local $aPlans
Local $aUnique302 = GetArrayFromFile($aPlans, $sFilePath & $sFileName, 2, 1, $aColumnToRead)
$aUniquePlans = _ArrayUnique(_ArrayExtract($aPlans, -1, -1, 1, 1), 0, 0, 0, 0)

If Not IsArray($aPlans) Or Not IsArray($aUnique302) Then
	ToLog("$aUniquePlans or $aUnique302 contains nothing")
	Exit
EndIf



Local $aColumnToRead302Part1 = [1, 2]
Local $sFileName302Part1 = "302part1.csv"
Local $a302Part1
GetArrayFromFile($a302Part1, $sFilePath & $sResources & $sFileName302Part1, 1, 0, $aColumnToRead302Part1)
ClearNumericValues($a302Part1)


Local $aColumnToRead302Part2 = [1, 2]
Local $sFileName302Part2 = "302part2.csv"
Local $a302Part2
GetArrayFromFile($a302Part2, $sFilePath & $sResources & $sFileName302Part2, 1, 0, $aColumnToRead302Part2)
ClearNumericValues($a302Part2)


Local $aColumnToReadPrice = [1, 2]
Local $sFileNamePrice = "price.csv"
Local $aPrice
GetArrayFromFile($aPrice, $sFilePath & $sResources & $sFileNamePrice, 1, 0, $aColumnToReadPrice)
;~ _ArrayDisplay($aPrice)

Local $aColumnToReadPriceMatching = [1]
Local $sFileNamePriceMatching = "price_matching.csv"
Local $aPriceMatching
GetArrayFromFile($aPriceMatching, $sFilePath & $sResources & $sFileNamePriceMatching, 1, 0, $aColumnToReadPriceMatching)
;~ _ArrayDisplay($aPriceMatching)



Parse302String($aUnique302, $a302Part1, $a302Part2)
ParsePlans($aUniquePlans, $aUnique302)
CreateExcelFile($aUniquePlans)


Func CreateExcelFile($aArray)
	If Not UBound($aArray, $UBOUND_ROWS) Then
		ToLog("CreateExcelFile array doesn't contain rows")
		Return
	EndIf

	Local $oExcel = _Excel_Open()
	If Not IsObj($oExcel) Then
		ToLog("CreateExcelFile cannot create excel instance")
		Return
	EndIf

	Local $oBook = _Excel_BookNew($oExcel, 1)
	If Not IsObj($oBook) Then
		ToLog("CreateExcelFile cannot create excel book")
		Return
	EndIf

	For $i = 0 To UBound($aArray, $UBOUND_ROWS) - 1
		_Excel_RangeWrite($oBook, Default, $aArray[$i][0], "A" & ($i + 1))
		_Excel_RangeWrite($oBook, Default, "Лист " & $i + 2, "B" & ($i + 1))
	Next

	For $i = 0 To UBound($aArray, $UBOUND_ROWS) - 1
		_Excel_SheetAdd($oBook, Default, False)

		_Excel_RangeWrite($oBook, Default, $aArray[$i][0], "A1")
		Local $nRow = 2

		For $y = 1 To 2
			Local $aSubArray = $aArray[$i][$y]
			If Not $aSubArray Then ContinueLoop

			$aSubArray = StringSplit($aSubArray, "|", $STR_NOCOUNT)

			For $x = 0 To UBound($aSubArray) - 1
				Local $sCurrent = $aSubArray[$x]
				If Not $sCurrent Then ContinueLoop
				If StringInStr($sCurrent, "@") Then
					$sCurrent = StringSplit($sCurrent, "@", $STR_NOCOUNT)
					_Excel_RangeWrite($oBook, Default, $sCurrent[0], "A" & $nRow)
					_Excel_RangeWrite($oBook, Default, $sCurrent[1], "B" & $nRow)
				Else
					_Excel_RangeWrite($oBook, Default, $sCurrent, "A" & $nRow)
				EndIf
				$nRow += 1
			Next
		Next
	Next



	_Excel_BookSaveAs($oBook, $sFilePath & "results_" & @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & ".xlsx")
	_Excel_Close($oExcel, False, True)
EndFunc


Func ParsePlans(ByRef $aPlansToCalculate, $aStandalonePlans)
	If Not IsArray($aPlansToCalculate) Or Not IsArray($aStandalonePlans) Then
		ToLog("ParsePlans parameters are not arrays")
		Return
	EndIf

	If UBound($aStandalonePlans, $UBOUND_COLUMNS) <> 3 Then
		ToLog("ParsePlans second array must have 3 columns")
		Return
	EndIf

	_ArrayColInsert($aPlansToCalculate, 1)
	_ArrayColInsert($aPlansToCalculate, 1)

	For $i = 0 To UBound($aPlansToCalculate, $UBOUND_ROWS) - 1
		Local $aCurrent = StringSplit($aPlansToCalculate[$i][0], "|", $STR_NOCOUNT)
		If Not IsArray($aCurrent) Then
			ToLog("Cannot split string: " & $aPlansToCalculate[$i])
			ContinueLoop
		EndIf

		For $x = 0 To UBound($aCurrent) - 1
			If Not $aCurrent[$x] Then ContinueLoop

			Local $nIndex = _ArraySearch($aStandalonePlans, $aCurrent[$x])
			If $nIndex = -1 Then
				ToLog("Cannot find values for: " & $aCurrent[$x])
				ContinueLoop
			EndIf

			$aPlansToCalculate[$i][1] &= $aStandalonePlans[$nIndex][1]
			$aPlansToCalculate[$i][2] &= $aStandalonePlans[$nIndex][2]
		Next

;~ 		_ArrayDisplay($aPlansToCalculate)

		For $x = 1 To 2
			If Not $aPlansToCalculate[$i][$x] Then ContinueLoop
			Local $sCurrent = StringLeft($aPlansToCalculate[$i][$x], StringLen($aPlansToCalculate[$i][$x]) - 1)
			Local $aCurrent = StringSplit($sCurrent, "|", $STR_NOCOUNT)
			$aCurrent = _ArrayUnique($aCurrent, 0, 0, 0, 0)

			For $y = 0 To UBound($aCurrent) - 1
				Local $nIndex = _ArraySearch($aPriceMatching, $aCurrent[$y])

				ToLog("$nIndex: " & $nIndex)

				If $nIndex = -1 Then
					ToLog("cannot find " & $aCurrent[$y] & " in priceMatching")
					ContinueLoop
				EndIf

				Local $sSearchTmp = $aPriceMatching[$nIndex][1] & ".1"
				$sSearchTmp = StringReplace($sSearchTmp, "|", "")
				ToLog("$sSearchTmp: " & $sSearchTmp)

				Local $nIndex2 = _ArraySearch($aPrice, $sSearchTmp)

				ToLog("$nIndex2: " & $nIndex2)

				If $nIndex2 = -1 Then
					ToLog("cannot find " & $sSearchTmp & " in price")
					ContinueLoop
				EndIf

				$aCurrent[$y] &= "@" & $aPrice[$nIndex2][2]

			Next



			$aPlansToCalculate[$i][$x] = _ArrayToString($aCurrent)
		Next

		$aPlansToCalculate[$i][0] = StringReplace(StringLeft($aPlansToCalculate[$i][0], _
			StringLen($aPlansToCalculate[$i][0]) - 1), "|", " | ")
	Next

;~ 	_ArrayDisplay($aPlansToCalculate)
EndFunc


Func Parse302String(ByRef $aArray, $a302Part1, $a302Part2)
	_ArrayColInsert($aArray, 1)
	_ArrayColInsert($aArray, 1)

	For $i = 0 To UBound($aArray) - 1
		Local $aTmp[0]
		Local $sCur = ""
		Local $nLen = StringLen($aArray[$i][0])

		For $y = 1 To $nLen
			Local $sTmp = StringMid($aArray[$i][0], $y, 1)

			If StringIsDigit($sTmp) Or ($sCur And $sTmp = ".") Then
				$sCur &= $sTmp
				If $y = $nLen Then _ArrayAdd($aTmp, $sCur)
			Else
				If $sCur Then
					_ArrayAdd($aTmp, $sCur)
					$sCur = ""
				EndIf
			EndIf
		Next

		If UBound($aTmp) <> 2 Then
			ToLog("!!! 302 string doesn't contain two digits")
			ContinueLoop
		EndIf

		For $x = 0 To 1
			While StringRight($aTmp[$x], 1) = "."
				$aTmp[$x] = StringLeft($aTmp[$x], StringLen($aTmp[$x]) - 1)
			WEnd

			$aArray[$i][$x + 1] = $aTmp[$x]
		Next

		Local $sWhatToSearch = ""
		Local $aArrayToSearch = ""

		If StringInStr($aArray[$i][1], ".") Then
			$sWhatToSearch = $aArray[$i][1]
			If $aArray[$i][2] = 1 Then
				$aArrayToSearch = $a302Part1
			ElseIf $aArray[$i][2] = 2 Then
				$aArrayToSearch = $a302Part2
			Else
				ToLog("Cannot parse 302 string to correct value")
				ContinueLoop
			EndIf
		Else
			If $aArray[$i][1] = 2 Then
				$sWhatToSearch = $aArray[$i][2]
				$aArrayToSearch = $a302Part2
			ElseIf $aArray[$i][1] = 1 And StringInStr($aArray[$i][2], ".") Then
				$sWhatToSearch = $aArray[$i][2]
				$aArrayToSearch = $a302Part1
			Else
				$sWhatToSearch = $aArray[$i][1]
				$aArrayToSearch = $a302Part2
			EndIf
		EndIf

		If $sWhatToSearch And IsArray($aArrayToSearch) Then
			Local $nIndex = _ArraySearch($aArrayToSearch, $sWhatToSearch, 0, 0)
			If $nIndex = -1 Then
				ToLog("Cannot find value: " & $sWhatToSearch)
				ContinueLoop
			EndIf

			$aArray[$i][1] = $aArrayToSearch[$nIndex][1]
			$aArray[$i][2] = $aArrayToSearch[$nIndex][2]
		Else
			ToLog("Cannot define what to search or array to search")
		EndIf
	Next
EndFunc


Func ClearNumericValues(ByRef $aArray)
	If Not IsArray($aArray) Then
		ToLog("ClearNumericValues $aArray is not an array")
		Return
	EndIf

	For $i = 0 To UBound($aArray, $UBOUND_ROWS) - 1
		$sCur = $aArray[$i][0]
		$sClear = ""
		For $x = 1 To StringLen($sCur)
			Local $sSymbol = StringMid($sCur, $x, 1)
			If StringIsDigit($sSymbol) Or ($sClear And $sSymbol = ".") Then _
				$sClear &= $sSymbol
			If $sClear And $sSymbol = " " Then ExitLoop
		Next

		If Not $sClear Then
			ToLog("ClearNumericValues cannot clear value")
			ContinueLoop
		EndIf

		If StringRight($sClear, 1) = "." Then $sClear = StringLeft($sClear, StringLen($sClear) - 1)
		$aArray[$i][0] = $sClear
	Next
EndFunc


Func GetArrayFromFile(ByRef $toReturn, $sFullFileName, $nRowStart, $nColumnToSearch, $aColumnsToRead)
	If Not FileExists($sFullFileName) Then
		ToLog("File: " & $sFullFileName & " doesn't exist")
		Return
	EndIf

	Local $aFileContent
	_FileReadToArray($sFullFileName, $aFileContent, BitOr($FRTA_NOCOUNT, $FRTA_ENTIRESPLIT, $FRTA_INTARRAYS), ";")
	If Not IsArray($aFileContent) Then
		ToLog("File: " & $sFullFileName & " doesn't contain any rows")
		Return
	EndIf

	Local $aResult[0][1 + UBound($aColumnsToRead)]
	Local $aUniqueValues[0]
	Local $nTotalRow = UBound($aFileContent) - 1

	For $nMainRow = $nRowStart To $nTotalRow
		Local $aCurrentRow = $aFileContent[$nMainRow]

		If $aCurrentRow[$nColumnToSearch] Then
			Local $aTemp[1][1 + UBound($aColumnsToRead)]
			$aTemp[0][0] = StringReplace($aCurrentRow[$nColumnToSearch], '"', "")

			Do
				For $nColumnElement = 0 To UBound($aColumnsToRead) - 1
					Local $sCurrent = $aCurrentRow[$aColumnsToRead[$nColumnElement]]
					If Not $sCurrent Then ContinueLoop
					If StringInStr($sCurrent, "!") Then ContinueLoop

					If _ArraySearch($aUniqueValues, $sCurrent) = -1 Then _ArrayAdd($aUniqueValues, $sCurrent)

					$aTemp[0][$nColumnElement + 1] &= StringReplace($sCurrent, '"', "") & "|"
				Next

				$nMainRow += 1
				If $nMainRow > $nTotalRow Then ExitLoop
				$aCurrentRow = $aFileContent[$nMainRow]
			Until ($aFileContent[$nMainRow])[$nColumnToSearch]

			$nMainRow -= 1

			_ArrayAdd($aResult, $aTemp)
		EndIf
	Next

	If UBound($aResult) Then
		$toReturn = $aResult
		Return $aUniqueValues
	EndIf

	$toReturn = 0
EndFunc


Func ToLog($sText)
	ConsoleWrite($sText & @CRLF)
EndFunc