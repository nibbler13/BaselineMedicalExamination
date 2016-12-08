#include <File.au3>
#include <Excel.au3>

Local $sFileName = "gbuz.csv"
Local $sFilePath = @ScriptDir & "\"
Local $sResources = "Resources\"

Local $aSystemMarks[] = ["!", "m", "w", "]"]


Local $nWorkerQuantity = 1516
Local $nWorkerCode = 1
If $nWorkerQuantity > 50 And $nWorkerQuantity <= 100 Then
	$nWorkerCode = 2
ElseIf $nWorkerQuantity > 100 And $nWorkerQuantity <= 300 Then
	$nWorkerCode = 3
ElseIf $nWorkerQuantity > 300 And $nWorkerQuantity <= 500 Then
	$nWorkerCode = 4
ElseIf $nWorkerQuantity > 500 Then
	$nWorkerCode = 5
EndIf

Local $aColumnToRead[] = [2]
Local $aColumnManWomanWoman40[] = [3, 4, 5]
Local $aUniquePlans
Local $aPlans
Local $aUnique302 = GetArrayFromFile($aPlans, $sFilePath & $sFileName, 5, 1, $aColumnToRead);, $aColumnManWomanWoman40)
$aUniquePlans = _ArrayUnique(_ArrayExtract($aPlans, -1, -1, 1, 1), 0, 0, 0, 0)
_ArrayDisplay($aUniquePlans)

If Not IsArray($aPlans) Or Not IsArray($aUnique302) Then
	ToLog("$aUniquePlans or $aUnique302 contains nothing")
	Exit
EndIf

Local $aColumnToRead302Part1 = [1, 2]
Local $sFileName302Part1 = "302part1_new.csv"
Local $a302Part1
GetArrayFromFile($a302Part1, $sFilePath & $sResources & $sFileName302Part1, 1, 0, $aColumnToRead302Part1)
ClearNumericValues($a302Part1)
;~ _ArrayDisplay($a302Part1)


Local $aColumnToRead302Part2 = [1, 2]
Local $sFileName302Part2 = "302part2_new.csv"
Local $a302Part2
GetArrayFromFile($a302Part2, $sFilePath & $sResources & $sFileName302Part2, 1, 0, $aColumnToRead302Part2)
ClearNumericValues($a302Part2)
;~ _ArrayDisplay($a302Part2)


Local $aColumnToReadPrice = [1, 2]
Local $sFileNamePrice = "price_new.csv"
Local $aPrice
GetArrayFromFile($aPrice, $sFilePath & $sResources & $sFileNamePrice, 1, 0, $aColumnToReadPrice)
;~ _ArrayDisplay($aPrice)

Local $aColumnToReadPriceMatching = [1, 2]
Local $sFileNamePriceMatching = "service_matching_new.csv"
Local $aPriceMatching
GetArrayFromFile($aPriceMatching, $sFilePath & $sResources & $sFileNamePriceMatching, 1, 0, $aColumnToReadPriceMatching)
;~ _ArrayDisplay($aPriceMatching)

Local $aColumnToReadNecesserily = [1]
Local $sFileNameNecesserily = "necesserily.csv"
Local $aNecesserily
GetArrayFromFile($aNecesserily, $sFilePath & $sResources & $sFileNameNecesserily, 0, 0, $aColumnToReadNecesserily)
;~ _ArrayDisplay($aNecesserily)


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

	Local $aHeader[] = ["Название мероприятий", _
						"Код услуги", _
						"Название услуги", _
						"Мужчины", _
						"Женщины до 40 лет", _
						"Женщины после 40 лет"]

	_ArrayColInsert($aArray, 2)
	_ArrayColInsert($aArray, 3)
	_ArrayColInsert($aArray, 4)

	For $nPlanRow = 0 To UBound($aArray, $UBOUND_ROWS) - 1
		_Excel_RangeWrite($oBook, Default, $aArray[$nPlanRow][0], "A1")

		For $nHeaderRow = 0 To UBound($aHeader) - 1
			_Excel_RangeWrite($oBook, Default, $aHeader[$nHeaderRow], _
				_Excel_ColumnToLetter($nHeaderRow + 1) & "2")
		Next

		Local $nRow = 3

		Local $nTotalMan = 0
		Local $nTotalWomen = 0
		Local $nTotalWomen40 = 0

		Local $aSubArray = $aArray[$nPlanRow][1]
		If Not IsArray($aSubArray) Then ContinueLoop

		Local $bOptional = False

		For $x = 0 To UBound($aSubArray, $UBOUND_ROWS) - 1
			Local $sName = $aSubArray[$x][0]
			Local $sCode = $aSubArray[$x][1]
			Local $sServiceName = $aSubArray[$x][2]
			Local $nPrice = $aSubArray[$x][3]
			Local $sDescription = $aSubArray[$x][4]

			If StringLeft($sName, 1) = "w" Then
				$sName = StringReplace($sName, "w", "для_женщин_")
				$nTotalWomen += $nPrice
				$nTotalWomen40 += $nPrice
				_Excel_RangeWrite($oBook, Default, "---", "D" & $nRow)
				_Excel_RangeWrite($oBook, Default, $nPrice, "E" & $nRow)
				_Excel_RangeWrite($oBook, Default, $nPrice, "F" & $nRow)
			ElseIf StringLeft($sName, 1) = "]" Then
				$sName = StringReplace($sName, "]", "для_женщин_старше40_")
				$nTotalWomen40 += $nPrice
				_Excel_RangeWrite($oBook, Default, "---", "D" & $nRow)
				_Excel_RangeWrite($oBook, Default, "---", "E" & $nRow)
				_Excel_RangeWrite($oBook, Default, $nPrice, "F" & $nRow)
			ElseIf StringLeft($sName, 1) = "!" Then
				$sName = StringReplace($sName, "!", "")

				If Not $bOptional Then
					$bOptional = True
					$nRow += 2
					_Excel_RangeWrite($oBook, Default, "Опциональные услуги:", "A" & $nRow)
					$nRow += 1
				EndIf

				If StringLeft($sName, 1) = "m" Then
					$sName = StringReplace($sName, "m", "для_мужчин_")
					_Excel_RangeWrite($oBook, Default, $nPrice, "D" & $nRow)
					_Excel_RangeWrite($oBook, Default, "---", "E" & $nRow)
					_Excel_RangeWrite($oBook, Default, "---", "F" & $nRow)
				ElseIf StringLeft($sName, 1) = "w" Then
					$sName = StringReplace($sName, "w", "для_женщин_")
					_Excel_RangeWrite($oBook, Default, "---", "D" & $nRow)
					_Excel_RangeWrite($oBook, Default, $nPrice, "E" & $nRow)
					_Excel_RangeWrite($oBook, Default, $nPrice, "F" & $nRow)
				ElseIf StringLeft($sName, 1) = "]" Then
					$sName = StringReplace($sName, "m", "для_женщин_старше40_")
					_Excel_RangeWrite($oBook, Default, "---", "D" & $nRow)
					_Excel_RangeWrite($oBook, Default, "---", "E" & $nRow)
					_Excel_RangeWrite($oBook, Default, $nPrice, "F" & $nRow)
				Else
					_Excel_RangeWrite($oBook, Default, $nPrice, "D" & $nRow)
					_Excel_RangeWrite($oBook, Default, $nPrice, "E" & $nRow)
					_Excel_RangeWrite($oBook, Default, $nPrice, "F" & $nRow)
				EndIf
			ElseIf StringLeft($sName, 1) = "m" Then
				$sName = StringReplace($sName, "m", "для_мужчин_")
				$nTotalMan += $nPrice
				_Excel_RangeWrite($oBook, Default, $nPrice, "D" & $nRow)
				_Excel_RangeWrite($oBook, Default, "---", "E" & $nRow)
				_Excel_RangeWrite($oBook, Default, "---", "F" & $nRow)
			Else
				$nTotalMan += $nPrice
				$nTotalWomen += $nPrice
				$nTotalWomen40 += $nPrice
				_Excel_RangeWrite($oBook, Default, $nPrice, "D" & $nRow)
				_Excel_RangeWrite($oBook, Default, $nPrice, "E" & $nRow)
				_Excel_RangeWrite($oBook, Default, $nPrice, "F" & $nRow)
			EndIf

			_Excel_RangeWrite($oBook, Default, $sName, "A" & $nRow)
			_Excel_RangeWrite($oBook, Default, $sCode, "B" & $nRow)
			_Excel_RangeWrite($oBook, Default, $sServiceName, "C" & $nRow)

			$nRow += 1
		Next

		$aArray[$nPlanRow][2] = $nTotalMan
		$aArray[$nPlanRow][3] = $nTotalWomen
		$aArray[$nPlanRow][4] = $nTotalWomen40

		_Excel_RangeWrite($oBook, Default, $nTotalMan, "D1")
		_Excel_RangeWrite($oBook, Default, $nTotalWomen, "E1")
		_Excel_RangeWrite($oBook, Default, $nTotalWomen40, "F1")

		$oBook.ActiveSheet.Range("A1:F2").Font.Bold = True
		$oBook.ActiveSheet.Columns("A").ColumnWidth = 50
		$oBook.ActiveSheet.Columns("B").ColumnWidth = 10
		$oBook.ActiveSheet.Columns("C").ColumnWidth = 90
		$oBook.ActiveSheet.Columns("D").ColumnWidth = 10
		$oBook.ActiveSheet.Columns("E").ColumnWidth = 20
		$oBook.ActiveSheet.Columns("F").ColumnWidth = 22

		Sleep(1000)

		_Excel_SheetAdd($oBook, Default, False)
	Next

	_Excel_RangeWrite($oBook, Default, "Состав плана", "A1")
	_Excel_RangeWrite($oBook, Default, "Мужчины", "B1")
	_Excel_RangeWrite($oBook, Default, "Женщины до 40 лет", "C1")
	_Excel_RangeWrite($oBook, Default, "Женщины после 40 лет", "D1")
	_Excel_RangeWrite($oBook, Default, "Расположение", "E1")

	For $i = 0 To UBound($aArray, $UBOUND_ROWS) - 1
		_Excel_RangeWrite($oBook, Default, $aArray[$i][0], "A" & ($i + 2))
		_Excel_RangeWrite($oBook, Default, $aArray[$i][2], "B" & ($i + 2))
		_Excel_RangeWrite($oBook, Default, $aArray[$i][3], "C" & ($i + 2))
		_Excel_RangeWrite($oBook, Default, $aArray[$i][4], "D" & ($i + 2))
		_Excel_RangeWrite($oBook, Default, "Лист " & $i + 1, "E" & ($i + 2))
	Next

	$oBook.ActiveSheet.Range("A1:E1").Font.Bold = True
	$oBook.ActiveSheet.Columns("A").ColumnWidth = 140
	$oBook.ActiveSheet.Columns("B").ColumnWidth = 10
	$oBook.ActiveSheet.Columns("C").ColumnWidth = 19
	$oBook.ActiveSheet.Columns("D").ColumnWidth = 22
	$oBook.ActiveSheet.Columns("E").ColumnWidth = 14


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

	For $iMainCurrentRow = 0 To UBound($aPlansToCalculate, $UBOUND_ROWS) - 1
		Local $aCurrent = StringSplit($aPlansToCalculate[$iMainCurrentRow][0], "|", $STR_NOCOUNT)
		If Not IsArray($aCurrent) Then
			ToLog("Cannot split string: " & $aPlansToCalculate[$iMainCurrentRow])
			ContinueLoop
		EndIf

		For $x = 0 To UBound($aCurrent) - 1
			If Not $aCurrent[$x] Then ContinueLoop

			Local $nIndex = _ArraySearch($aStandalonePlans, $aCurrent[$x])
			If $nIndex = -1 Then
				ToLog("Cannot find values for: " & $aCurrent[$x])
				ContinueLoop
			EndIf

			$aPlansToCalculate[$iMainCurrentRow][1] &= $aStandalonePlans[$nIndex][1]
			$aPlansToCalculate[$iMainCurrentRow][2] &= $aStandalonePlans[$nIndex][2]
		Next

		$aPlansToCalculate[$iMainCurrentRow][1] &= $aPlansToCalculate[$iMainCurrentRow][2]

		Local $aParsedServices[0]
		If $aPlansToCalculate[$iMainCurrentRow][1] Then
			$aParsedServices = StringSplit($aPlansToCalculate[$iMainCurrentRow][1], "|", $STR_NOCOUNT)
		Else
			ContinueLoop
		EndIf

		If IsArray($aNecesserily) Then
			Local $aTmp = _ArrayExtract($aNecesserily, -1, -1, 0, 0)
			_ArrayAdd($aParsedServices, $aTmp)
		EndIf

		If Not UBound($aParsedServices) Then ContinueLoop

		$aParsedServices = _ArrayUnique($aParsedServices, 0, 0, 0, 0)

		For $i = 0 To 3
			_ArrayColInsert($aParsedServices, 1)
		Next

		Local $aTempResult[0][5]

		For $iParsedServicesRow = 0 To UBound($aParsedServices, $UBOUND_ROWS) - 1
			Local $sWhatToSearch = $aParsedServices[$iParsedServicesRow][0]
			If Not $sWhatToSearch Then ContinueLoop

			If _ArraySearch($aSystemMarks, StringLeft($sWhatToSearch, 1)) > -1 Then _
				$sWhatToSearch = StringRight($sWhatToSearch, StringLen($sWhatToSearch) - 1)

			Local $nIndex = _ArraySearch($aPriceMatching, $sWhatToSearch)

			If $nIndex = -1 Then
				ToLog("cannot find '" & $sWhatToSearch & "' in priceMatching")
				ContinueLoop
			EndIf

			$aParsedServices[$iParsedServicesRow][4] = $aPriceMatching[$nIndex][2]

			If Not $aPriceMatching[$nIndex][1] Then ContinueLoop

			Local $aServiceCodes = StringSplit($aPriceMatching[$nIndex][1], "|", $STR_NOCOUNT)

			For $iServiceCodesElement = 0 To UBound($aServiceCodes, $UBOUND_ROWS) - 1
				If Not $aServiceCodes[$iServiceCodesElement] Then ContinueLoop

				Local $aCurrentService[1][5]
				$aCurrentService[0][0] = $aParsedServices[$iParsedServicesRow][0]
				$aCurrentService[0][4] = $aParsedServices[$iParsedServicesRow][4]

				If Not StringIsDigit(StringLeft($aServiceCodes[$iServiceCodesElement], 1)) Then
					If StringLeft($aCurrentService[0][0], 1) = "!" Then
						$aCurrentService[0][0] = "!" & StringLeft($aServiceCodes[$iServiceCodesElement], 1) & _
							StringRight($aCurrentService[0][0], StringLen($aCurrentService[0][0]) - 1)
					Else
						$aCurrentService[0][0] = StringLeft($aServiceCodes[$iServiceCodesElement], 1) & _
							$aCurrentService[0][0]
					EndIf

					$aServiceCodes[$iServiceCodesElement] = StringRight($aServiceCodes[$iServiceCodesElement], _
						StringLen($aServiceCodes[$iServiceCodesElement]) - 1)
				EndIf

				Local $sAddPart = "." & $nWorkerCode
				If StringLen($aServiceCodes[$iServiceCodesElement]) < 6 Or _
					StringLeft($aServiceCodes[$iServiceCodesElement], 1) <> 8 Then _
					$sAddPart = ""

				Local $nIndex2 = _ArraySearch($aPrice, $aServiceCodes[$iServiceCodesElement] & $sAddPart)

				If $nIndex2 = -1 Then
					ToLog("cannot find " & $sWhatToSearch & " in price")
					$aCurrentService[0][1] = $aServiceCodes[$iServiceCodesElement]
				Else
					$aCurrentService[0][1] = $aPrice[$nIndex2][0]
					$aCurrentService[0][2] = StringReplace($aPrice[$nIndex2][1], "|", "")
					$aCurrentService[0][3] = StringReplace($aPrice[$nIndex2][2], "|", "")
				EndIf

				_ArrayAdd($aTempResult, $aCurrentService)
			Next
		Next

		Local $aResult[0][5]

		For $i = 0 To UBound($aTempResult, $UBOUND_ROWS) - 1
			If _ArraySearch($aResult, $aTempResult[$i][1], _
				Default, Default, Default, Default, Default, 1) = -1 Or _
				Not $aTempResult[$i][1] Then
				Local $aTmp = _ArrayExtract($aTempResult, $i, $i)
				If StringInStr($aTmp[0][1], "851000") Then
					If _ArraySearch($aResult, "851005*", Default, Default, Default, Default, Default, 1) > -1 Or _
						_ArraySearch($aResult, "851004*", Default, Default, Default, Default, Default, 1) > -1 Then
						ConsoleWrite("Skipping 851000 because present 851005 or 851004" & @CRLF)
						ContinueLoop
					EndIf
				EndIf
				_ArrayAdd($aResult, $aTmp)
			EndIf
		Next

		_ArraySort($aResult, True)

;~ 		_ArrayDisplay($aResult)

		$aPlansToCalculate[$iMainCurrentRow][1] = $aResult
		$aPlansToCalculate[$iMainCurrentRow][0] = StringReplace(StringLeft($aPlansToCalculate[$iMainCurrentRow][0], _
			StringLen($aPlansToCalculate[$iMainCurrentRow][0]) - 1), "|", " | ")
	Next

	_ArrayColDelete($aPlansToCalculate, 2)

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


Func GetArrayFromFile(ByRef $toReturn, $sFullFileName, $nRowStart, $nColumnToSearch, $aColumnsToRead, $aPeople = 0)
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
;~ 					If StringInStr($sCurrent, "!") Then ContinueLoop


;~ 					If StringInStr($sFullFileName, "match") Then
;~ 						ConsoleWrite($aTemp[0][0] & " : " & $aColumnsToRead[$nColumnElement] & " - " & $sCurrent & @CRLF)
;~ 					EndIf


					If _ArraySearch($aUniqueValues, $sCurrent) = -1 Then _ArrayAdd($aUniqueValues, $sCurrent)

					$aTemp[0][$nColumnElement + 1] &= StringReplace($sCurrent, '"', "") & "|"
				Next

				$nMainRow += 1
				If $nMainRow > $nTotalRow Then ExitLoop
				$aCurrentRow = $aFileContent[$nMainRow]
			Until ($aFileContent[$nMainRow])[$nColumnToSearch]

;~ 			If IsArray($aPeople) Then
;~ 				For $nColumnElement = 0 To UBound
;~ 			EndIf

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