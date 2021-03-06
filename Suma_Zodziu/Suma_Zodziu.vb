﻿Public Class Suma_Zodziu
    ''' <summary>
    ''' skaitmeninio tipo kintamojo keitimas į sumą žodžiu
    ''' </summary>
    ''' <returns>gražina sumą žodžiu String, didžiausiai suma baigiasi žodžiu 'tūkstančiai'</returns>
    Public Function GET_SumaZodziu(ByVal Num As Double) As String
        Dim SumaZodziu As String = String.Empty
        Dim Dollars As String = String.Empty
        Dim Cents As String = String.Empty
        Dim Valiuta As String = String.Empty
        Dim DecimalPlace, Count As Integer
        Dim Number As String = String.Empty

        If Num < 0 Then
            Return "-"
        Else
            Number = Num.ToString("#.00")
        End If

        DecimalPlace = InStr(Number, ",")
        If DecimalPlace > 0 Then
            Cents = Left(Mid(Number, DecimalPlace + 1) &
                  "00", 2) & " ct."
            Number = Trim(Left(Number, DecimalPlace - 1))
        Else
            Cents = "00 ct."
        End If

        Count = Len(Number)
        If Number = "" Then
            Dollars = "Nulis eurų "
        Else
            Select Case Count
                Case > 3
                    Dollars = Dollars & GET_Thousand(Right(Number, 6))
                Case > 2
                    Dollars = Dollars & GET_Hundreds(Right(Number, 3))
                Case > 1
                    Dollars = Dollars & GET_Tens(Right(Number, 2))
                    Valiuta = GET_Valiuta(Right(Number, 2))
                Case = 1
                    Dollars = Dollars & GET_Digit(Right(Number, 1))
                    Valiuta = GET_Valiuta(Right(Number, 1))
                Case Else
                    Dollars = "N/G"
            End Select
        End If
        'If Count = 1 Then
        '    Dollars = "Nulis "
        'End If
        SumaZodziu = UCase(Left(Dollars, 1))
        Dollars = Right(Dollars, Len(Dollars) - 1)

        SumaZodziu = SumaZodziu & Dollars & Valiuta & Cents '& " ct."

        Return SumaZodziu
    End Function

    Private Function GET_Valiuta(ByVal Skaic_text)
        If Len(Skaic_text) > 1 Then
            If Left(Skaic_text, 1) = "1" Then
                GET_Valiuta = "eurų "
                Exit Function
            Else
                Skaic_text = Right(Skaic_text, 1)
            End If
        End If
        Select Case Skaic_text
            Case "0" : GET_Valiuta = "eurų "
            Case "1" : GET_Valiuta = "euras "
            Case Else : GET_Valiuta = "eurai "
        End Select
    End Function

    Private Function GET_Digit(Skaic_text)

        Select Case Skaic_text
            Case "1" : GET_Digit = "vienas "
            Case "2" : GET_Digit = "du "
            Case "3" : GET_Digit = "trys "
            Case "4" : GET_Digit = "keturi "
            Case "5" : GET_Digit = "penki "
            Case "6" : GET_Digit = "šeši "
            Case "7" : GET_Digit = "septyni "
            Case "8" : GET_Digit = "aštuoni "
            Case "9" : GET_Digit = "devyni "
            Case Else : GET_Digit = ""
        End Select

    End Function

    Private Function GET_Tens(TensText)
        Dim Result As String
        Result = ""           ' Valome rezultatą
        If Val(Left(TensText, 1)) = 1 Then   'Jei tarp 10-19...
            Select Case Val(TensText)
                Case "10" : Result = "dešimt "
                Case "11" : Result = "venuolika "
                Case "12" : Result = "dvylika "
                Case "13" : Result = "trylika "
                Case "14" : Result = "keturiolika "
                Case "15" : Result = "penkiolika "
                Case "16" : Result = "šešiolika "
                Case "17" : Result = "septyniolika "
                Case "18" : Result = "aštuoniolika "
                Case "19" : Result = "devyniolika "
                Case Else
            End Select
        Else  ' jei tarp 20-99...
            Select Case Val(Left(TensText, 1))
                Case 2 : Result = "dvidešimt "
                Case 3 : Result = "trisdešimt "
                Case 4 : Result = "keturiasdešimt "
                Case 5 : Result = "penkiasdešimt "
                Case 6 : Result = "šešiasdešimt "
                Case 7 : Result = "septyniasdešimt "
                Case 8 : Result = "aštuoniasdešimt "
                Case 9 : Result = "devyniasdešimt "
                Case Else
            End Select
            Result = Result & GET_Digit(Right(TensText, 1))  ' vienetų vieta
        End If
        GET_Tens = Result
    End Function

    Private Function GET_Hundreds(ByVal Number)
        Number = Left(Number, 1)
        Select Case Number
            Case "0" : GET_Hundreds = ""
            Case "1" : GET_Hundreds = GET_Digit(Number) & "šimtas "
            Case Else : GET_Hundreds = GET_Digit(Number) & "šimtai "
        End Select
    End Function

    Private Function GET_Thousand(ByVal Number)
        Dim Tukst As String
        Tukst = Left(Number, Len(Number) - 3)
        If Len(Tukst) > 2 Then GET_Thousand = GET_Hundreds(Tukst)
        If Len(Tukst) > 1 And Right(Tukst, 2) > 9 Then
            GET_Thousand = GET_Thousand & GET_Tens(Right(Tukst, 2))
        Else
            GET_Thousand = GET_Thousand & GET_Digit(Right(Tukst, 1))
        End If
        If Tukst > 10 And Right(Tukst, 2) > 10 And Right(Tukst, 2) < 20 Then
            GET_Thousand = GET_Thousand & "tūkstančių "
        Else
            If Right(Tukst, 1) = "1" Then
                GET_Thousand = GET_Thousand & "tūkstantis "
            ElseIf Right(Tukst, 1) = "0" Then
                GET_Thousand = GET_Thousand & "tūkstančių "
            Else
                GET_Thousand = GET_Thousand & "tūkstančiai "
            End If
        End If

    End Function

End Class
