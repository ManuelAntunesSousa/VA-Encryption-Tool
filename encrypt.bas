Attribute VB_Name = "encryption sequence"

Private Function StrToPsd(ByVal Txt As String) As Long
    Dim xVal As Long
    Dim xCh As Long
    Dim xSft1 As Long
    Dim xSft2 As Long
    Dim I As Integer
    Dim xLen As Integer
    xLen = Len(Txt)
    For I = 1 To xLen
        xCh = Asc(Mid$(Txt, I, 1))
        xVal = xVal Xor (xCh * 2 ^ xSft1)
        xVal = xVal Xor (xCh * 2 ^ xSft2)
        xSft1 = (xSft1 + 7) Mod 19
        xSft2 = (xSft2 + 13) Mod 23
    Next I
    StrToPsd = xVal
End Function

Private Function Encryption(ByVal Psd As String, ByVal InTxt As String, Optional ByVal Enc As Boolean = True) As String
    Dim xOffset As Long
    Dim xLen As Integer
    Dim I As Integer
    Dim xCh As Integer
    Dim xOutTxt As String
    xOffset = StrToPsd(Psd)
    Rnd -1
    Randomize xOffset
    xLen = Len(InTxt)
    For I = 1 To xLen
        xCh = Asc(Mid$(InTxt, I, 1))
        If xCh >= 32 And xCh <= 126 Then
            xCh = xCh - 32
            xOffset = Int((96) * Rnd)
            If Enc Then
                xCh = ((xCh + xOffset) Mod 95)
            Else
                xCh = ((xCh - xOffset) Mod 95)
                If xCh < 0 Then xCh = xCh + 95
            End If
            xCh = xCh + 32
            xOutTxt = xOutTxt & Chr$(xCh)
        End If
    Next I
    Encryption = xOutTxt
End Function

Private Function Decryption(ByVal Psd As String, ByVal InTxt As String) As String
    Dim xOffset As Long
    Dim xLen As Integer
    Dim I As Integer
    Dim xCh As Integer
    Dim xOutTxt As String
    xOffset = StrToPsd(Psd)
    Rnd -1
    Randomize xOffset
    xLen = Len(InTxt)
    For I = 1 To xLen
        xCh = Asc(Mid$(InTxt, I, 1))
        If xCh >= 32 And xCh <= 126 Then
            xCh = xCh - 32
            xOffset = Int((96) * Rnd)
            xCh = ((xCh - xOffset) Mod 95)
            If xCh < 0 Then xCh = xCh + 95
            xCh = xCh + 32
            xOutTxt = xOutTxt & Chr$(xCh)
        End If
    Next I
    Decryption = xOutTxt
End Function

Sub EncryptionRange()
    Dim xRg As Range
    Dim xPsd As String
    Dim xRet As Variant
    Dim xCell As Range
    Dim wsM As Worksheet
    Dim wsP As Worksheet
    

    Set wsM = ThisWorkbook.Sheets("sheet name 1")
    Set wsP = ThisWorkbook.Sheets("sheet name 2")
    
    ' Prompt password
    xPsd = InputBox("Enter password:")
    If xPsd = "" Then
        MsgBox "Password cannot be empty"
        Exit Sub
    End If
    
    ' Prompt
    xRet = Application.InputBox("Type 1 to encrypt cell(s); Type 2 to decrypt cell(s)",  , , , , , 1)
    If TypeName(xRet) = "Boolean" Then Exit Sub
    
    If xRet > 0 Then
        If xRet Mod 2 = 1 Then ' Encryption

            EncryptDecryptColumns wsM.Range("L5:L" & wsM.Cells(wsM.Rows.Count, "L").End(xlUp).Row), xPsd, True
            EncryptDecryptColumns wsM.Range("O5:O" & wsM.Cells(wsM.Rows.Count, "O").End(xlUp).Row), xPsd, True
            EncryptDecryptColumns wsM.Range("M5:M" & wsM.Cells(wsM.Rows.Count, "M").End(xlUp).Row), xPsd, True
            EncryptDecryptColumns wsM.Range("J5:J" & wsM.Cells(wsM.Rows.Count, "J").End(xlUp).Row), xPsd, True

            EncryptDecryptColumns wsP.Range("B4:B" & wsP.Cells(wsP.Rows.Count, "B").End(xlUp).Row), xPsd, True
            EncryptDecryptColumns wsP.Range("D4:D" & wsP.Cells(wsP.Rows.Count, "D").End(xlUp).Row), xPsd, True
            EncryptDecryptColumns wsP.Range("C4:C" & wsP.Cells(wsP.Rows.Count, "C").End(xlUp).Row), xPsd, True
            EncryptDecryptColumns wsP.Range("E4:E" & wsP.Cells(wsP.Rows.Count, "E").End(xlUp).Row), xPsd, True
            EncryptDecryptColumns wsP.Range("H4:H" & wsP.Cells(wsP.Rows.Count, "H").End(xlUp).Row), xPsd, True
            EncryptDecryptColumns wsP.Range("F4:F" & wsP.Cells(wsP.Rows.Count, "F").End(xlUp).Row), xPsd, True
            Else ' Decryption

            EncryptDecryptColumns wsM.Range("L5:L" & wsM.Cells(wsM.Rows.Count, "L").End(xlUp).Row), xPsd, False
            EncryptDecryptColumns wsM.Range("O5:O" & wsM.Cells(wsM.Rows.Count, "O").End(xlUp).Row), xPsd, False
            EncryptDecryptColumns wsM.Range("M5:M" & wsM.Cells(wsM.Rows.Count, "M").End(xlUp).Row), xPsd, False
            EncryptDecryptColumns wsM.Range("J5:J" & wsM.Cells(wsM.Rows.Count, "J").End(xlUp).Row), xPsd, False

            EncryptDecryptColumns wsP.Range("B4:B" & wsP.Cells(wsP.Rows.Count, "B").End(xlUp).Row), xPsd, False
            EncryptDecryptColumns wsP.Range("D4:D" & wsP.Cells(wsP.Rows.Count, "D").End(xlUp).Row), xPsd, False
            EncryptDecryptColumns wsP.Range("C4:C" & wsP.Cells(wsP.Rows.Count, "C").End(xlUp).Row), xPsd, False
            EncryptDecryptColumns wsP.Range("E4:E" & wsP.Cells(wsP.Rows.Count, "E").End(xlUp).Row), xPsd, False
            EncryptDecryptColumns wsP.Range("H4:H" & wsP.Cells(wsP.Rows.Count, "H").End(xlUp).Row), xPsd, False
            EncryptDecryptColumns wsP.Range("F4:F" & wsP.Cells(wsP.Rows.Count, "F").End(xlUp).Row), xPsd, False
        End If
    End If
End Sub


Sub EncryptDecryptColumns(xRg As Range, xPsd As String, xEnc As Boolean)
    Dim xCell As Range
    Dim xPrefix As String

    For Each xCell In xRg
        If xCell.Value <> "" Then
            If xEnc Then
                If xRg.Worksheet.Name = "sheet name 1" Then
                    If xRg.Column = 12 Then 
                        xPrefix = "Manager"
                    ElseIf xRg.Column = 15 Then 
                        xPrefix = "Client"
                    ElseIf xRg.Column = 10 Then 
                        xPrefix = "NIPC"
                    Else
                        xPrefix = ""
                    End If
                ElseIf xRg.Worksheet.Name = "sheet name 2" Then
                    If xRg.Column = 2 Then 
                        xPrefix = "Manager"
                    ElseIf xRg.Column = 4 Or xRg.Column = 5 Then 
                        xPrefix = "Client"
                    Else
                        xPrefix = ""
                    End If
                End If
                xCell.Value = xPrefix & Encryption(xPsd, xCell.Value, True)
            Else
                If xRg.Worksheet.Name = "sheet name 1" Then
                    If xRg.Column = 12 Then 
                        xPrefix = "Manager"
                    ElseIf xRg.Column = 15 Then 
                        xPrefix = "Client"
                    ElseIf xRg.Column = 10 Then 
                        xPrefix = "NIPC"
                    Else
                        xPrefix = ""
                    End If
                ElseIf xRg.Worksheet.Name = "sheet name 2" Then
                    If xRg.Column = 2 Then
                        xPrefix = "Manager"
                    ElseIf xRg.Column = 4 Or xRg.Column = 5 Then 
                        xPrefix = "Client"
                    Else
                        xPrefix = ""
                    End If
                End If
                xCell.Value = Decryption(xPsd, Mid(xCell.Value, Len(xPrefix) + 1))
            End If
        End If
    Next
End Sub
