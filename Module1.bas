Attribute VB_Name = "Module1"
Public railCn As ADODB.Connection
Public saveUpdate As Integer
Public userAccountID As Long
Public trainNo As Long
Sub main()
    Set railCn = New ADODB.Connection
    railCn.CursorLocation = adUseClient
    railCn.Provider = "microsoft.ACE.OLEDB.12.0"
    railCn.ConnectionString = "data source=" & App.Path & "\railway.accdb"
    railCn.Open
    If railCn.State = 1 Then
        frmSplash.Show
    End If
End Sub

Public Function comboSearch(combo As ComboBox, fieldID As Integer) As Integer
    For i = 0 To combo.ListCount - 1
        If combo.ItemData(i) = fieldID Then
            comboSearch = i
            Exit For
        End If
    Next
End Function
Public Function convertTime(time1 As String) As String
'debug.print((time1, 10, 2) = "PM" And Val(Mid(time1, 1, 1)) >= 1)
    If Mid(time1, 1, 2) = "12" And Mid(time1, 10, 2) = "AM" Then
        convertTime = "00" & ":" & Mid(time1, 4, 2)
    ElseIf Mid(time1, 1, 2) = "12" And Mid(time1, 10, 2) = "PM" Then
        convertTime = "12" & ":" & Mid(time1, 4, 2)
    ElseIf Mid(time1, 10, 2) = "PM" Or Mid(time1, 9, 2) = "PM" Then
        If Mid(time1, 2, 1) = ":" Then
            If Val(Mid(time1, 1, 1)) >= 1 Then
                convertTime = (12 + Val(Mid(time1, 1, 2))) & ":" & Mid(time1, 3, 2)
            End If
        Else
            If Val(Mid(time1, 1, 1)) >= 1 Then
                convertTime = (12 + Val(Mid(time1, 1, 2))) & ":" & Mid(time1, 4, 2)
            End If
        End If
    Else
         If Mid(time1, 5, 1) <> ":" Then
            convertTime = Mid(time1, 1, 5)
         Else
            convertTime = "0" & Mid(time1, 1, 4)
         End If
    End If
End Function

Public Sub validation(num As Integer, key As Integer, txtBox As TextBox)
    Select Case (num)
        Case 1  'for numeric
                If IsNumeric(Chr(key)) = False And Not key = 8 Then
                    key = 0
                End If
        Case 2  'for alphabets
                If (key < 65 Or key > 90) And (key < 97 Or key > 122) And Not key = 8 And Not key = 32 Then
                    key = 0
                End If
        Case 3  'uppercase conversion
                If (key < 65 Or key > 90) And (key < 97 Or key > 122) And Not key = 8 And Not key = 32 Then
                    key = 0
                ElseIf (key >= 97 And key <= 122) Then
                    key = key - 32
                End If
        End Select
End Sub
Public Function addTime(time1 As String, time2 As String) As String
Dim time3
Dim temp As String
    If (Val(Mid(time1, 4, 2)) + Val(Mid(time2, 4, 2))) >= 60 Then
        time3 = (Val(Mid(time1, 4, 2)) + Val(Mid(time2, 4, 2))) - 60
        If Len(time3) = 1 Then
            time3 = "0" & time3
        End If
        temp = Val(Mid(time1, 1, 2)) + Val(Mid(time2, 1, 2)) + 1
        If Val(temp) < 10 Then
            time3 = "0" & (Val(Mid(time1, 1, 2)) + Val(Mid(time2, 1, 2)) + 1) & ":" & time3
        Else
            time3 = (Val(Mid(time1, 1, 2)) + Val(Mid(time2, 1, 2)) + 1) & ":" & time3
        End If
    Else
        time3 = Val(Mid(time1, 4, 2)) + Val(Mid(time2, 4, 2))
        If Len(time3) = 1 Then
            time3 = "0" & time3
        End If
        temp = Val(Mid(time1, 1, 2)) + Val(Mid(time2, 1, 2))
        If Val(temp) < 10 Then
            time3 = "0" & (Val(Mid(time1, 1, 2)) + Val(Mid(time2, 1, 2))) & ":" & time3
        Else
            time3 = (Val(Mid(time1, 1, 2)) + Val(Mid(time2, 1, 2))) & ":" & time3
        End If
    End If
    If Len(time3) = 4 Then
        time3 = Mid(time3, 1, 2) & ":0" & Mid(time3, 4, 1)
    End If
    addTime = time3
End Function
Public Function diffTime(time1 As String, time2 As String) As String
    Dim time3
    If Val(Mid(time1, 1, 2)) = Val(Mid(time2, 1, 2)) Then
        If Val(Mid(time1, 4, 2)) > Val(Mid(time2, 4, 2)) Then
            time3 = "00" & ":" & (Val(Mid(time1, 4, 2)) - Val(Mid(time2, 4, 2)))
'        Else
'            time3 = "00" & ":" & (Val(Mid(time2, 4, 2)) - Val(Mid(time1, 4, 2)))
'        End If
'        If Val(Mid(time1, 4, 2)) < Val(Mid(time2, 4, 2)) Then
'            time3 = time3 & ":" & (Val(Mid(time2, 4, 2)) - Val(Mid(time1, 4, 2)))
        Else
            time3 = 24 - Val(Mid(time2, 1, 2)) + Val(Mid(time1, 1, 2))
            If Val(time3) < 10 Then
                time3 = "0" & time3
            End If
            time3 = Val(time3) - 1
            If Val(time3) < 10 Then
                time3 = "0" & time3
            End If
            time3 = time3 & ":" & (Val(Mid(time2, 4, 2)) - Val(Mid(time1, 4, 2)))
        End If
    ElseIf Val(Mid(time1, 1, 2)) > Val(Mid(time2, 1, 2)) Then
        time3 = Val(Mid(time1, 1, 2)) - Val(Mid(time2, 1, 2))
        If Val(time3) < 10 Then
            time3 = "0" & time3
        End If
        If Val(Mid(time1, 4, 2)) > Val(Mid(time2, 4, 2)) Then
            time3 = time3 & ":" & (Val(Mid(time1, 4, 2)) - Val(Mid(time2, 4, 2)))
        Else
            time3 = Val(time3) - 1
            If Val(time3) < 10 Then
                time3 = "0" & time3
            End If
            time3 = time3 & ":" & 60 - (Val(Mid(time2, 4, 2)) - Val(Mid(time1, 4, 2)))
        End If
    ElseIf Val(Mid(time1, 1, 2)) < Val(Mid(time2, 1, 2)) Then
        time3 = 24 - Val(Mid(time2, 1, 2)) + Val(Mid(time1, 1, 2))
        If Val(time3) < 10 Then
            time3 = "0" & time3
        End If

        If Val(Mid(time1, 4, 2)) < Val(Mid(time2, 4, 2)) Then
            time3 = time3 & ":" & (Val(Mid(time2, 4, 2)) - Val(Mid(time1, 4, 2)))
        Else
            time3 = Val(time3) - 1
            If Val(time3) < 10 Then
                time3 = "0" & time3
            End If
            time3 = time3 & ":" & 60 - (Val(Mid(time1, 4, 2)) - Val(Mid(time2, 4, 2)))
        End If
    End If
    If Val(Mid(time3, 4, 2)) = 60 Then
        If Val(Mid(time3, 1, 2)) < 10 Then
            time3 = "0" & (Val(Mid(time3, 1, 2)) + 1) & ":00"
        Else
            time3 = (Val(Mid(time3, 1, 2)) + 1) & ":00"
        End If
    End If
    If Len(time3) = 4 Then
        time3 = Mid(time3, 1, 2) & ":0" & Mid(time3, 4, 1)
    End If
    diffTime = time3
End Function

