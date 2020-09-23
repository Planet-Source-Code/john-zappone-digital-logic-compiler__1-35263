Attribute VB_Name = "CodeMod"
Sub ColorCode()
For i = 0 To Len(Form1.Code.Text)
strCommentStart = InStr(1, Form1.Code.Text, "//")
If strCommentStart <> 0 = True Then
strCommentEnd = InStr(strCommentStart, Form1.Code.Text, ";")
    If strCommentEnd <> 0 = True Then
    Form1.Code.SelStart = strCommentStart - 1
    Form1.Code.SelLength = strCommentEnd - 1
    Form1.Code.SelColor = vbGreen
    Form1.Code.SelStart = strCommentEnd + 1
    Form1.Code.SelColor = vbBlack
    End If
End If
Next i
End Sub

Sub Compile()
Dim strCode(1 To 15)
Dim strPos(1 To 2)
Dim strSize(1 To 2)

strCode(1) = "messagebox"
strCode(2) = "input"
strCode(3) = "showform"
strCode(4) = "hideform"
strCode(5) = "label"
strCode(6) = "text"
strCode(7) = "picture"
strCode(8) = "caption"
strCode(9) = "button"
strCode(10) = "end"
strCode(11) = "text.hide"
strCode(12) = "picture.hide"
strCode(13) = "button.hide"
strCode(14) = "label.hide"
strCode(15) = "form"

strPos(1) = "posX"
strPos(2) = "posY"

strSize(1) = "width"
strSize(2) = "height"

For i = 1 To Len(Form1.Code.Text)
    For Z = 1 To 15
    
        strStart = InStr(i, Form1.Code.Text, strCode(Z))
        If strStart <> 0 = True Then
                    
            Select Case strCode(Z)              'Shows/Hides
                Case "text.hide"
                    Form2.Text1.Visible = False
                Case "picture.hide"
                    Form2.Picture1.Visible = False
                Case "button.hide"
                    Form2.Command1.Visible = False
                Case "label.hide"
                    Form2.Label1.Visible = False
                Case Else
            End Select                          'Shows/Hides

        strBracStart = InStr(strStart, Form1.Code.Text, "(")
        If strBracStart <> 0 = True Then
            strEnd = InStr(strBracStart, Form1.Code.Text, ");")
        If strEnd <> 0 = True Then
            strData = Mid(Form1.Code.Text, strBracStart + 1, strEnd - strBracStart - 1)
        
        Select Case strCode(Z)
        Case "messagebox"       'Message Box
        MsgBox (strData)
        Case "showform"         'Show Form
        Form2.Show
        Form2.Caption = strData
        Case "hideform"         'Hide Form
        Form2.Hide
        Case "end"              'End Program
        Unload Form2
        i = Len(Form1.Code.Text)
        Exit Sub
        End Select
        
                End If
                        End If
        
        'Start Of More Data Parse Coding
        'Postion Parse
        
        strPosBegin = InStr(strEnd, Form1.Code.Text, "{")
        
        'Start of Finding X value
        If strPosBegin <> 0 = True Then
            strNewPosX = InStr(strPosBegin, Form1.Code.Text, "posX=")
        If strNewPosX <> 0 = True Then
            strNewPosXend = InStr(strNewPosX, Form1.Code.Text, ";")
        If strNewPosXend <> 0 = True Then
        strDataPosX = Mid(Form1.Code.Text, strNewPosX + 5, strNewPosXend - strNewPosX - 5)
            End If
                End If
                
        'Start of Finding Y value
            strNewPosY = InStr(strPosBegin, Form1.Code.Text, "posY=")
        If strNewPosY <> 0 = True Then
            strNewPosYend = InStr(strNewPosY, Form1.Code.Text, ";")
        If strNewPosYend <> 0 = True Then
            strDataPosY = Mid(Form1.Code.Text, strNewPosY + 5, strNewPosYend - strNewPosY - 5)
    
            End If
                End If
                
        'Start of Finding Height Value
            strNewHeight = InStr(strPosBegin, Form1.Code.Text, "height=")
        If strNewHeight <> 0 = True Then
            strNewHeightend = InStr(strNewHeight, Form1.Code.Text, ";")
        If strNewHeightend <> 0 = True Then
            strHeight = Mid(Form1.Code.Text, strNewHeight + 7, strNewHeightend - strNewHeight - 7)
        End If
            End If
            
        'Start of Finding Width value
            strNewWidth = InStr(strPosBegin, Form1.Code.Text, "width=")
        If strNewWidth <> 0 = True Then
            strNewWidthEnd = InStr(strNewWidth, Form1.Code.Text, ";")
        If strNewWidthEnd <> 0 = True Then
            strWidth = Mid(Form1.Code.Text, strNewWidth + 6, strNewWidthEnd - strNewWidth - 6)
        End If
            End If
            
            strPosEnd = InStr(strPosBegin, Form1.Code.Text, "}")
        If strPosEnd <> 0 = True Then
            End If
                End If
        
        If strCode(Z) = "text" = True Then      'Textbox
        Form2.Text1.Text = strData
        Form2.Text1.Left = strDataPosX
        Form2.Text1.Top = strDataPosY
        Form2.Text1.Height = strHeight
        Form2.Text1.Width = strWidth
        Form2.Text1.Visible = True
        End If
        
        If strCode(Z) = "picture" = True Then   'Picture
        Form2.Picture1.Visible = True
        Form2.Picture1.Picture = LoadPicture(strData)
        Form2.Picture1.Left = strDataPosX
        Form2.Picture1.Top = strDataPosY
        Form2.Picture1.Height = strHeight
        Form2.Picture1.Width = strWidth
        End If
        
        If strCode(Z) = "button" = True Then    'Button
        Form2.Command1.Visible = True
        Form2.Command1.Caption = strData
        Form2.Command1.Left = strDataPosX
        Form2.Command1.Top = strDataPosY
        Form2.Command1.Height = strHeight
        Form2.Command1.Width = strWidth
        End If
        
        If strCode(Z) = "input" = True Then     'Input
        a = InputBox(strData)
        'direct data to objects in form2 aka text
        End If
        
        If strCode(Z) = "label" = True Then     'Label
        Form2.Label1.Visible = True
        Form2.Label1.Caption = strData
        Form2.Label1.Left = strDataPosX
        Form2.Label1.Top = strDataPosY
        Form2.Label1.Height = strHeight
        Form2.Label1.Width = strWidth
        End If
        
        'End Of It
        
        i = strEnd
        
        End If
    Next Z
    
    If Form1.stopcompile.Text = "1" = True Then
    i = Len(Form1.Code.Text)
    End If
Next i
End Sub
