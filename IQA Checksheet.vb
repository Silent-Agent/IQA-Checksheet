Option Explicit
Function ValidateForm() As Boolean
    txtpartcode.BackColor = vbWhite
    txtpartname.BackColor = vbWhite
    txtpartno.BackColor = vbWhite
    
    ValidateForm = True
    
    If Trim(txtpartcode.Value) = "" Then
        MsgBox "Box can't be left blank.", vbOKOnly + vbInformation, "Part Code"
        txtpartcode.BackColor = vbRed
        txtpartcode.Activate
        ValidateForm = False
        
    ElseIf optProd.Value = False And optqc.Value = False Then
    
        MsgBox "Please select department.", vbOKOnly + vbInformation, "department"
        ValidateForm = False
        
    ElseIf Trim(txtpartname.Value) = "" Then
        MsgBox "Box can't be left blank", vbOKOnly + vbInformation, "Part Name"
        txtpartname.BackColor = vbRed
        txtpartname.Activate
        ValidateForm = False
        
  
    ElseIf Trim(txtpartno.Value) = "" Then
        MsgBox "Box can't be left blank", vbOKOnly + vbInformation, "Part No."
        txtpartno.BackColor = vbRed
        txtpartno.Activate
        ValidateForm = False
    
    End If
    
End Function

Private Sub cmddelete_Click()

  ' Unprotect the sheet to enable changes to protected cells
    Sheet3.Unprotect Password:="" 'replace "password" with your sheet password, if any
    'Dim i As Integer
    Dim i As Long
    Dim lastRow As Long
    Dim iDelete As VbMsgBoxResult
    
    lastRow = Cells.SpecialCells(xlLastCell).Row
    
    'For i = 1 To Range("A1048576").End(xlUp).Row - 1
    For i = lastRow To 2 Step -1
        'If ListBox.Selected(i) Then
            'Dim iDelete As VbMsgBoxResult
           'iDelete = MsgBox("confirm if you want to delete", vbQuestion + vbYesNo, "Data Entry Form")
            'If iDelete = vbYes Then
            'Rows(i + 1).Select
            'Selection.Delete
            'End If
        'End If
        If ListBox.ListIndex = i - 2 Then
            iDelete = MsgBox("Are you sure you want to delete this row?", vbQuestion + vbYesNo, "Data Entry Form")
            If iDelete = vbYes Then
                Rows(i).Delete
            End If
        End If
        
    Next i
       ' Reprotect the sheet to prevent accidental changes
    Sheet3.Protect Password:="" 'replace "password" with your sheet password, if any
    
End Sub


Private Sub cmdexit_Click()
'Dim iExit As VbMsgBoxResult
    'iExit = MsgBox("confirm if you want to exit", vbQuestion + vbYesNo, "Data Entry Form")
    'If iExit = vbYes Then
    Unload Me
    'End If

End Sub

Private Sub cmdreset_Click()

txtitem.Text = ""
txtextprovider.Text = ""
txtlotno.Text = ""
txtreceivingqty.Text = ""
txtdono.Text = ""
optattach.Value = False
optnotattach.Value = False
optappearanceok.Value = False
optappearancenotgood.Value = False
txtsamplingqty.Text = ""
txtaql.Text = ""
txtrejectqty.Text = ""
txtdefect.Text = ""
optdefectok.Value = False
optdefectok.Value = False
txtactiontaken.Text = ""
txtremark.Text = ""
txtstopcardno.Text = ""

End Sub

Private Sub cmdsave_Click()

    Dim wks As Worksheet
    Dim addnew As Range
    Set wks = Sheet3
    
     ' Unprotect the sheet to enable changes to protected cells
    wks.Unprotect Password:="" 'replace "password" with your sheet password, if any

     Set addnew = wks.Range("A1048576").End(xlUp).Offset(1, 0)

    addnew.Offset(0, 0).Value = Format(Now, "mm/dd/yyyy hh:mm:ss")
    addnew.Offset(0, 1).Value = txtitem.Text
    addnew.Offset(0, 2).Value = txtextprovider.Text
    addnew.Offset(0, 3).Value = txtlotno.Text
    addnew.Offset(0, 4).Value = txtreceivingqty.Text
    addnew.Offset(0, 5).Value = txtdono.Text
    addnew.Offset(0, 6).Value = IIf(optattach.Value = True, "Attach", "Not Attach")
    addnew.Offset(0, 7).Value = IIf(optappearanceok.Value = True, "OK", "Not Good")
    addnew.Offset(0, 8).Value = txtsamplingqty.Text
    addnew.Offset(0, 9).Value = txtaql.Text
    addnew.Offset(0, 10).Value = txtrejectqty.Text
    addnew.Offset(0, 11).Value = txtdefect.Text
    addnew.Offset(0, 12).Value = IIf(optdefectok.Value = True, "OK", "Not Good")
    addnew.Offset(0, 13).Value = txtactiontaken.Text
    addnew.Offset(0, 14).Value = txtremark.Text
    addnew.Offset(0, 15).Value = txtstopcardno.Text
    
       ' Reprotect the sheet to prevent accidental changes
    wks.Protect Password:="" 'replace "password" with your sheet password, if any
    ListBox.ColumnCount = 5
    ListBox.RowSource = "A2:D1048576"

    Call cmdreset_Click

End Sub

