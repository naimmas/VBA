VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   ClientHeight    =   9690.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13170
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()
ListBox1.Clear

If (OptionButton2.Value = True And OptionButton3.Value = False) Then
q = 0
aaa = CLng(Label4.Caption) + 1

For i = 2 To aaa
If InStr(CStr(Sheets("1FullData").Cells(i, 1).Value), CStr(TextBox1.Value)) Then
With ListBox1
.AddItem
.List(q, 0) = Sheets("1FullData").Cells(i, 1).Value
.List(q, 1) = Sheets("1FullData").Cells(i, 2).Value
.List(q, 2) = Sheets("1FullData").Cells(i, 3).Value
.List(q, 3) = Sheets("1FullData").Cells(i, 4).Value


End With
q = q + 1
End If
Next i

End If
'----------------------------------------------------

If (OptionButton1.Value = True And OptionButton3.Value = False) Then
q = 0
aaa = CLng(Label4.Caption) + 1

For i = 2 To aaa
If InStr(CStr(Sheets("1FullData").Cells(i, 2).Value), CStr(TextBox1.Value)) Then
With ListBox1
.AddItem
.List(q, 0) = Sheets("1FullData").Cells(i, 1).Value
.List(q, 1) = Sheets("1FullData").Cells(i, 2).Value
.List(q, 2) = Sheets("1FullData").Cells(i, 3).Value
.List(q, 3) = Sheets("1FullData").Cells(i, 4).Value


End With
q = q + 1
End If
Next i

End If
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------

If OptionButton2.Value = True And OptionButton3.Value = True Then
If TextBox6.Value = "" Then MsgBox ("íÑÌì ßÊÇÈÉ ÚÏÏ ÇáÇÎÊÈÇÑÇÊ ÇáÊãåíÏíÉ"): Exit Sub

q = 0
aaa = CLng(Label4.Caption) + 1
ggg = CLng(TextBox6.Value)

For y = 1 To ggg
For i = 2 To aaa
If IsNumeric(Sheets("1FullData").Cells(i, 4).Value) Then
mmm = CLng(Sheets("1FullData").Cells(i, 4).Value)
If InStr(CStr(Sheets("1FullData").Cells(i, 1).Value), CStr(TextBox1.Value)) And mmm = y Then
With ListBox1
.AddItem
.List(q, 0) = Sheets("1FullData").Cells(i, 1).Value
.List(q, 1) = Sheets("1FullData").Cells(i, 2).Value
.List(q, 2) = Sheets("1FullData").Cells(i, 3).Value
.List(q, 3) = Sheets("1FullData").Cells(i, 4).Value
End With
q = q + 1

GoTo ppp
End If
End If



Next i
ppp:
Next y
',,,,,,,,,,,,,,,,,,,,,,
For v = 2 To aaa

If InStr(CStr(Sheets("1FullData").Cells(v, 1).Value), CStr(TextBox1.Value)) And Sheets("1FullData").Cells(v, 4).Value = CStr("A") Then
With ListBox1
.AddItem
.List(q, 0) = Sheets("1FullData").Cells(v, 1).Value
.List(q, 1) = Sheets("1FullData").Cells(v, 2).Value
.List(q, 2) = Sheets("1FullData").Cells(v, 3).Value
.List(q, 3) = Sheets("1FullData").Cells(v, 4).Value


End With
q = q + 1
GoTo oop1
End If
Next v
',,,,,,,,,,,,,,,,,,,,,,,,,
oop1:
For m = 2 To aaa

If InStr(CStr(Sheets("1FullData").Cells(m, 1).Value), CStr(TextBox1.Value)) And Sheets("1FullData").Cells(m, 4).Value = CStr("B") Then
With ListBox1
.AddItem
.List(q, 0) = Sheets("1FullData").Cells(m, 1).Value
.List(q, 1) = Sheets("1FullData").Cells(m, 2).Value
.List(q, 2) = Sheets("1FullData").Cells(m, 3).Value
.List(q, 3) = Sheets("1FullData").Cells(m, 4).Value


End With
q = q + 1
GoTo oop2
End If

Next m
',,,,,,,,,,,,,,,,,,,,,,,,,
oop2:

For x = 2 To aaa

If InStr(CStr(Sheets("1FullData").Cells(x, 1).Value), CStr(TextBox1.Value)) And Sheets("1FullData").Cells(x, 4).Value = CStr("C") Then
With ListBox1
.AddItem
.List(q, 0) = Sheets("1FullData").Cells(x, 1).Value
.List(q, 1) = Sheets("1FullData").Cells(x, 2).Value
.List(q, 2) = Sheets("1FullData").Cells(x, 3).Value
.List(q, 3) = Sheets("1FullData").Cells(x, 4).Value


End With
q = q + 1
GoTo oop3
End If
Next x
',,,,,,,,,,,,,,,,,,,,,,,,,
oop3:

End If



'>>>>>>>>
'>>>>>>>>>
qwe = ListBox1.ListCount - 1
For u = 0 To qwe
If IsNumeric(ListBox1.List(u, 3)) Then dfg = CLng(dfg) + CLng(ListBox1.List(u, 2)): kkk = kkk + 1
If ListBox1.List(u, 3) = "A" Then dfgg = CLng(dfgg) + CLng(ListBox1.List(u, 2))
If ListBox1.List(u, 3) = "B" Then dfggg = CLng(dfggg) + CLng(ListBox1.List(u, 2))
If ListBox1.List(u, 3) = "C" Then dfgggg = CLng(dfgggg) + CLng(ListBox1.List(u, 2))

Next u

If TextBox6.Value = "" Then bnbn = kkk Else bnbn = CLng(TextBox6.Value)
If CLng(bnbn) = 0 Then bnbn = 1
Label6.Caption = Round(dfg / bnbn, 2)
Label7.Caption = dfgg
Label8.Caption = dfggg
Label9.Caption = dfgggg

qlabel6 = dfg / bnbn * CLng(TextBox2.Value) / 100
qlabel7 = dfgg * CLng(TextBox3.Value) / 100
qLabel8 = dfggg * CLng(TextBox4.Value) / 100
qlabel9 = dfgggg * CLng(TextBox5.Value) / 100

Label18.Caption = Round(qlabel6 + qlabel7 + qLabel8 + qlabel9, 2)

End Sub

Private Sub CommandButton1_Click()
If OptionButton2.Value = True And OptionButton3.Value = True Then
If TextBox6.Value = "" Then MsgBox ("íÑÌì ßÊÇÈÉ ÚÏÏ ÇáÇÎÊÈÇÑÇÊ ÇáÊãåíÏíÉ"): Exit Sub
q = 0
aaa = CLng(Label4.Caption) + 1
ggg = CLng(TextBox6.Value)
lll = CLng(Label5.Caption) + 1

For ll = 2 To lll



For y = 1 To ggg
For i = 2 To aaa
If IsNumeric(Sheets("1FullData").Cells(i, 4).Value) Then
mmm = CLng(Sheets("1FullData").Cells(i, 4).Value)
If InStr(CStr(Sheets("1FullData").Cells(i, 1).Value), CStr(Sheets("1FullData").Cells(ll, 7).Value)) And mmm = y Then
qaz = qaz + CDbl(Sheets("1FullData").Cells(i, 3).Value)


GoTo ppp
End If
End If



Next i
ppp:
Next y



Sheets("1FullData").Cells(ll, 9) = Round(qaz / CLng(TextBox6.Value), 2)



',,,,,,,,,,,,,,,,,,,,,,
For v = 2 To aaa

If InStr(CStr(Sheets("1FullData").Cells(v, 1).Value), CStr(Sheets("1FullData").Cells(ll, 7).Value)) And Sheets("1FullData").Cells(v, 4).Value = CStr("A") Then
qaza = qaza + CDbl(Sheets("1FullData").Cells(v, 3).Value)

GoTo oop1
End If
Next v

oop1:

Sheets("1FullData").Cells(ll, 10) = Round(qaza, 2)
',,,,,,,,,,,,,,,,,,,,,,,,,

For m = 2 To aaa

If InStr(CStr(Sheets("1FullData").Cells(m, 1).Value), CStr(Sheets("1FullData").Cells(ll, 7).Value)) And Sheets("1FullData").Cells(m, 4).Value = CStr("B") Then
qazb = qazb + CDbl(Sheets("1FullData").Cells(m, 3).Value)


GoTo oop2
End If

Next m

oop2:
Sheets("1FullData").Cells(ll, 11) = Round(qazb, 2)
',,,,,,,,,,,,,,,,,,,,,,,,,
For x = 2 To aaa

If InStr(CStr(Sheets("1FullData").Cells(x, 1).Value), CStr(Sheets("1FullData").Cells(ll, 7).Value)) And Sheets("1FullData").Cells(x, 4).Value = CStr("C") Then
qazc = qazc + CDbl(Sheets("1FullData").Cells(x, 3).Value)

GoTo oop3
End If
Next x
',,,,,,,,,,,,,,,,,,,,,,,,,
oop3:
Sheets("1FullData").Cells(ll, 12) = Round(qazc, 2)


qlabel6 = qaz / CLng(TextBox6.Value) * CLng(TextBox2.Value) / 100
qlabel7 = qaza * CLng(TextBox3.Value) / 100
qLabel8 = qazb * CLng(TextBox4.Value) / 100
qlabel9 = qazc * CLng(TextBox5.Value) / 100

Sheets("1FullData").Cells(ll, 13) = Round(qlabel6 + qlabel7 + qLabel8 + qlabel9, 2)

qaz = 0
qaza = 0
qazb = 0
qazc = 0
Label25.Caption = CLng(ll)
DoEvents

Next ll
'>>>>>>>>
'>>>>>>>>>
MsgBox ("Finished")
End If


End Sub

Private Sub Label12_Click()

End Sub

Private Sub OptionButton3_Change()
OptionButton2.Value = True
OptionButton1.Enabled = False
TextBox6.Value = ""
TextBox6.Enabled = True
Label24.Enabled = True
CommandButton1.Enabled = True
End Sub

Private Sub OptionButton3_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()
If IsNumeric(TextBox2.Value) Then a1 = CLng(TextBox2.Value)
If IsNumeric(TextBox3.Value) Then a2 = CLng(TextBox3.Value)
If IsNumeric(TextBox4.Value) Then a3 = CLng(TextBox4.Value)
If IsNumeric(TextBox5.Value) Then a4 = CLng(TextBox5.Value)


Label20 = a1 + a2 + a3 + a4
End Sub

Private Sub TextBox3_Change()
If IsNumeric(TextBox2.Value) Then a1 = CLng(TextBox2.Value)
If IsNumeric(TextBox3.Value) Then a2 = CLng(TextBox3.Value)
If IsNumeric(TextBox4.Value) Then a3 = CLng(TextBox4.Value)
If IsNumeric(TextBox5.Value) Then a4 = CLng(TextBox5.Value)


Label20 = a1 + a2 + a3 + a4
End Sub

Private Sub TextBox4_Change()
If IsNumeric(TextBox2.Value) Then a1 = CLng(TextBox2.Value)
If IsNumeric(TextBox3.Value) Then a2 = CLng(TextBox3.Value)
If IsNumeric(TextBox4.Value) Then a3 = CLng(TextBox4.Value)
If IsNumeric(TextBox5.Value) Then a4 = CLng(TextBox5.Value)


Label20 = a1 + a2 + a3 + a4
End Sub

Private Sub TextBox5_Change()
If IsNumeric(TextBox2.Value) Then a1 = CLng(TextBox2.Value)
If IsNumeric(TextBox3.Value) Then a2 = CLng(TextBox3.Value)
If IsNumeric(TextBox4.Value) Then a3 = CLng(TextBox4.Value)
If IsNumeric(TextBox5.Value) Then a4 = CLng(TextBox5.Value)


Label20 = a1 + a2 + a3 + a4
End Sub

Private Sub UserForm_Click()
OptionButton3.Value = False
OptionButton1.Enabled = True
TextBox6.Value = ""
TextBox6.Enabled = False
Label24.Enabled = False
CommandButton1.Enabled = False

End Sub

Private Sub UserForm_Initialize()
Label4.Caption = Application.WorksheetFunction.CountA(Sheets("1FullData").Range("a2:a1000000"))
Label5.Caption = Application.WorksheetFunction.CountA(Sheets("1FullData").Range("g2:g1000000"))
End Sub
