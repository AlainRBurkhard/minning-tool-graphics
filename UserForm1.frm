VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "INPUT DATA"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6720
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Range("B6").Value = TextBox1.Value
Range("B8").Value = TextBox2.Value
Range("C21").Value = TextBox3.Value
Range("A23").Value = TextBox4.Value
Range("B23").Value = TextBox5.Value
Range("C23").Value = TextBox6.Value
Range("A24").Value = TextBox7.Value
Range("B24").Value = TextBox8.Value
Range("C24").Value = TextBox9.Value
Range("A25").Value = TextBox10.Value
Range("B25").Value = TextBox11.Value
Range("C25").Value = TextBox12.Value
Range("A27").Value = TextBox13.Value
Range("B27").Value = TextBox14.Value
Range("C27").Value = TextBox15.Value
Range("A28").Value = TextBox16.Value
Range("B28").Value = TextBox17.Value
Range("C28").Value = TextBox18.Value
Range("A29").Value = TextBox19.Value
Range("B29").Value = TextBox20.Value
Range("C29").Value = TextBox21.Value

UserForm1.Hide

End Sub


Private Sub Label1_Click()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub userform_activate()
Dim Rs As Worksheet

Label1.Caption = "Name of the sample: " & Range("b1").Value
        

End Sub
