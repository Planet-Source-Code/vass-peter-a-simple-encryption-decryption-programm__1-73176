VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Krypton"
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   1620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Decrypt!"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ascii As Integer
Dim tmp As String
Dim string2 As String
Dim string1 As String

Private Sub Command1_Click()
string2 = Text1
If Text1 = "" Then
MsgBox "This field cann not be blank!Please enter a valid encryption key(only numbers)!", vbCritical, "Error"
Exit Sub
End If
For f = 1 To Len(Text1)
ascii = Asc(Mid(string2, f, 1))
If ascii < 48 Or ascii > 58 Then
MsgBox "Please enter a valid encryption key(only numbers)", vbCritical, "Error"
Exit Sub
End If
Next f
CommonDialog1.ShowOpen
Open CommonDialog1.FileTitle For Input As #1
While EOF(1) = 0
Line Input #1, tmp
string1 = string1 & Chr(tmp / Text1)
Wend
MsgBox string1
Close #1
Open App.Path & "\decrypted.txt" For Output As #1
Print #1, string1
Close #1
string1 = ""
szoveg = "Decryption was succesful!!" & "Filepath: " & App.Path & "\decrypted.txt"
MsgBox szoveg, vbInformation, "SUCCES!!"
End Sub

Private Sub Form_Load()
MsgBox "Created by Vass Péter whit vb 6.0!", vbInformation, " "
End Sub
