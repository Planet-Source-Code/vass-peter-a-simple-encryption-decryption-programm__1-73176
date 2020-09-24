VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Krypton(Encryption version)"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   0
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt!"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2520
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Dim szoveg As String
Dim tmp As String
Dim s As String
Dim string2 As String
Dim string1 As String

Private Sub Command1_Click()
string1 = Text1
string2 = Text2
'**********ERROR**********'
If string1 = "" Then
MsgBox "This field cann not be blank!", vbCritical, "Error"
Exit Sub
End If

If Text2 = "" Then
MsgBox "This field cann not be blank! Please enter a valid encryption key(only numbers)!", vbCritical, "Error"
Exit Sub
End If

For f = 1 To Len(string2)
ascii = Asc(Mid(string2, f, 1))
If ascii < 48 Or ascii > 58 Then
MsgBox "Please enter a valid encryption key(only numbers)!", vbCritical, "Error"
Text2 = ""
Exit Sub
End If
Next f
'**********ENCRYPTION**********'
Open App.Path & "\encrypted.txt" For Output As #1
For i = 1 To Len(string1)
tmp = Asc(Mid(string1, i, 1)) * Text2
Print #1, tmp
Next i
Close #1
MsgBox "Encryption was succesful!!" & "Filepath:" & App.Path & "\encrypted.txt", vbInformation, "SUCCES!!"
End Sub
Private Sub Form_Load()
MsgBox "Created by Vass Peter whit vb 6.0!", vbInformation, ""
End Sub
