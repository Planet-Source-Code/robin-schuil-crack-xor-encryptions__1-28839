VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "XOR Cracker"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Encrypt"
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   5175
      Begin VB.TextBox Text2 
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Plaintext"
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5175
      Begin VB.TextBox Text1 
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Text            =   "insecure"
         Top             =   2160
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Crack"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Result:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label lblPassword 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "................"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5760
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const strMessage = "A substitution cipher is an extremely simple example of " & vbCrLf _
                & "conventional cryptography. A substitution cipher subsitutes " & vbCrLf _
                & "one piece of information for another. This is most frequently " & vbCrLf _
                & "done by offsetting letters of the alphabet using XOR. " & vbCrLf _
                & "However, this is not a safe method to encrypt your data. " & vbCrLf _
                & "This project demonstrates how to crack a XOR encrypted text " & vbCrLf _
                & "file. It retrieves the password for you which was used to " & vbCrLf _
                & "encrypt the message. " & vbCrLf _
                & "It can crack *most* of the messages, as long as the message " & vbCrLf _
                & "is at least 4 or 5 times as long as the password."

Private strEncrypted As String

Private Sub Command1_Click()
    
    Command1.Enabled = False
    XorCrack strEncrypted
    Command1.Enabled = True
    
    MsgBox XorCrypt(strEncrypted, lblPassword.Caption), vbInformation, "Decrypted message"
    
End Sub

Private Sub Command2_Click()

    strEncrypted = XorCrypt(Text1.Text, Text3.Text)
    
    Text2.Text = strEncrypted
    
End Sub

Private Sub Form_Load()
    
    Text1.Text = strMessage
    Text2.Text = ""
    Text3.Text = "insecure"

End Sub
