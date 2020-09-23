VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Serial Number Generator"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSerial 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1035
      TabIndex        =   6
      Top             =   1860
      Width           =   3330
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Default         =   -1  'True
      Height          =   390
      Left            =   2175
      TabIndex        =   5
      Top             =   1290
      Width           =   1260
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   4
      Left            =   1050
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   1365
      Width           =   945
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   3
      Left            =   1515
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   1020
      Width           =   495
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   2
      Left            =   1050
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1020
      Width           =   450
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   1
      Left            =   1035
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   645
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   0
      Left            =   1035
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   270
      Width           =   2475
   End
   Begin VB.Label Label5 
      Caption         =   "Serial"
      Height          =   300
      Left            =   195
      TabIndex        =   11
      Top             =   1860
      Width           =   810
   End
   Begin VB.Label Label4 
      Caption         =   "Reserved"
      Height          =   240
      Left            =   135
      TabIndex        =   10
      Top             =   1410
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   "Version"
      Height          =   255
      Left            =   210
      TabIndex        =   9
      Top             =   1035
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Program:"
      Height          =   255
      Left            =   225
      TabIndex        =   8
      Top             =   645
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "User"
      Height          =   255
      Left            =   225
      TabIndex        =   7
      Top             =   285
      Width           =   795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    On Error GoTo EH
    
    Dim s As New Serial
    If (Val(txtData(1)) > 255) Or (Val(txtData(2)) > 15) Or (Val(txtData(3)) > 15) Then
        Err.Raise vbObjectError + 1
    End If
    txtSerial = s.GenerateSerial(txtData(0), Val(txtData(1)), Val(txtData(2)), Val(txtData(3)), Val(txtData(4)))
    Form2.Text1 = txtData(0)
    Form2.Text2 = txtSerial
    
    Exit Sub

EH:
    MsgBox "Invalid Data, program E [0,255], Ver* E [0,15], Reserved E [0,255]"
End Sub

Private Sub Form_Load()
    
    
    txtData(0) = "John Doe"
    txtData(1) = "1"
    txtData(2) = "1"
    txtData(3) = "0"
    txtData(4) = "8"
    
    Left = 500
    Top = 500
    Form2.Left = 600 + Me.Width
    Form2.Top = Top
    Form2.Show
End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index > 0 Then
        'Numbers only!
        If KeyAscii > Asc(" ") Then 'Allow special codes
            If (KeyAscii < Asc("0")) Or (KeyAscii > Asc("9")) Then
                Beep
                KeyAscii = 0
            End If
        End If
    End If
End Sub
