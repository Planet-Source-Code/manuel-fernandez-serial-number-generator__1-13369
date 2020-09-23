VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Register ProgramName"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   795
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1395
      Width           =   3165
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Default         =   -1  'True
      Height          =   360
      Left            =   3000
      TabIndex        =   2
      Top             =   1005
      Width           =   960
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   810
      TabIndex        =   1
      Top             =   600
      Width           =   3120
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   795
      TabIndex        =   0
      Top             =   225
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "S/N"
      Height          =   225
      Left            =   90
      TabIndex        =   5
      Top             =   630
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "User"
      Height          =   195
      Left            =   75
      TabIndex        =   4
      Top             =   255
      Width           =   675
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim v As Collection
    Dim s As New Serial
    
    Set v = s.CheckSerial(Text1, Text2)
    
    If v("ValidKey") Then
        Text3 = "Key is OK, program: " & v("ProgramCode") & " v" & v("AppMajor") & "." & v("AppMinor") & vbCrLf & _
        "Key Date (m/y): " & v("KeyDate") & vbCrLf & _
        "Reserved Field: " & v("Reserved")
    Else
        Text3 = "Key is wrong"
    End If
    
    Set v = Nothing
End Sub
