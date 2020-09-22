VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   2430
      TabIndex        =   0
      Top             =   780
      Width           =   1290
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   180
      Picture         =   "frmAbout.frx":0000
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Abdul Gafoor.GK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   1305
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   180
      Width           =   1440
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "gafoorgk@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   3
      Left            =   1020
      MouseIcon       =   "frmAbout.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   420
      Width           =   1950
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "by"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   1020
      TabIndex        =   1
      Top             =   180
      Width           =   165
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Private Const strEMail As String = "gafoorgk@hotmail.com"

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub lbl_Click(Index As Integer)
    If (Index = 3) Then
        Call ShellExecute(0&, vbNullString, _
                          "mailto:" & strEMail, vbNullString, _
                          "C:\", SW_SHOWNORMAL)
    End If
End Sub

Public Sub ActivateForm(strID As String)
    lbl(3).Caption = strEMail
    
    If (strID = "DATEPICKER") Then
        Me.Caption = "About Date Picker"
        Me.Icon = LoadResPicture("DATEPICKER", vbResIcon)
    Else
        Me.Caption = "About Month Control"
        Me.Icon = LoadResPicture("MONTH", vbResIcon)
    End If
    Set img.Picture = Me.Icon
    
    Me.Show vbModal
End Sub
