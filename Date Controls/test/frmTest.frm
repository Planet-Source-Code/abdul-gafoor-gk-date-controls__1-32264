VERSION 5.00
Object = "*\A..\source\DateControls.vbp"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin DateControls.gkMonth gkMonth1 
      Height          =   2835
      Left            =   2190
      TabIndex        =   2
      Top             =   165
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   5001
      BeginProperty DayHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleFontSize   =   8.25
      DayHeaderFontSize=   8.25
      CurrentMonthFontSize=   8.25
      PreMonthFontSize=   8.25
      PostMonthFontSize=   8.25
      ActiveDayFontSize=   8.25
      ActiveDayFontItalic=   0   'False
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   615
      Width           =   1815
   End
   Begin DateControls.gkDatePicker DatePicker1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      CalendarActiveDayForeColor=   12582912
      MaxDate         =   38748
      MinDate         =   36161
      CustomFormat    =   "MMMM dd, yyyy"
      Value           =   37322
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Debug.Print DatePicker1.IsNull
End Sub
