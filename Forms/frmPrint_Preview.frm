VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPrint_Preview 
   Caption         =   "Print Preview"
   ClientHeight    =   5505
   ClientLeft      =   -120
   ClientTop       =   450
   ClientWidth     =   7875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   7875
   Begin RichTextLib.RichTextBox rtb 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7858
      _Version        =   393217
      TextRTF         =   $"frmPrint_Preview.frx":0000
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      Caption         =   "lblNotes"
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   8000
   End
End
Attribute VB_Name = "frmPrint_Preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'Insert notes here
lblNotes.Caption = ""
End Sub

Private Sub Form_Resize()
Shape1.width = Me.width
lblNotes.width = Me.ScaleWidth - (lblNotes.Left)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPrint_Preview = Nothing
End Sub
