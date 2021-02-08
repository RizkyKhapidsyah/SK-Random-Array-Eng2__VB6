VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rnd Array"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   2025
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstArr 
      Height          =   2595
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   3480
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar prgArr 
      Height          =   2595
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   4577
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Min             =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Label lblT 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblW 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const MaxNum As Integer = 10000
Const MinNum As Integer = 1 'Allways greather then 0

Private Sub cmdGo_Click()
  
   On Error Resume Next
   
   Dim RndArr(MinNum To MaxNum) As Integer
   
   Dim iCount As Integer
   Dim jCount As Integer
   
   Dim Limit As Integer
   Dim Exist As Boolean
   
   Dim Sum As Long
   Dim start As Single
   
   
   Limit = CInt(InputBox("Input number of elements ...", "Rnd", 10))
   
   
   
   'Check input ...
   If Limit > MaxNum Or Limit < MinNum Then Exit Sub
      
   lblT.Visible = True
   lblW.Visible = True
   lstArr.Clear
   
   prgArr.Visible = True
   prgArr.Max = Limit
   prgArr.Min = 1
   
   Sum = 0
   start = Timer
   For iCount = 1 To Limit
      Randomize
      Do
         Exist = False
         
         'Little optimization
         'Sum of first n integers = n*(n+1)/2
         'When iCount =Limit, simple subb n*(n+1)/2-Sum
         If iCount = Limit Then
            RndArr(iCount) = Limit * (Limit + 1) / 2 - Sum
         Else
            'Try next random value ...
            RndArr(iCount) = Int((Rnd * Limit) + 1)
         End If
         
         'Check if allready exist ...
         
         If iCount < Limit Then
            For jCount = 1 To iCount - 1
               If RndArr(iCount) = RndArr(jCount) Then Exist = True
            Next jCount
         End If
         
         DoEvents
         lblT.Caption = Round(Timer - start, 2)
         lblW.Caption = "Writing " & iCount

      Loop Until Not Exist
      
      lstArr.AddItem (RndArr(iCount))
      prgArr.Value = iCount
      Sum = Sum + RndArr(iCount)
   Next iCount
   
   prgArr.Visible = False
   lblW.Visible = False
   lblT.Visible = False
   
   MsgBox "Time elapsed  for " & Limit & " numbers" & vbNewLine & vbNewLine & " => " & lblT.Caption & " sec", , "Rnd"
   
End Sub

Private Sub Form_Load()

   Set Me.Icon = Nothing
   prgArr.Visible = False
   
End Sub

