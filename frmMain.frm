VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmMain 
   Caption         =   "File Compare"
   ClientHeight    =   7140
   ClientLeft      =   1530
   ClientTop       =   765
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   9360
   Begin RichTextLib.RichTextBox rtbView 
      Height          =   4695
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8281
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin VB.CommandButton cmdViewDifference 
         Caption         =   "&View Differences"
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdCompare 
         Caption         =   "&Compare Files"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtFileOne 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txtFileTwo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   3855
      End
      Begin VB.CommandButton cmdOpenOne 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   2
         Top             =   510
         Width           =   735
      End
      Begin VB.CommandButton cmdOpenTwo 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3960
         TabIndex        =   1
         Top             =   1110
         Width           =   735
      End
      Begin VB.Label lblAreDifferent 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   5400
         TabIndex        =   7
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "File One"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "File Two"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   750
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblWhichOne 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4650
      TabIndex        =   11
      Top             =   2040
      Width           =   75
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************
' K. Juryea 2001                        *
' Filecompare                           *
'****************************************
Option Explicit
Const TempFile = "C:\$$$diff"
Dim Filebox As String

Public Function CompareFiles(fileOne As String, fileTwo As String) As Boolean

    '***Declare All Variables***'
    Dim fileOneContent, fileTwoContent As String
    Dim temp, Status As String
    Dim FirstFile, NextFile, Line As Integer
    
    Call ClearAll 'Clear everything to start fresh
    FirstFile = FreeFile
    On Error GoTo Error:
    Open fileOne For Input As #FirstFile 'Users First Choice
    Do Until EOF(FirstFile)
        Line Input #FirstFile, temp
        fileOneContent = fileOneContent + temp 'Populate First Variable
    Loop
    Close #FirstFile


    NextFile = FreeFile
    Open fileTwo For Input As #NextFile 'Users Second Choice
    Do Until EOF(NextFile)
        Line Input #NextFile, temp
        fileTwoContent = fileTwoContent + temp 'Populate Second Variable
    Loop
    Close #NextFile


    If fileOneContent = fileTwoContent Then 'Compare the 2 Variables
        CompareFiles = True
        Status = "identical" 'Text string for Label (files are the Same)
    Else
        CompareFiles = False
        Status = "different" 'Text string for Label (files are Different)
        cmdViewDifference.Visible = True 'Since they are Different,Show View Button
    End If
    lblAreDifferent.Caption = "Files are " & Status 'Text string for Label (whole string)
    txtFileOne.SetFocus
    Exit Function
Error:
    Select Case Err.Number
        Case 75 'Error if no file is chosen in either text box
            MsgBox "Please Select Files To Compare", vbExclamation
            txtFileOne.SetFocus
    End Select
End Function

Public Sub fileopen()
    
    '***Declare All Variables***'
    Dim Message As String, Style As String, Title As String, Response As String, LastOpen As String
    Dim sString As String
    Dim lLength As Long
    Dim SaveDir As String
    Dim InputFile As String
    
    'Show File Open Dialog Box
    Filebox = ""
    CommonDialog1.filename = ""
    CommonDialog1.DialogTitle = "Select File To Convert"
    CommonDialog1.Filter = "All Files  *.*"
    CommonDialog1.ShowOpen
    InputFile = CommonDialog1.filename
    Filebox = InputFile
    If InputFile = "" Then
        MsgBox "Open File Action Canceled", vbInformation + vbOKOnly, "LogSCAN Convert"
    End If
End Sub

Private Sub cmdCompare_Click()
    'Compare the 2 Files
    Call CompareFiles(txtFileOne.Text, txtFileTwo.Text)
End Sub

Private Sub cmdOpenOne_Click()
    Call ClearAll 'Clear Everything
    Call fileopen 'Show File Open Dialog
    txtFileOne = Filebox
End Sub

Private Sub cmdOpenTwo_Click()
    Call ClearAll 'Clear Everything
    Call fileopen 'Show File Open Dialog
    txtFileTwo = Filebox
End Sub

Private Sub cmdViewDifference_Click()
    'Function to View Differences (showing lines in the second file)
    If exists(TempFile) Then
        Kill (TempFile) 'Kill old file if it exists
    End If
    Call WriteDiffer(txtFileOne.Text, txtFileTwo.Text, TempFile)
    lblWhichOne.Visible = True
    lblWhichOne.Caption = "The lines below are different in " & txtFileTwo.Text
    rtbView.LoadFile TempFile
End Sub

Private Sub ClearAll()
    cmdViewDifference.Visible = False
    rtbView.Text = ""
    lblAreDifferent.Caption = ""
    lblWhichOne.Visible = False
End Sub

Private Sub Form_Load()
    Call ClearAll 'Make sure everything is cleared out
End Sub
