VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMAPIOutXpress 
   Caption         =   "Send File Attachment via Outlook Express"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4905
   Icon            =   "frmMAPIOutXpress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin MSMAPI.MAPIMessages MAPIMessage1 
      Left            =   4200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      TabIndex        =   7
      Top             =   1320
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   885
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   285
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSendTo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1365
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   4200
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "View Outlook Express Transmission"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   765
      TabIndex        =   6
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label lblShortPath 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " <-- Select a file to attach "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1365
      TabIndex        =   2
      Top             =   600
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   405
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmMAPIOutXpress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***********************************************************************
'*       Project: Attach a file to Outlook Express with MAPI Component *
'*                               MSMAPI32.OCX                          *
'*       email: johnpc7@cox.net                                        *
'*       web:   http://members.cox.net/johnpc7/                        *
'***********************************************************************
Dim FileToAttach As String
Dim ActualNumberOfOccurrences As Integer
Dim AnsParseStr As String
Dim position As Integer
Dim NumCharInString As Integer
Dim checkit As Boolean
Dim RetVal As Variant

Private Sub cmdBrowse_Click()

'***********************************************************************
'Give the file selection window a title.
    cdlOpen.DialogTitle = "Pick A File To Attach To Outlook Express"
    'The file selection window will start in the
    'applications directory.
    cdlOpen.InitDir = App.Path
        
    'Show all files
    cdlOpen.Filter = "All Files (*.*)|*.*"
                       
    'Open the file selection window.
    cdlOpen.ShowOpen
    
    FileToAttach = cdlOpen.FileName
'***********************************************************************
'How many back slashes \ in our files path name
    position = 1

Do Until position = 0

   position = InStr(position + 1, FileToAttach, "\")
   ActualNumberOfOccurrences = ActualNumberOfOccurrences + 1
  
Loop
  
'Parse the string to get the part after the last backslash \
 AnsParseStr = ParseStr(FileToAttach, "\", ActualNumberOfOccurrences, ActualNumberOfOccurrences)
  

 '****************************************************************
    'Check to see if a file has been selected.
    Select Case FileToAttach
    
       Case vbNullString
            'Do not allow empty strings.
            checkit = False
            Exit Sub
            
        Case Else
            checkit = True
            lblShortPath.Visible = True
            lblShortPath = AnsParseStr
            
    End Select

End Sub

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Center the Form on the Screen
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub Option1_Click(Index As Integer)

If txtSendTo = "" Then
    RetVal = MsgBox("Send To not filled in", 16, "Send An Attachment in Outlook Express")
    Option1(0).Value = False
    Option1(1).Value = False
    txtSendTo.SetFocus
    Exit Sub
End If

If checkit = False Then
    RetVal = MsgBox("No File Selected", 16, "Send An Attachment in Outlook Express")
    Option1(0).Value = False
    Option1(1).Value = False
    cmdBrowse.SetFocus
    Exit Sub
End If

'Add the MAPI components (MSMAPI32.OCX)

MAPISession1.SignOn
MAPISession1.DownLoadMail = False
DoEvents
    MAPIMessage1.SessionID = MAPISession1.SessionID
    MAPIMessage1.Compose
    
    MAPIMessage1.RecipAddress = txtSendTo
    MAPIMessage1.ResolveName
    MAPIMessage1.MsgSubject = AnsParseStr
    MAPIMessage1.AttachmentPathName = FileToAttach
    MAPIMessage1.AttachmentName = AnsParseStr
'**********************************************************************
'**********************************************************************
'Supply your own information here
    MAPIMessage1.MsgNoteText = "Attached File:  - " & AnsParseStr _
        & vbCrLf & vbCrLf & vbCrLf & "Regards," & vbCrLf & vbCrLf _
        & "Your Name" & vbCrLf & vbCrLf & "YourEmail@someserver.com"
'**********************************************************************
'**********************************************************************
    Select Case Index 'Do you want to open Outlook Express?
    
        Case 0
            MAPIMessage1.Send True 'yes we want to see it
            
        Case Else
            MAPIMessage1.Send False 'no we don't need to see it
                          'just send it
    End Select
'**************************************************************************************

MAPISession1.SignOff

End Sub

Private Sub txtSendTo_GotFocus()
    txtSendTo.SelStart = 0
    txtSendTo.SelLength = Len(txtSendTo)
End Sub
Function ParseStr(ByVal Text, ByVal separator, ByVal start As Integer, _
ByVal toEnd As Integer) As String
    
    Dim i As Integer, Temp As String, result As String
    Dim ParseStrBegin As Integer, t As Integer, Count As Integer
    Dim ParseStrEnd As Integer, Found As Integer
    
    ParseStr = ""
    If Text = "" Then Exit Function
    If separator = "" Then Exit Function
    If Not (start > 0) Then start = 1
    If toEnd < start Then toEnd = start
    'Find first instance of the separator
     t = InStr(1, Text, separator)
    
    'If no occurence return original string and exit
    If t = 0 Then
    ParseStr = Text
    Exit Function
    End If


    'If first ParseStr, return left most data and exit

    If (start = 1) And (start = toEnd) Then
    If t = 1 Then
        ParseStr = ""
        Exit Function
    Else
        ParseStr = Left(Text, t - 1)
        Exit Function
    End If
    End If
    
    ParseStrBegin = 1
    For i = 1 To start - 1
       t = InStr(ParseStrBegin, Text, separator)
       If t = 0 Then Exit For
       ParseStrBegin = t + 1
       Next i
    
    ' If there is no separator exit function with "" result
    If t = 0 Then Exit Function
    
    'If only one ParseStr to return, find it and exit
    If start = toEnd Then
    t = InStr(ParseStrBegin, Text, separator)
    If t = 0 Then t = Len(Text) + 1
    result = Left(Text, t - 1)
    ParseStr = Right(result, t - ParseStrBegin)
    Exit Function
    End If
    
    'Find last ParseStr then exit

    ParseStrEnd = t + 1
    If start = 1 Then start = 2
    For i = start To toEnd

    t = InStr(ParseStrEnd, Text, separator)
    If t = 0 Then
        t = Len(Text) + 1
        Exit For
    End If
    ParseStrEnd = t + 1
    Next i
    If t = 0 Then t = Len(Text) + 1
    result = Left(Text, t - 1)
    ParseStr = Right(result, t - ParseStrBegin)

End Function

