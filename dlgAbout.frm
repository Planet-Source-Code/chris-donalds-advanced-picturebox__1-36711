VERSION 5.00
Begin VB.Form dlgAbout 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2940
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5715
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "dlgAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5715
   StartUpPosition =   1  'CenterOwner
   Tag             =   "About SalesDesk"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   90
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   2490
      Width           =   1260
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1320
      Left            =   60
      Picture         =   "dlgAbout.frx":058A
      Stretch         =   -1  'True
      Top             =   60
      Width           =   1320
   End
   Begin VB.Label lblNote 
      Caption         =   "Web"
      Height          =   315
      Left            =   1680
      MouseIcon       =   "dlgAbout.frx":063E
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1830
      Width           =   3915
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2610
      X2              =   7080
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OMIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   1740
      TabIndex        =   5
      Tag             =   "CompanyProduct"
      Top             =   60
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advanced PictureBox"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   1890
      TabIndex        =   4
      Tag             =   "CompanyProduct"
      Top             =   120
      Width           =   3405
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Copyright"
      Height          =   840
      Left            =   1680
      TabIndex        =   3
      Top             =   870
      Width           =   3915
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   1650
      X2              =   5610
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   1680
      X2              =   5580
      Y1              =   2205
      Y2              =   2205
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "TransPicturebox ActiveX Control, version 1.0"
      Height          =   195
      Left            =   1920
      TabIndex        =   2
      Tag             =   "Version"
      Top             =   540
      Width           =   3165
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warning: ..."
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   1680
      TabIndex        =   1
      Tag             =   "Warning: ..."
      Top             =   2310
      Width           =   3915
   End
End
Attribute VB_Name = "dlgAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Private Const SW_SHOW = 5       ' Displays Window in its current size and position
Private Const SW_SHOWNORMAL = 1 ' Restores Window if Minimized or Maximized

Private Sub Form_Load()
    lblCopyright = App.LegalCopyright + Chr$(13) + App.LegalTrademarks
    lblDisclaimer = "* Warning: No part of this software product may be reproduced, retransmitted or modified in any form."
    lblNote = "Visit us on the Web at omitinc.com"
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub openDefBrowser(url As String)
    Dim filename As String
    Dim Dummy As String
    Dim BrowserExec As String * 255
    Dim RetVal As Long
    Dim FileNumber As Integer
    
    ' First, create a known, temporary HTML file
    BrowserExec = Space(255)
    filename = "C:\temphtm.htm"
    FileNumber = FreeFile                    ' Get unused file number
    Open filename For Output As #FileNumber  ' Create temp HTML file
        Write #FileNumber, "<HTML> <\HTML>"  ' Output text
    Close #FileNumber                        ' Close file
    
    ' Then find the application associated with it
    RetVal = FindExecutable(filename, Dummy, BrowserExec)
    BrowserExec = Trim(BrowserExec)
    
    ' If an application is found, launch it!
    If RetVal <= 32 Or IsEmpty(BrowserExec) Then ' Error
        MsgBox "Could not find associated Browser", vbExclamation, "Browser Not Found"
    Else
        RetVal = ShellExecute(hWnd, "open", BrowserExec, url, Dummy, SW_SHOWNORMAL)
        If RetVal <= 32 Then        ' Error
            MsgBox "Web Page not Opened", vbExclamation, App.Title
        End If
    End If
    Kill filename                   ' delete temp HTML file
End Sub

Private Sub lblNote_Click()
    openDefBrowser ("http://www.omitinc.com")
End Sub
