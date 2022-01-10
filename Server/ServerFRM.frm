VERSION 5.00
Begin VB.Form ServerFRM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BackDoor   [Server]"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   8370
      PasswordChar    =   "*"
      TabIndex        =   11
      Text            =   "1019"
      Top             =   1305
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Timer tmrchk 
      Interval        =   1000
      Left            =   7995
      Top             =   3585
   End
   Begin VB.TextBox txtport 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8370
      TabIndex        =   8
      Text            =   "1019"
      Top             =   1005
      Width           =   1530
   End
   Begin VB.CommandButton cmdconn 
      Caption         =   "Listen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7965
      TabIndex        =   5
      Top             =   1605
      Width           =   1935
   End
   Begin VB.CommandButton senddataCMD 
      Caption         =   "&Send"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Left            =   6405
      TabIndex        =   4
      Top             =   5610
      Width           =   975
   End
   Begin VB.TextBox dataTXT 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      MaxLength       =   100
      TabIndex        =   0
      Top             =   5670
      Width           =   6210
   End
   Begin VB.TextBox MainTXT 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5400
      Left            =   75
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "ServerFRM.frx":0000
      Top             =   105
      Width           =   7605
   End
   Begin VB.Label lbltimecon 
      AutoSize        =   -1  'True
      Caption         =   "Time : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7815
      TabIndex        =   14
      Top             =   2625
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pass"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7980
      TabIndex        =   13
      Top             =   1350
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lblPort 
      AutoSize        =   -1  'True
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7980
      TabIndex        =   12
      Top             =   990
      Width           =   300
   End
   Begin VB.Label lblstdev 
      AutoSize        =   -1  'True
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   6000
      Width           =   450
   End
   Begin VB.Label lblremip 
      AutoSize        =   -1  'True
      Caption         =   "Remote IP : 127.0.0.1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7815
      TabIndex        =   9
      Top             =   2340
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label YouripLBL 
      AutoSize        =   -1  'True
      Caption         =   "Your IP : 127.0.0.1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7830
      TabIndex        =   7
      Top             =   225
      Width           =   1620
   End
   Begin VB.Label lblstat 
      AutoSize        =   -1  'True
      Caption         =   "Close..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   270
      Left            =   7890
      TabIndex        =   6
      Top             =   510
      Width           =   840
   End
   Begin VB.Label nickCLIENT 
      Height          =   135
      Left            =   585
      TabIndex        =   3
      Top             =   5340
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label nickSERVER 
      Height          =   135
      Left            =   345
      TabIndex        =   2
      Top             =   5340
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "ServerFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
Dim WithEvents Winsock As CSocketMaster
Attribute Winsock.VB_VarHelpID = -1

'======================Displays info if disconnected========================
Private Sub Winsock_Close()
    MainTXT.SelText = "¨'°*·º·*°'¨ Disconnected to Client ¨'°*·º·*°'¨" & vbCrLf 'displays text if u get disconnected from the client"
    senddataCMD.Enabled = False 'this wont let you send text anymore since client is disconnected
End Sub

Private Sub cmdconn_Click()
If cmdconn.Caption = "Listen" Then
    Winsock.CloseSck 'closes any previous connections
    Winsock.LocalPort = CLng(txtport.Text)  '187 is the port
    Winsock.Listen 'listens to see if client wants to connect
  cmdconn.Caption = "Disconnect"
dataTXT.SetFocus
Else
  Winsock.CloseSck
End If

End Sub

Private Sub Form_Load()
Set Winsock = New CSocketMaster
YouripLBL.Caption = "Your IP:  " & Winsock.LocalIP
lblstat.ForeColor = RGB(180, 0, 0)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock.CloseSck
End Sub

Private Sub tmrchk_Timer()
lblstdev.Caption = "State:" & Winsock.State
If Winsock.State = sckListening Then lblstat = "Ready...": lblstat.ForeColor = RGB(0, 0, 180)
If Winsock.State = sckClosed Then
  lblstat = "Close..."
  lblstat.ForeColor = RGB(180, 0, 0)
  lbltimecon.Caption = "Time : "
  lblremip.Visible = False
  cmdconn.Caption = "Listen"
  senddataCMD.Enabled = False
End If
If Winsock.State = sckConnected Then
  lblstat = "Connected ..."
  lblstat.ForeColor = RGB(0, 180, 0)
  lblremip.Caption = "Remote IP: " & Winsock.RemoteHostIP
  lbltimecon.Caption = FormatDateTime(Now)
  lblremip.Visible = True
  senddataCMD.Enabled = True
End If
End Sub

'===========================================================================
'======================ALLOWS CLIENT TO CONNECT=============================
Private Sub winsock_ConnectionRequest(ByVal requestID As Long)

If Winsock.State <> sckClosed Then Winsock.CloseSck 'closes connection if one is already open
Winsock.Accept requestID 'allows new connection

Me.Show 'shows server frm when WINSOCK connects to the client
Unload ConnectFRM 'unloads connect frm when WINSOCK connects to the client
Winsock.SendData "C" & nickSERVER.Caption  'send data to server telling to load serverfrm
ServerFRM.Caption = "BoyDread BackDoor     [Welcome, " & nickSERVER.Caption & "!]" 'renames the form approiately
MainTXT.SelText = "~~~~~~~~~~~~~~~~~ Connected to Client ~~~~~~~~~~~~~~~~~" & vbCrLf 'displays that connection worked"

End Sub
'===========================================================================
'=======================RETREIVES DATA FROM CLIENT==========================
Private Sub winsock_DataArrival(ByVal bytesTotal As Long)
Dim strData, strData2 As String 'where the data sent by the client will be stored
Call Winsock.GetData(strData, vbString) 'gets the data sent by the client

strData2 = Left(strData, 1) 'saves the first variable's value to strData2
strData = Mid(strData, 2) 'saves the text the server sent to strData

If strData2 = "T" Then MainTXT.SelText = nickCLIENT.Caption & ":     " & strData & vbCrLf  'adds the data to the txtbox
If strData2 = "N" Then nickCLIENT.Caption = strData 'loads the client's username from data sent

End Sub
'===========================================================================
'========================SENDS DATA TYPED TO CLIENT=========================
Private Sub senddataCMD_Click()
MainTXT.SelText = nickSERVER.Caption & ":     " & dataTXT.Text & vbCrLf 'puts what u typed in ur maintxt
Winsock.SendData "T" & dataTXT.Text 'sends the data to the client
dataTXT.Text = "" 'clears the txtbox u typed in

End Sub
'===========================================================================
