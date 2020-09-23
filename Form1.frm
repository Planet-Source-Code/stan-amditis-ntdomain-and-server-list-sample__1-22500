VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "SQL Servers Only"
      Height          =   315
      Index           =   1
      Left            =   4890
      TabIndex        =   4
      Top             =   270
      Width           =   1635
   End
   Begin VB.OptionButton Option1 
      Caption         =   "All Servers"
      Height          =   315
      Index           =   0
      Left            =   3600
      TabIndex        =   3
      Top             =   270
      Value           =   -1  'True
      Width           =   1065
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   135
      TabIndex        =   1
      Top             =   600
      Width           =   3030
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   3585
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.Label Label2 
      Caption         =   "Domains"
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   5
      Top             =   315
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scanning.   Please wait..."
      Height          =   2790
      Left            =   3585
      TabIndex        =   2
      Top             =   600
      Width           =   3030
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim binit As Boolean
   ' General definitions
   Const ERROR_SUCCESS = 0
   Const ERROR_MORE_DATA = 234
   Const SV_TYPE_SERVER = &H2   'Server type mask, all types of servers
   Const SV_TYPE_SQLSERVER As Long = &H4
   Const SIZE_SI_101 = 24

   Private Type SERVER_INFO_101
      dwPlatformId As Long
      lpszServerName As Long
      dwVersionMajor As Long
      dwVersionMinor As Long
      dwType As Long
      lpszComment As Long
   End Type

   Private Declare Function NetServerEnum Lib "Netapi32.dll" ( _
      ByVal ServerName As String, _
      ByVal Level As Long, _
      Buffer As Long, _
      ByVal PrefMaxLen As Long, _
      EntriesRead As Long, _
      TotalEntries As Long, _
      ByVal servertype As Long, _
      ByVal Domain As String, _
      ResumeHandle As Long) As Long

   Private Declare Function NetAPIBufferFree Lib "Netapi32.dll" Alias _
      "NetApiBufferFree" (BufPtr As Any) As Long
   Private Declare Sub RtlMoveMemory Lib "KERNEL32" ( _
      hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
   Private Declare Function lstrcpyW Lib "KERNEL32" ( _
      ByVal lpszDest As String, ByVal lpszSrc As Long) As Long

   Private Function PointerToString(lpszString As Long) As String
      Dim lpszStr1 As String, lpszStr2 As String, nRes As Long
      lpszStr1 = String(1000, "*")
      nRes = lstrcpyW(lpszStr1, lpszString)
      lpszStr2 = (StrConv(lpszStr1, vbFromUnicode))
      PointerToString = Left(lpszStr2, InStr(lpszStr2, Chr$(0)) - 1)
   End Function

Private Sub RefreshServerList()
    Dim pszTemp As String, pszServer As String, pszDomain As String
    Dim nLevel As Long, I As Long, BufPtr As Long, TempBufPtr As Long
    Dim nPrefMaxLen As Long, nEntriesRead As Long, nTotalEntries As Long
    Dim nServerType As Long, nResumeHandle As Long, nRes As Long
    Dim ServerInfo As SERVER_INFO_101
    Dim nItem As Integer
    Dim nCount As Integer
    
    If List2.SelCount = 0 Then Exit Sub
    
    pszServer = vbNullString
    
    nCount = List2.ListCount
    nItem = 0
    Do
        If nItem >= nCount Then Exit Do
        
        If List2.Selected(nItem) Then
            Label1.Caption = "Scanning " & List2.List(nItem) & "." & vbNewLine & vbNewLine & "Please wait..."
            Me.Refresh
            pszTemp = List2.List(nItem)
            pszDomain = StrConv(pszTemp, vbUnicode)
            
            nLevel = 101
            BufPtr = 0
            nPrefMaxLen = &HFFFFFFFF
            nEntriesRead = 0
            nTotalEntries = 0
            nServerType = IIf(Option1(0).Value, SV_TYPE_SERVER, SV_TYPE_SQLSERVER)
            nResumeHandle = 0
            Do
                nRes = NetServerEnum(pszServer, _
                            nLevel, _
                            BufPtr, _
                            nPrefMaxLen, _
                            nEntriesRead, _
                            nTotalEntries, _
                            nServerType, _
                            pszDomain, _
                            nResumeHandle)
                If (nRes = ERROR_SUCCESS) Or (nRes = ERROR_MORE_DATA) Then
                    TempBufPtr = BufPtr
                    For I = 1 To nEntriesRead
                        RtlMoveMemory ServerInfo, TempBufPtr, SIZE_SI_101
                        List1.AddItem PointerToString(ServerInfo.lpszServerName)
                        TempBufPtr = TempBufPtr + SIZE_SI_101
                    Next I
                Else
                    MsgBox "NetServerEnum failed (" & nRes & ") for Server " & List2.List(nItem)
                End If
                NetAPIBufferFree (BufPtr)
            Loop While nEntriesRead < nTotalEntries
        
        End If
        nItem = nItem + 1
    Loop

End Sub

 
Private Sub Form_Activate()
    Dim vvar As Variant
    Dim DomainList As Variant
    
    If binit = True Then
        Screen.MousePointer = vbHourglass
        
        Label1.Caption = "Scanning Network for Domains." & vbNewLine & vbNewLine & "Please wait..."
        Me.Refresh
        DomainList = EnumDomains()
        
        For Each vvar In DomainList
          List2.AddItem vvar
        Next
        
        Label1.Caption = "Please select a domain to scan."
        Screen.MousePointer = vbNormal
        Me.Refresh
        
        binit = False
    End If
      
End Sub

Private Sub Form_Load()
    binit = True
    Me.Show
      
End Sub

Private Sub List2_Click()
    List1.Clear
    If binit Then Exit Sub
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    List1.Visible = False
    Me.Refresh
    DoEvents
    
    RefreshServerList
    
    Me.Enabled = True
    Screen.MousePointer = vbNormal
    List1.Visible = (List1.ListCount <> 0)
    Label1.Caption = IIf(List1.ListCount <> 0, "", "No servers found.")
    Me.Refresh
    DoEvents
End Sub

Private Sub Option1_Click(Index As Integer)
    If binit Then Exit Sub
    List2_Click
End Sub
