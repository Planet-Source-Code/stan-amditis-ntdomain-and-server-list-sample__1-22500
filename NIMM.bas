Attribute VB_Name = "NIMM"
Option Explicit
' WARNING:
'     If you're not an advanced VB Programmer,
' do not modify anything.  If you are, then
' below you will find utter crazyness! :-)
'
'   - Patrick Herbst


' ---------------------------------------------
' API calls
' ---------------------------------------------
Private Declare Function NetAPIBufferFree Lib "Netapi32.dll" Alias "NetApiBufferFree" (ByVal Ptr As Long) As Long
Private Declare Function NetAPIBufferAllocate Lib "Netapi32.dll" Alias "NetApiBufferAllocate" (ByVal ByteCount As Long, Ptr As Long) As Long
Private Declare Function NetGetDCName Lib "Netapi32.dll" (ServerName As Byte, DomainName As Byte, DCNPtr As Long) As Long
Private Declare Function NetGroupAddUser Lib "Netapi32.dll" (ServerName As Byte, GroupName As Byte, UserName As Byte) As Long
Private Declare Function NetGroupDelUser Lib "Netapi32.dll" (ServerName As Byte, GroupName As Byte, UserName As Byte) As Long
Private Declare Function NetGroupEnum Lib "Netapi32.dll" (ServerName As Byte, ByVal Level As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHandle As Long) As Long
Private Declare Function NetGroupEnumUsers Lib "Netapi32.dll" Alias "NetGroupGetUsers" (ServerName As Byte, GroupName As Byte, ByVal Level As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHandle As Long) As Long
Private Declare Function NetLocalGroupAddMembers Lib "Netapi32.dll" (ByVal psServer As Long, ByVal psLocalGroupName As Long, ByVal Level As Long, pPtrBuffer As Long, ByVal membercount As Long) As Long
Private Declare Function NetLocalGroupDelMembers Lib "Netapi32.dll" (ByVal psServer As Long, ByVal psLocalGroup As Long, ByVal lLevel As Long, uMember As LOCALGROUP_MEMBERS_INFO_0, ByVal lMemberCount As Long) As Long
Private Declare Function NetLocalGroupEnum Lib "Netapi32.dll" (ServerName As Byte, ByVal Level As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHandle As Long) As Long
Private Declare Function NetLocalGroupEnumUsers Lib "Netapi32.dll" Alias "NetLocalGroupGetMembers" (ServerName As Byte, GroupName As Byte, ByVal Level As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHandle As Long) As Long
Private Declare Function NetServerSetInfo Lib "Netapi32.dll" (sServerName As Byte, ByVal lLevel As Long, vBuffer As Long, ParmError As Long) As Long
Private Declare Function NetUserAdd Lib "Netapi32.dll" (ServerName As Byte, ByVal Level As Long, Buffer As USER_INFO_3, parm_err As Long) As Long
Private Declare Function NetUserChangePassword Lib "Netapi32.dll" (ByVal DomainName As String, ByVal UserName As String, ByVal OldPassword As String, ByVal NewPassword As String) As Long
Private Declare Function NetUserDel Lib "Netapi32.dll" (ServerName As Byte, UserName As Byte) As Long
Private Declare Function NetUserEnum Lib "Netapi32.dll" (ServerName As Byte, ByVal Level As Long, ByVal Filter As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHwnd As Long) As Long
Private Declare Function NetUserGetGroups Lib "Netapi32.dll" (ServerName As Byte, UserName As Byte, ByVal Level As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long) As Long
Private Declare Function NetUserGetInfo Lib "Netapi32.dll" (ServerName As Byte, UserName As Byte, ByVal Level As Long, Buffer As Long) As Long
Private Declare Function NetUserGetLocalGroups Lib "Netapi32.dll" (lpServer As Any, UserName As Byte, ByVal Level As Long, ByVal Flags As Long, lpBuffer As Long, ByVal MaxLen As Long, lpEntriesRead As Long, lpTotalEntries As Long) As Long
Private Declare Function NetUserLogon Lib "Advapi32" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As Any, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Private Declare Function NetUserSetInfo Lib "Netapi32.dll" (ByVal ServerName As String, ByVal UserName As String, ByVal Level As Long, UserInfo As Any, ParmError As Long) As Long

Private Declare Sub CopyMem Lib "kernel32.dll" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function PtrToInt Lib "kernel32.dll" Alias "lstrcpynW" (RetVal As Any, ByVal Ptr As Long, ByVal nCharCount As Long) As Long
Private Declare Function PtrToStr Lib "kernel32.dll" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long
Private Declare Function lstrcpyW Lib "kernel32.dll" (bRet As Byte, ByVal lPtr As Long) As Long
Private Declare Function StrCopyA Lib "kernel32.dll" Alias "lstrcpyA" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare Function StrLenA Lib "kernel32.dll" Alias "lstrlenA" (ByVal Ptr As Long) As Long
Private Declare Function StrLenW Lib "kernel32.dll" Alias "lstrlenW" (ByVal Ptr As Long) As Long
Private Declare Function StrToPtr Lib "kernel32.dll" Alias "lstrcpyW" (ByVal Ptr As Long, Source As Byte) As Long

Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal p_lngEnumHwnd As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal pEnumHwnd As Long, lpcCount As Long, lpBuffer As NETRESOURCE, lpBufferSize As Long) As Long
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lppEnumHwnd As Long) As Long


' ---------------------------------------------
' Possible errors with API call
' ---------------------------------------------

Private Const ERROR_SUCCESS As Long = 0&
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const NERR_BASE As Long = 2100
Private Const NERR_GroupExists As Long = NERR_BASE + 123
Private Const NERR_NotPrimary As Long = NERR_BASE + 126
Private Const NERR_UserExists As Long = NERR_BASE + 124
Private Const NERR_PasswordTooShort As Long = NERR_BASE + 145
Private Const NERR_InvalidComputer As Long = NERR_BASE + 251
Private Const NERR_Success As Long = 0&

' ---------------------------------------------
' General constants used
' ---------------------------------------------

Private Const USER_PRIV_MASK = &H3
Private Const USER_PRIV_GUEST = &H0
Private Const USER_PRIV_USER = &H1
Private Const USER_PRIV_ADMIN = &H2
Private Const constUserInfoLevel3 As Long = 3
Private Const TIMEQ_FOREVER As Long = -1&
Private Const MAX_PATH As Long = 260&
Private Const DOMAIN_GROUP_RID_USERS As Long = &H201&
Private Const USER_MAXSTORAGE_UNLIMITED As Long = -1&
Private Const LocalGroupMembersInfo3 As Long = 3&
Private Const MAX_RESOURCES As Long = 256
Private Const MAX_COMPUTERNAME As Long = 15
Private Const MAX_USERNAME As Long = 256
Private Const NOT_A_CONTAINER As Long = -1
Private Const RESOURCE_GLOBALNET As Long = &H2&
Private Const RESOURCETYPE_ANY As Long = &H0&
Private Const RESOURCEUSAGE_ALL As Long = &H0&
Private Const NO_ERROR As Long = 0&
Private Const RESOURCE_ENUM_ALL As Long = &HFFFF

' ---------------------------------------------
' Constants used by LogonUser
' ---------------------------------------------

Private Const LOGON32_PROVIDER_DEFAULT As Long = 0&
Private Const LOGON32_PROVIDER_WINNT35 As Long = 1&
Private Const LOGON32_LOGON_INTERACTIVE As Long = 2&
Private Const LOGON32_LOGON_NETWORK As Long = 3&
Private Const LOGON32_LOGON_BATCH As Long = 4&
Private Const LOGON32_LOGON_SERVICE As Long = 5&

' ---------------------------------------------
' Used by usri3_flags element of data structure
' ---------------------------------------------
Private Const UF_SCRIPT = &H1
Private Const UF_ACCOUNTDISABLE = &H2
Private Const UF_HOMEDIR_REQUIRED = &H8
Private Const UF_LOCKOUT = &H10
Private Const UF_PASSWD_NOTREQD = &H20
Private Const UF_PASSWD_CANT_CHANGE = &H40
Private Const UF_NORMAL_ACCOUNT = &H200
Private Const UF_DONT_EXPIRE_PASSWD As Long = &H10000
Private Const UF_SERVER_TRUST_ACCOUNT As Long = &H2000&
Private Const UF_TEMP_DUPLICATE_ACCOUNT As Long = &H100&
Private Const UF_INTERDOMAIN_TRUST_ACCOUNT As Long = &H800&
Private Const UF_WORKSTATION_TRUST_ACCOUNT As Long = &H1000&
Private Const STILL_ACTIVE As Long = &H103&
Private Const PROCESS_QUERY_INFORMATION As Long = &H400&
Private Const FILTER_NORMAL_ACCOUNT = &H2

Private Type MungeLong
  X As Long
  Dummy As Integer
End Type

Private Type MungeInt
  XLo As Integer
  XHi As Integer
  Dummy As Integer
End Type

Private Type TUser0                    ' Level 0
  ptrName As Long
End Type

Private Type TUser1                    ' Level 1
  ptrName As Long
  ptrPassword As Long
  dwPasswordAge As Long
  dwPriv As Long
  ptrHomeDir As Long
  ptrComment As Long
  dwFlags As Long
  ptrScriptPath As Long
End Type

' ---------------------------------------------
' The USER_INFO_3 data structure
' ---------------------------------------------

Private Type USER_INFO_3
  usri3_name As Long
  usri3_password As Long
  usri3_password_age As Long
  usri3_priv As Long
  usri3_home_dir As Long
  usri3_comment As Long
  usri3_flags As Long
  usri3_script_path As Long
  usri3_auth_flags As Long
  usri3_full_name As Long
  usri3_usr_comment As Long
  usri3_parms As Long
  usri3_workstations As Long
  usri3_last_logon As Long
  usri3_last_logoff As Long
  usri3_acct_expires As Long
  usri3_max_storage As Long
  usri3_units_per_week As Long
  usri3_logon_hours As Long
  usri3_bad_pw_count As Long
  usri3_num_logons As Long
  usri3_logon_server As Long
  usri3_country_code As Long
  usri3_code_page As Long
  usri3_user_id As Long
  usri3_primary_group_id As Long
  usri3_profile As Long
  usri3_home_dir_drive As Long
  usri3_password_expired As Long
End Type

Private Type USERINFO_2_API
  usri2_name As Long
  usri2_password As Long
  usri2_password_age As Long
  usri2_priv As Long
  usri2_home_dir As Long
  usri2_comment As Long
  usri2_flags As Long
  usri2_script_path As Long
  usri2_auth_flags As Long
  usri2_full_name As Long
  usri2_usr_comment As Long
  usri2_parms As Long
  usri2_workstations As Long
  usri2_last_logon As Long
  usri2_last_logoff As Long
  usri2_acct_expires As Long
  usri2_max_storage As Long
  usri2_units_per_week As Long
  usri2_logon_hours As Long
  usri2_bad_pw_count As Long
  usri2_num_logons As Long
  usri2_logon_server As Long
  usri2_country_code As Long
  usri2_code_page As Long
End Type

Private Type USER_INFO_10_API
  Name As Long
  Comment As Long
  UsrComment As Long
  FullName As Long
End Type

Private Type USER_INFO_1003
  usri1003_password As Long
End Type

Public Type USER_INFO
  Name As String
  FullName As String
  Comment As String
  UserComment As String
End Type

Private Type LOCALGROUP_MEMBERS_INFO_0
  pSID As Long
End Type

Private Type LOCALGROUP_MEMBERS_INFO_3
  DomainAndName As Long
End Type

' Type used by NetServerSetInfo

Private Type SERVER_INFO_1005
  sv1005_comment As Long
End Type

Private Type NETRESOURCE
  dwScope As Long
  dwType As Long
  dwDisplayType As Long
  dwUsage As Long
  pLocalName As Long
  pRemoteName As Long
  pComment As Long
  pProvider As Long
End Type

' *******************************************************
' Add a user either to NT -- you *MUST* have admin or
' account operator priviledges to successfully run
' this function
' Use on NT Only
' *******************************************************

Public Function AddNewUser(ByVal Server As String, ByVal UserName As String, ByVal Password As String, Optional ByVal FullName As String = vbNullString, Optional ByVal UserComment As String = vbNullString) As Boolean
  Dim p_strErr As String
  Dim p_lngRtn As Long
  Dim p_lngPtrUserName As Long
  Dim p_lngPtrPassword As Long
  Dim p_lngPtrUserFullName As Long
  Dim p_lngPtrUserComment As Long
  Dim p_lngParameterErr As Long
  Dim p_lngFlags As Long
  Dim p_abytServerName() As Byte
  Dim p_abytUserName() As Byte
  Dim p_abytPassword() As Byte
  Dim p_abytUserFullName() As Byte
  Dim p_abytUserComment() As Byte
  Dim p_typUserInfo3 As USER_INFO_3

  If Server <> "" And Left(Server, 2) <> "\\" Then Server = "\\" & Server

  If FullName = vbNullString Then
    FullName = UserName
  End If

  ' ------------------------------------------
  ' Create byte arrays to avoid Unicode hassles
  ' ------------------------------------------

  p_abytServerName = Server & vbNullChar
  p_abytUserName = UserName & vbNullChar
  p_abytUserFullName = FullName & vbNullChar
  p_abytPassword = Password & vbNullChar
  p_abytUserComment = UserComment & vbNullChar

  ' ------------------------------------------
  ' Allocate buffer space
  ' ------------------------------------------

  p_lngRtn = NetAPIBufferAllocate(UBound(p_abytUserName), p_lngPtrUserName)

  p_lngRtn = NetAPIBufferAllocate(UBound(p_abytUserFullName), p_lngPtrUserFullName)

  p_lngRtn = NetAPIBufferAllocate(UBound(p_abytPassword), p_lngPtrPassword)

  p_lngRtn = NetAPIBufferAllocate(UBound(p_abytUserComment), p_lngPtrUserComment)

  ' ------------------------------------------
  ' Get pointers to the byte arrays
  ' ------------------------------------------

  p_lngPtrUserName = VarPtr(p_abytUserName(0))
  p_lngPtrUserFullName = VarPtr(p_abytUserFullName(0))
  p_lngPtrPassword = VarPtr(p_abytPassword(0))
  p_lngPtrUserComment = VarPtr(p_abytUserComment(0))

  ' ------------------------------------------
  ' Fill the VB structure
  ' ------------------------------------------

  p_lngFlags = UF_NORMAL_ACCOUNT Or _
  UF_SCRIPT Or _
  UF_DONT_EXPIRE_PASSWD

  With p_typUserInfo3
    .usri3_acct_expires = TIMEQ_FOREVER ' Never expires
    .usri3_comment = p_lngPtrUserComment ' Comment
    .usri3_flags = p_lngFlags ' There are a number of variations
    .usri3_full_name = p_lngPtrUserFullName ' User's full name
    .usri3_max_storage = USER_MAXSTORAGE_UNLIMITED ' Can use any amount
    'of disk space
    .usri3_name = p_lngPtrUserName ' Name of user account
    .usri3_password = p_lngPtrPassword ' Password for user account
    .usri3_primary_group_id = DOMAIN_GROUP_RID_USERS ' You MUST use this
    'constant for NetUserAdd
    .usri3_script_path = 0& ' Path of user's logon script
    .usri3_auth_flags = 0& ' Ignored by NetUserAdd
    .usri3_bad_pw_count = 0& ' Ignored by NetUserAdd
    .usri3_code_page = 0& ' Code page for user's language
    .usri3_country_code = 0& ' Country code for user's language
    .usri3_home_dir = 0& ' Can specify path of home directory of this
    'user
    .usri3_home_dir_drive = 0& ' Drive letter assign to user's
    'profile
    .usri3_last_logoff = 0& ' Not needed when adding a user
    .usri3_last_logon = 0& ' Ignored by NetUserAdd
    .usri3_logon_hours = 0& ' Null means no restrictions
    .usri3_logon_server = 0& ' Null means logon to domain server
    .usri3_num_logons = 0& ' Ignored by NetUserAdd
    .usri3_parms = 0& ' Used by specific applications
    .usri3_password_age = 0& ' Ignored by NetUserAdd
    .usri3_password_expired = 0& ' None-zero means user must change
    'password at next logon
    .usri3_priv = 0& ' Ignored by NetUserAdd
    .usri3_profile = 0& ' Path to a user's profile
    .usri3_units_per_week = 0& ' Ignored by NetUserAdd
    .usri3_user_id = 0& ' Ignored by NetUserAdd
    .usri3_usr_comment = 0& ' User comment
    .usri3_workstations = 0& ' Workstations a user can log onto (null
    '= all stations)
  End With

  ' ------------------------------------------
  ' Attempt to add the user
  ' ------------------------------------------

  p_lngRtn = NetUserAdd(p_abytServerName(0), _
  constUserInfoLevel3, p_typUserInfo3, p_lngParameterErr)

  ' ------------------------------------------
  ' Check for error
  ' ------------------------------------------

  If p_lngRtn <> 0 Then
    AddNewUser = False
    Select Case p_lngRtn
      Case ERROR_ACCESS_DENIED
        p_strErr = "User doesn't have sufficient access rights."
      Case NERR_GroupExists
        p_strErr = "The group already exists."
      Case NERR_NotPrimary
        p_strErr = "Can only do this operation on the PDC of the domain."
      Case NERR_UserExists
        p_strErr = "The user account already exists."
      Case NERR_PasswordTooShort
        p_strErr = "The password is shorter than required."
      Case NERR_InvalidComputer
        p_strErr = "The computer name is invalid."
      Case Else
        p_strErr = "Unknown error #" & CStr(p_lngRtn)
    End Select
    On Error GoTo 0
    Err.Raise p_lngRtn, p_strErr & vbCrLf & "Error in parameter " & p_lngParameterErr & " when attempting to add the user, " & UserName
  Else
    AddNewUser = True
  End If

  ' ------------------------------------------
  ' Be a good programmer and free the memory
  ' you've allocated
  ' ------------------------------------------

  p_lngRtn = NetAPIBufferFree(p_lngPtrUserName)
  p_lngRtn = NetAPIBufferFree(p_lngPtrPassword)
  p_lngRtn = NetAPIBufferFree(p_lngPtrUserFullName)
  p_lngRtn = NetAPIBufferFree(p_lngPtrUserComment)

End Function

Public Function DelUser(ByVal Server As String, ByVal UserName As String) As Boolean
  Dim UNArray() As Byte
  Dim SNArray() As Byte
  
  If Server <> "" And Left(Server, 2) <> "\\" Then Server = "\\" & Server
  
  UNArray = UserName & vbNullChar
  SNArray = Server & vbNullChar
  DelUser = IIf(NetUserDel(SNArray(0), UNArray(0)), False, True) ' If result is 0, then success
End Function
   
Public Function DelUserFromGlobalGroup(ByVal Server As String, ByVal UserName As String, ByVal GlobalGroup As String) As Boolean
  '
  ' This only deletes users from global groups - not local groups
  '
  Dim SNArray() As Byte, GNArray() As Byte, UNArray() As Byte, Result As Long
  
  If Server <> "" And Left(Server, 2) <> "\\" Then Server = "\\" & Server
  
  SNArray = Server & vbNullChar
  GNArray = GlobalGroup & vbNullChar
  UNArray = UserName & vbNullChar
  Result = NetGroupDelUser(SNArray(0), GNArray(0), UNArray(0))
  If Result = 2220 Then
    Err.Raise 2220, "DelUserFromGlobalGroup", "There is no **GLOBAL** group '" & GlobalGroup & "'"
    DelUserFromGlobalGroup = False
    Exit Function
  End If
  DelUserFromGlobalGroup = IIf(Result, False, True) ' If 0 it's success
End Function

Public Function DelUserFromLocalGroup(ByVal Server As String, ByVal UserName As String, ByVal LocalGroup As String) As Boolean

  Dim p_lngPtrGroupName As Long
  Dim p_lngPtrUserName As Long
  Dim p_lngPtrServerName As Long
  Dim p_lngMemberCount As Long
  Dim p_lngRtn As Long
  Dim p_usersid As LOCALGROUP_MEMBERS_INFO_0

  If Server <> "" And Left(Server, 2) <> "\\" Then Server = "\\" & Server

  ' Convert the server name to a pointer
  If Len(Trim$(Server)) = 0 Then
    p_lngPtrServerName = 0&
  Else
    p_lngPtrServerName = StrPtr(Server)
  End If

  ' Convert the group name to a pointer
  p_lngPtrGroupName = StrPtr(LocalGroup)

  ' Convert the user name to a pointer
  p_lngPtrUserName = StrPtr(UserName)
  p_usersid.pSID = p_lngPtrUserName

  ' Add the user
  p_lngMemberCount = 1

  p_lngRtn = NetLocalGroupDelMembers(p_lngPtrServerName, p_lngPtrGroupName, LocalGroupMembersInfo3, p_usersid, p_lngMemberCount)

  If p_lngRtn = NERR_Success Then
    DelUserFromLocalGroup = True
  Else
    DelUserFromLocalGroup = False
  End If

End Function
 
' Works only on Win NT
Public Function GetPDCName(Optional ByVal Domain As String, Optional ByVal Server As String) As String
  Dim Result As Long, DCName As String, DCNPtr As Long
  Dim DNArray() As Byte, MNArray() As Byte, DCNArray(100) As Byte
  
  If Server <> "" And Left(Server, 2) <> "\\" Then Server = "\\" & Server
  
  MNArray = Server & vbNullChar
  DNArray = Domain & vbNullChar
  Result = NetGetDCName(MNArray(0), DNArray(0), DCNPtr)
  If Result <> 0 Then
    GetPDCName = ""
    Exit Function
  End If
  Result = PtrToStr(DCNArray(0), DCNPtr)
  Result = NetAPIBufferFree(DCNPtr)
  DCName = DCNArray()
  GetPDCName = DCName
End Function


Public Function AddUserToGlobalGroup(ByVal Server As String, ByVal UserName As String, ByVal GlobalGroup As String) As Boolean
  Dim SNArray() As Byte, GNArray() As Byte, UNArray() As Byte, Result As Long
  
  If Server <> "" And Left(Server, 2) <> "\\" Then Server = "\\" & Server
  
  SNArray = Server & vbNullChar
  GNArray = GlobalGroup & vbNullChar
  UNArray = UserName & vbNullChar
  Result = NetGroupAddUser(SNArray(0), GNArray(0), UNArray(0))
  If Result = 2220 Then
    Err.Raise Result, "AddUserToGlobalGroup", "There is no **GLOBAL** group '" & GlobalGroup & "'"
    AddUserToGlobalGroup = False
    Exit Function
  End If
  AddUserToGlobalGroup = IIf(Result, False, True) ' 0 (aka false) actually means success
End Function

' Use on NT Only
Public Function AddUserToLocalGroup(ByVal Server As String, ByVal UserName As String, ByVal LocalGroup As String) As Boolean
  Dim p_lngPtrGroupName As Long
  Dim p_lngPtrUserName As Long
  Dim p_lngPtrServerName As Long
  Dim p_lngMemberCount As Long
  Dim p_lngRtn As Long

  If Server <> "" And Left(Server, 2) <> "\\" Then Server = "\\" & Server
  
  ' Convert the server name to a pointer
  If Len(Trim$(Server)) = 0 Then
    p_lngPtrServerName = 0&
  Else
    p_lngPtrServerName = StrPtr(Server)
  End If

  ' Convert the group name to a pointer
  p_lngPtrGroupName = StrPtr(LocalGroup)

  ' Convert the user name to a pointer
  p_lngPtrUserName = StrPtr(UserName)

  ' Add the user
  p_lngMemberCount = 1

  p_lngRtn = NetLocalGroupAddMembers(p_lngPtrServerName, p_lngPtrGroupName, LocalGroupMembersInfo3, p_lngPtrUserName, p_lngMemberCount)

  If p_lngRtn = NERR_Success Then
    AddUserToLocalGroup = True
  Else
    AddUserToLocalGroup = False
  End If

End Function
 
' Works on Win 95 & NT
Public Function SetServerComment(ByVal Comment As String, Optional ByVal Server As String = "") As Boolean
  Dim p_bytServerName() As Byte
  Dim p_lngRtn As Long
  Dim p_lngSrvInfoRtn As Long
  Dim p_lngServEnumLevel As Long
  Dim p_lngParmError As Long
  Dim p_lngStrPtr As Long
  
  If Server <> "" And Left(Server, 2) <> "\\" Then Server = "\\" & Server
  
  ' Initialize the variables
  If Trim$(Server) = vbNullString Then
    p_bytServerName = vbNullChar
  Else
    p_bytServerName = Trim$(Server) & vbNullChar
  End If

  p_lngServEnumLevel = 1005
  p_lngStrPtr = StrPtr(Comment)
  p_lngRtn = NetServerSetInfo(p_bytServerName(0), p_lngServEnumLevel, p_lngStrPtr, p_lngParmError)

  If p_lngRtn = 0 Then
    SetServerComment = True
  Else
    SetServerComment = False
    Err.Raise Err.LastDllError, "SetServerComment"
  End If

End Function

' Works on Win 95 & NT
Public Function Login(ByVal UserName As String, ByVal Password As String) As Boolean

  On Error Resume Next ' Don't accept errors here

  Dim p_lngToken As Long
  Dim p_lngRtn As Long

  p_lngRtn = NetUserLogon(UserName, 0&, Password, LOGON32_LOGON_NETWORK, LOGON32_PROVIDER_DEFAULT, p_lngToken)

  If p_lngRtn = 0 Then
    Login = False
  Else
    Login = True
  End If

  On Error GoTo 0

End Function

' Works on Win 95 & NT
Public Function EnumDomains() As Variant
  Dim p_avntDomains As Variant
  Dim p_lngNumItems As Long
  Dim p_lngRtn As Long
  Dim p_lngEnumHwnd As Long
  Dim p_lngCount As Long
  Dim p_lngLoop As Long
  Dim p_lngBufSize As Long
  Dim p_astrDomainNames() As String
  Dim p_atypNetAPI(0 To MAX_RESOURCES) As NETRESOURCE
  
  ' First time through, find the root level
  p_lngEnumHwnd = 0&
  p_lngRtn = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, RESOURCEUSAGE_ALL, 0&, p_lngEnumHwnd)
 
  If p_lngRtn = NO_ERROR Then
    
    p_lngCount = RESOURCE_ENUM_ALL
    p_lngBufSize = UBound(p_atypNetAPI) * Len(p_atypNetAPI(0))
    p_lngRtn = WNetEnumResource(p_lngEnumHwnd, p_lngCount, p_atypNetAPI(0), p_lngBufSize)
 
  End If
  If p_lngEnumHwnd <> 0 Then
    Call WNetCloseEnum(p_lngEnumHwnd)
  End If

  ' Now we are going for the second level, which should contain the domain names
  p_lngRtn = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, RESOURCEUSAGE_ALL, p_atypNetAPI(0), p_lngEnumHwnd)
  
  If p_lngRtn = NO_ERROR Then
    p_lngCount = RESOURCE_ENUM_ALL
    p_lngBufSize = UBound(p_atypNetAPI) * Len(p_atypNetAPI(0))
    p_lngRtn = WNetEnumResource(p_lngEnumHwnd, p_lngCount, p_atypNetAPI(0), p_lngBufSize)
    If p_lngCount > 0 Then
      ReDim p_astrDomainNames(1 To p_lngCount) As String
      For p_lngLoop = 0 To p_lngCount - 1
        p_astrDomainNames(p_lngLoop + 1) = PointerToAsciiStr(p_atypNetAPI(p_lngLoop).pRemoteName)
      Next p_lngLoop
    End If
  End If
 
  If p_lngEnumHwnd <> 0 Then
    Call WNetCloseEnum(p_lngEnumHwnd)
  End If
  
  EnumDomains = p_astrDomainNames
End Function

Public Function EnumGlobalGroups(ByVal Server As String, Optional ByVal UserName As String) As Variant
  ' Enumerates global groups only - not local groups
  ' Returns an array of global groups
  ' If a username is specified, it only returns
  ' groups that that user is a member of
  Dim Result As Long
  Dim BufPtr As Long
  Dim EntriesRead As Long
  Dim TotalEntries As Long
  Dim ResumeHandle As Long
  Dim BufLen As Long
  Dim SNArray() As Byte
  Dim GNArray(99) As Byte
  Dim UNArray() As Byte
  Dim GName As String
  Dim I As Integer
  Dim UNPtr As Long
  Dim TempPtr As MungeLong
  Dim TempStr As MungeInt
  Dim Groups() As String
  Dim Pass As Long

  If Server <> "" And Left(Server, 2) <> "\\" Then Server = "\\" & Server

  SNArray = Server & vbNullChar      ' Move to byte array
  UNArray = UserName & vbNullChar    ' Move to Byte array
  BufLen = 255                       ' Buffer size
  ResumeHandle = 0                   ' Start with the first entry

  Pass = 0
  Do
    If UserName = "" Then
      Result = NetGroupEnum(SNArray(0), 0, BufPtr, BufLen, EntriesRead, TotalEntries, ResumeHandle)
    Else
      Result = NetUserGetGroups(SNArray(0), UNArray(0), 0, BufPtr, BufLen, EntriesRead, TotalEntries)
    End If
    EnumGlobalGroups = Result
    If Result <> 0 And Result <> 234 Then    ' 234 means multiple reads required
      Err.Raise Result, "EnumGlobalGroups", "Error " & Result & " enumerating global group " & EntriesRead & " of " & TotalEntries
      Exit Function
    End If
    For I = 1 To EntriesRead
      ' Get pointer to string from beginning of buffer
      ' Copy 4 byte block of memory in 2 steps
      PtrToInt TempStr.XLo, BufPtr + (I - 1) * 4, 2
      PtrToInt TempStr.XHi, BufPtr + (I - 1) * 4 + 2, 2
      LSet TempPtr = TempStr ' munge 2 Integers to a Long
      ' Copy string to array and convert to a string
      Result = PtrToStr(GNArray(0), TempPtr.X)
      GName = Left(GNArray, StrLenW(TempPtr.X))
      ReDim Preserve Groups(0 To Pass) As String
      Groups(Pass) = GName
      Pass = Pass + 1
    Next I
  Loop Until EntriesRead = TotalEntries
  ' The above condition only valid for reading accounts on NT
  ' but not OK for OS/2 or LanMan
  NetAPIBufferFree BufPtr         ' Don't leak memory
  
  EnumGlobalGroups = Groups()
End Function

Public Function EnumLocalGroups(ByVal Server As String, Optional ByVal UserName As String) As Variant
  ' Enumerates local groups only - not global groups
  ' Returns an array of local groups
  ' If a username is specified, it only returns
  ' groups that that user is a member of
  Dim Result As Long
  Dim BufPtr As Long
  Dim EntriesRead As Long
  Dim TotalEntries As Long
  Dim ResumeHandle As Long
  Dim BufLen As Long
  Dim SNArray() As Byte
  Dim GNArray(99) As Byte
  Dim UNArray() As Byte
  Dim GName As String
  Dim I As Integer
  Dim UNPtr As Long
  Dim TempPtr As MungeLong
  Dim TempStr As MungeInt
  Dim Pass As Long
  Dim Groups() As String

  If Server <> "" And Left(Server, 2) <> "\\" Then Server = "\\" & Server

  SNArray = Server & vbNullChar      ' Move to byte array
  UNArray = UserName & vbNullChar    ' Move to Byte array
  BufLen = 255                       ' Buffer size
  ResumeHandle = 0                   ' Start with the first entry
  
  Pass = 0
  Do
    If UserName = "" Then
      Result = NetLocalGroupEnum(SNArray(0), 0, BufPtr, BufLen, EntriesRead, TotalEntries, ResumeHandle)
    Else
      Result = NetUserGetLocalGroups(SNArray(0), UNArray(0), 0, 0, BufPtr, BufLen, EntriesRead, TotalEntries)
    End If
    
    If Result <> 0 And Result <> 234 Then    ' 234 means multiple reads required
      Err.Raise Result, "EnumLocalGroups", "Error enumerating local group " & EntriesRead & " of " & TotalEntries
      Exit Function
    End If

    For I = 1 To EntriesRead
      ' Get pointer to string from beginning of buffer
      ' Copy 4 byte block of memory in 2 steps
      PtrToInt TempStr.XLo, BufPtr + (I - 1) * 4, 2
      PtrToInt TempStr.XHi, BufPtr + (I - 1) * 4 + 2, 2
      LSet TempPtr = TempStr ' munge 2 Integers to a Long
      ' Copy string to array and convert to a string
      Result = PtrToStr(GNArray(0), TempPtr.X)
      GName = Left(GNArray, StrLenW(TempPtr.X))
      ReDim Preserve Groups(0 To Pass) As String
      Groups(Pass) = GName
      Pass = Pass + 1
    Next I
  Loop Until EntriesRead = TotalEntries
  ' The above condition only valid for reading accounts on NT
  ' but not OK for OS/2 or LanMan
  NetAPIBufferFree BufPtr         ' Don't leak memory
  
  EnumLocalGroups = Groups()
End Function


Public Function EnumUsers(ByVal Server As String) As Variant
  Dim Users() As String
  Dim Result As Long
  Dim BufPtr As Long
  Dim EntriesRead As Long
  Dim TotalEntries As Long
  Dim ResumeHandle As Long
  Dim BufLen As Long
  Dim SNArray() As Byte
  Dim GNArray() As Byte
  Dim UNArray(99) As Byte
  Dim UName As String
  Dim I As Integer
  Dim UNPtr As Long
  Dim TempPtr As MungeLong
  Dim TempStr As MungeInt
  Dim Pass As Long

  If Server <> "" And Left(Server, 2) <> "\\" Then Server = "\\" & Server

  SNArray = Server & vbNullChar       ' Move to byte array
  BufLen = 255                       ' Buffer size
  ResumeHandle = 0                   ' Start with the first entry
  
  Pass = 0
  Do
    Result = NetUserEnum(SNArray(0), 0, FILTER_NORMAL_ACCOUNT, BufPtr, BufLen, EntriesRead, TotalEntries, ResumeHandle)
    
    If Result <> 0 And Result <> 234 Then    ' 234 means multiple reads required
      Err.Raise Result, "EnumUsers", "Error enumerating user " & EntriesRead & " of " & TotalEntries
      Exit Function
    End If
    For I = 1 To EntriesRead
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(TempStr.XLo, BufPtr + (I - 1) * 4, 2)
      Result = PtrToInt(TempStr.XHi, BufPtr + (I - 1) * 4 + 2, 2)
      LSet TempPtr = TempStr ' munge 2 integers into a Long
      ' Copy string to array
      Result = PtrToStr(UNArray(0), TempPtr.X)
      UName = Left(UNArray, StrLenW(TempPtr.X))
      ReDim Preserve Users(0 To Pass)
      Users(Pass) = UName
      Pass = Pass + 1
    Next I
  Loop Until EntriesRead = TotalEntries
  Result = NetAPIBufferFree(BufPtr)         ' Don't leak memory
  EnumUsers = Users()
End Function

Public Function EnumUsersInGlobalGroup(ByVal Server As String, ByVal GlobalGroup As String) As Variant
  Dim Users() As String
  Dim Result As Long
  Dim BufPtr As Long
  Dim EntriesRead As Long
  Dim TotalEntries As Long
  Dim ResumeHandle As Long
  Dim BufLen As Long
  Dim SNArray() As Byte
  Dim GNArray() As Byte
  Dim UNArray(99) As Byte
  Dim UName As String
  Dim I As Integer
  Dim UNPtr As Long
  Dim TempPtr As MungeLong
  Dim TempStr As MungeInt
  Dim Pass As Long

  If Server <> "" And Left(Server, 2) <> "\\" Then Server = "\\" & Server

  SNArray = Server & vbNullChar       ' Move to byte array
  GNArray = GlobalGroup & vbNullChar       ' Move to Byte array
  BufLen = 255                       ' Buffer size
  ResumeHandle = 0                   ' Start with the first entry
  
  Pass = 0
  Do
    If GlobalGroup <> "" Then
      Result = NetGroupEnumUsers(SNArray(0), GNArray(0), 0, BufPtr, BufLen, EntriesRead, TotalEntries, ResumeHandle)
    Else
      Err.Raise 0, "EnumUsersInGlobalGroup", "You must specify a global group."
      Exit Function
    End If
    If Result <> 0 And Result <> 234 Then    ' 234 means multiple reads required
      If Result = 2220 Then Err.Raise 2220, "EnumUsersInGlobalGroup", "There is no global group '" & GlobalGroup & "'"
      Err.Raise Result, "EnumUsersInGlobalGroup", "Error enumerating user " & EntriesRead & " of " & TotalEntries
      Exit Function
    End If
    For I = 1 To EntriesRead
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(TempStr.XLo, BufPtr + (I - 1) * 4, 2)
      Result = PtrToInt(TempStr.XHi, BufPtr + (I - 1) * 4 + 2, 2)
      LSet TempPtr = TempStr ' munge 2 integers into a Long
      ' Copy string to array
      Result = PtrToStr(UNArray(0), TempPtr.X)
      UName = Left(UNArray, StrLenW(TempPtr.X))
      
      If Right(UName, 1) <> "$" Then    ' For some reason computer names
        ReDim Preserve Users(0 To Pass) ' are returned too, with a $ on the
        Users(Pass) = UName             ' end.  This is to weed them out.
        Pass = Pass + 1
      End If
    Next I
  Loop Until EntriesRead = TotalEntries
  Result = NetAPIBufferFree(BufPtr)         ' Don't leak memory
  EnumUsersInGlobalGroup = Users()
End Function

Public Function EnumUsersInLocalGroup(ByVal Server As String, ByVal LocalGroup As String) As Variant
  Dim Users() As String
  Dim Result As Long
  Dim BufPtr As Long
  Dim EntriesRead As Long
  Dim TotalEntries As Long
  Dim ResumeHandle As Long
  Dim BufLen As Long
  Dim SNArray() As Byte
  Dim GNArray() As Byte
  Dim UNArray(99) As Byte
  Dim UName As String
  Dim I As Integer
  Dim UNPtr As Long
  Dim TempPtr As MungeLong
  Dim TempStr As MungeInt
  Dim Pass As Long

  If Server <> "" And Left(Server, 2) <> "\\" Then Server = "\\" & Server

  SNArray = Server & vbNullChar       ' Move to byte array
  GNArray = LocalGroup & vbNullChar       ' Move to Byte array
  BufLen = 255                       ' Buffer size
  ResumeHandle = 0                   ' Start with the first entry
  
  Pass = 0
  Do
    If LocalGroup <> "" Then
      Result = NetLocalGroupEnumUsers(SNArray(0), GNArray(0), 0, BufPtr, BufLen, EntriesRead, TotalEntries, ResumeHandle)
    Else
      Err.Raise 0, EnumUsersInLocalGroup, "You must specify a local group"
      Exit Function
    End If
    If Result <> 0 And Result <> 234 Then    ' 234 means multiple reads required
      If Result = 2220 Then Err.Raise 2220, EnumUsersInLocalGroup, "There is no local group '" & LocalGroup & "'"
      Err.Raise Result, EnumUsersInLocalGroup, "Error enumerating user " & EntriesRead & " of " & TotalEntries
      Exit Function
    End If
    For I = 1 To EntriesRead
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(TempStr.XLo, BufPtr + (I - 1) * 4, 2)
      Result = PtrToInt(TempStr.XHi, BufPtr + (I - 1) * 4 + 2, 2)
      LSet TempPtr = TempStr ' munge 2 integers into a Long
      ' Copy string to array
      Result = PtrToStr(UNArray(0), TempPtr.X)
      UName = Left(UNArray, StrLenW(TempPtr.X))
      ReDim Preserve Users(0 To Pass)
      Users(Pass) = UName
      Pass = Pass + 1
    Next I
  Loop Until EntriesRead = TotalEntries
  Result = NetAPIBufferFree(BufPtr)         ' Don't leak memory
  EnumUsersInLocalGroup = Users()
End Function


Private Function PointerToAsciiStr(ByVal xi_lngPtrToString As Long) As String
  On Error Resume Next ' Don't accept an error here
  Dim p_lngLen As Long
  Dim p_strStringValue As String
  Dim p_lngNullPos As Long
  Dim p_lngRtn As Long
  
  p_lngLen = StrLenA(xi_lngPtrToString)
  If xi_lngPtrToString > 0 And p_lngLen > 0 Then
    p_strStringValue = Space$(p_lngLen + 1)
    p_lngRtn = StrCopyA(p_strStringValue, xi_lngPtrToString)
    p_lngNullPos = InStr(p_strStringValue, Chr$(0))
    If p_lngNullPos > 0 Then
      'Lose the null terminator...
      PointerToAsciiStr = Left$(p_strStringValue, p_lngNullPos - 1)
    Else
      PointerToAsciiStr = p_strStringValue 'Just pass the string...
    End If
  Else
    PointerToAsciiStr = ""
  End If
End Function

' Works on Win NT only
Public Function EnumUsersLoggedOn(ByVal Server As String) As Variant
  Dim p_lngRtn As Long
  Dim p_lngPtrBuffer As Long
  Dim p_lngPtrUserInfoBuf As Long
  Dim p_lngEntriesRead As Long
  Dim p_lngTotalEntries As Long
  Dim p_lngResumeHwnd As Long
  Dim p_lngLoop As Long
  Dim p_lngLastLogon As Long
  Dim p_lngLastLogoff As Long
  Dim p_strUserName As String
  Dim p_abytServerName() As Byte
  Dim p_abytUserName() As Byte
  Dim p_atypUserInfo() As USER_INFO_10_API
  Dim p_typUserInfo As USERINFO_2_API
  Dim LoggedOnUsers() As String
  Dim Pass As Long
  
  If Server <> "" And Left(Server, 2) <> "\\" Then Server = "\\" & Server
  
  ' ------------------------------------------
  ' Initialize the variable(s)
  ' ------------------------------------------
 
  p_abytServerName = Server & Chr$(0)

  ' ------------------------------------------
  ' Make appropriate API call and check for error
  ' ------------------------------------------

  p_lngRtn = NetUserEnum(p_abytServerName(0), 10, 0&, p_lngPtrBuffer, &H4000, p_lngEntriesRead, p_lngTotalEntries, p_lngResumeHwnd)

  If p_lngRtn <> 0 Then
    Err.Raise p_lngRtn, "EnumUsersLoggedOn", "Had an error enumerating users."
    Exit Function
  End If

  ' ------------------------------------------
  ' Exit if no entries found
  ' ------------------------------------------

  If p_lngEntriesRead < 1 Then
    Exit Function
  End If

  ' ------------------------------------------
  ' Redim the type array to hold this info
  ' ------------------------------------------

  ReDim p_atypUserInfo(0 To p_lngEntriesRead - 1)
 
  ' ------------------------------------------
  ' Copy the pointer to the buffer into the
  ' type array
  ' ------------------------------------------

  CopyMem p_atypUserInfo(0), ByVal p_lngPtrBuffer, Len(p_atypUserInfo(0)) * p_lngEntriesRead

  ' ------------------------------------------
  ' Fill-in the info needed to call the
  ' Add() method
  ' NOTE: We will always have +1 open pipe,
  ' since in making this call we create
  ' a pipe, "\PIPE\srvsvc"
  ' ------------------------------------------
  
  Pass = 0
  For p_lngLoop = 0 To p_lngEntriesRead - 1
    p_strUserName = PointerToUnicodeStr(p_atypUserInfo(p_lngLoop).Name)
    p_abytUserName = p_strUserName & Chr(0)
    p_lngRtn = NetUserGetInfo(p_abytServerName(0), p_abytUserName(0), 2, p_lngPtrUserInfoBuf)

    If p_lngRtn <> 0 Then
      Err.Raise p_lngRtn, "EnumUsersLoggedOn", "Had an error with EnumLoggedOnUsers"
      Exit Function
    End If
 
    CopyMem p_typUserInfo, ByVal p_lngPtrUserInfoBuf, Len(p_typUserInfo)
 
    p_lngLastLogon = p_typUserInfo.usri2_last_logon   ' The last time the user logged on
    p_lngLastLogoff = p_typUserInfo.usri2_last_logoff ' The last time the user logged off
 
    If (p_lngLastLogoff < p_lngLastLogon) Then        ' If the logoff time is less than the
      ReDim Preserve LoggedOnUsers(0 To Pass)
      LoggedOnUsers(Pass) = p_strUserName        ' logon time, then they're still logged on
      Pass = Pass + 1
    End If

    If p_lngPtrUserInfoBuf <> 0 Then NetAPIBufferFree p_lngPtrUserInfoBuf
  Next p_lngLoop
  
  ' ------------------------------------------
  ' Clean-up the buffer
  ' ------------------------------------------
  If p_lngPtrBuffer <> 0 Then NetAPIBufferFree p_lngPtrBuffer
  
  EnumUsersLoggedOn = LoggedOnUsers()
End Function



Private Function PointerToUnicodeStr(lpUnicodeStr As Long) As String
  On Error Resume Next ' Don't accept an error here
  Dim Buffer() As Byte
  Dim nLen As Long
  If lpUnicodeStr Then
    nLen = StrLenW(lpUnicodeStr) * 2
    If nLen Then
      ReDim Buffer(0 To (nLen - 1)) As Byte

      ' ------------------------------------
      ' Copy the pointer to the buffer into
      ' the type array
      ' ------------------------------------
      CopyMem Buffer(0), ByVal lpUnicodeStr, nLen
      PointerToUnicodeStr = Buffer
    End If
  End If
End Function

' Use on NT Only
Public Function ChangePassword(ByVal Domain As String, ByVal UserName As String, ByVal OldPassword As String, ByVal NewPassword As String) As Boolean
  Dim sServer As String, sUser As String
  Dim sNewPass As String, sOldPass As String
  Dim UI1003 As USER_INFO_1003
  Dim dwLevel As Long
  Dim lRet As String
  Dim sNew As String

  ' StrConv Functions are necessary since VB will perform
  ' UNICODE/ANSI translation before passing strings to the
  ' NETAPI functions

  sUser = StrConv(UserName, vbUnicode)
  sNewPass = StrConv(NewPassword, vbUnicode)
  'See if this is Domain or Computer referenced

  If Left(Domain, 2) = "\\" Then
    sServer = StrConv(Domain, vbUnicode)
  Else
    ' Domain was referenced, get the Primary Domain Controller
    sServer = StrConv(GetPDCName(Domain), vbUnicode)
  End If

  If OldPassword = "" Then
    ' Administrative over-ride of existing password.
    ' Does not require old password
    dwLevel = 1003
    sNew = NewPassword
    UI1003.usri1003_password = StrPtr(sNew)
    lRet = NetUserSetInfo(sServer, sUser, dwLevel, UI1003, 0&)
  Else
    ' Set the Old Password and attempt to change the user's password
    sOldPass = StrConv(OldPassword, vbUnicode)
    lRet = NetUserChangePassword(sServer, sUser, sOldPass, sNewPass)
  End If
  
  If lRet <> 0 Then
    ChangePassword = False
  Else
    ChangePassword = True
  End If
End Function

Public Function GetUserInfo(Server As String, UserName As String) As USER_INFO
  Dim bUsername() As Byte
  Dim bServername() As Byte
  Dim usrapi As USER_INFO_10_API
  Dim buff As Long

  If Server <> "" And Left(Server, 2) <> "\\" Then Server = "\\" & Server

  bUsername = UserName & Chr$(0)
  bServername = Server & Chr$(0)
   
  If NetUserGetInfo(bServername(0), bUsername(0), 10, buff) = ERROR_SUCCESS Then
    'copy the data from buff into the
    'API user_10 structure
    CopyMem usrapi, ByVal buff, Len(usrapi)

    'extract each member and return
    'as members of the UDT
    GetUserInfo.Name = GetPointerToByteStringW(usrapi.Name)
    GetUserInfo.FullName = GetPointerToByteStringW(usrapi.FullName)
    GetUserInfo.Comment = GetPointerToByteStringW(usrapi.Comment)
    GetUserInfo.UserComment = GetPointerToByteStringW(usrapi.UsrComment)

    NetAPIBufferFree buff
  Else
    Err.Raise Err.LastDllError, "GetUserInfo"
  End If

End Function

Private Function GetPointerToByteStringW(lpString As Long) As String
  Dim buff() As Byte
  Dim nSize As Long

  If lpString Then
    'its Unicode, so mult. by 2
    nSize = StrLenW(lpString) * 2
    If nSize Then
      ReDim buff(0 To (nSize - 1)) As Byte
      CopyMem buff(0), ByVal lpString, nSize
      GetPointerToByteStringW = buff
    End If
  End If
End Function

