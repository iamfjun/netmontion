Attribute VB_Name = "Module2"
Public cnn As New ADODB.Connection
Public cnn_1 As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rs_1 As New ADODB.Recordset
Public strcnn As String
Public strsq As String
Public strs As String
Public strSql As String
Public strSql_1 As String
Option Explicit
Public Const IP_STATUS_BASE = 11000
Public Const IP_SUCCESS = 0
Public Const IP_BUF_TOO_SMALL = (11000 + 1)
Public Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Public Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Public Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Public Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Public Const IP_NO_RESOURCES = (11000 + 6)
Public Const IP_BAD_OPTION = (11000 + 7)
Public Const IP_HW_ERROR = (11000 + 8)
Public Const IP_PACKET_TOO_BIG = (11000 + 9)
Public Const IP_REQ_TIMED_OUT = (11000 + 10)
Public Const IP_BAD_REQ = (11000 + 11)
Public Const IP_BAD_ROUTE = (11000 + 12)
Public Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Public Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Public Const IP_PARAM_PROBLEM = (11000 + 15)
Public Const IP_SOURCE_QUENCH = (11000 + 16)
Public Const IP_OPTION_TOO_BIG = (11000 + 17)
Public Const IP_BAD_DESTINATION = (11000 + 18)
Public Const IP_ADDR_DELETED = (11000 + 19)
Public Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Public Const IP_MTU_CHANGE = (11000 + 21)
Public Const IP_UNLOAD = (11000 + 22)
Public Const IP_ADDR_ADDED = (11000 + 23)
Public Const IP_GENERAL_FAILURE = (11000 + 50)
Public Const MAX_IP_STATUS = 11000 + 50
Public Const IP_PENDING = (11000 + 255)
Public Const PING_TIMEOUT = 200
Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1

Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128

Public Type ICMP_OPTIONS
Ttl             As Byte
Tos             As Byte
Flags           As Byte
OptionsSize     As Byte
OptionsData     As Long
End Type


Dim ICMPOPT As ICMP_OPTIONS

Public Type ICMP_ECHO_REPLY
Address         As Long
status          As Long
RoundTripTime   As Long
DataSize        As Integer
Reserved        As Integer
DataPointer     As Long
Options         As ICMP_OPTIONS
data            As String * 250
End Type

Public Type HOSTENT
hName           As Long
hAliases        As Long
hAddrType       As Integer
hLen            As Integer
hAddrList       As Long
End Type

Public Type WSADATA
wVersion        As Integer
wHighVersion    As Integer
szDescription(0 To MAX_WSADescription) As Byte
szSystemStatus(0 To MAX_WSASYSStatus) As Byte
wMaxSockets     As Integer
wMaxUDPDG       As Integer
dwVendorInfo    As Long
End Type


Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

Public Declare Function IcmpCloseHandle Lib "icmp.dll" _
(ByVal IcmpHandle As Long) As Long
   
Public Declare Function IcmpSendEcho Lib "icmp.dll" _
(ByVal IcmpHandle As Long, _
ByVal DestinationAddress As Long, _
ByVal RequestData As String, _
ByVal RequestSize As Integer, _
ByVal RequestOptions As Long, _
ReplyBuffer As ICMP_ECHO_REPLY, _
ByVal ReplySize As Long, _
ByVal Timeout As Long) As Long
    
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long

Public Declare Function WSAStartup Lib "WSOCK32.DLL" _
(ByVal wVersionRequired As Long, _
lpWSADATA As WSADATA) As Long
    
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Public Declare Function gethostname Lib "WSOCK32.DLL" _
(ByVal szHost As String, _
ByVal dwHostLen As Long) As Long
    
Public Declare Function gethostbyname Lib "WSOCK32.DLL" _
(ByVal szHost As String) As Long
   
Public Declare Sub RtlMoveMemory Lib "kernel32" _
(hpvDest As Any, _
ByVal hpvSource As Long, _
ByVal cbCopy As Long)

Public Function GetStatusCode(status As Long) As String

Dim msg As String

'Message Boxes
Select Case status
Case IP_SUCCESS:               msg = "IP Successful"
Case IP_BUF_TOO_SMALL:         msg = "IP Buffer Too Small"
Case IP_DEST_NET_UNREACHABLE:  msg = "IP Destination Net Unreachable"
Case IP_DEST_HOST_UNREACHABLE: msg = "IP Destination Host Unreachable"
Case IP_DEST_PROT_UNREACHABLE: msg = "IP Destination Protocol Unreachable"
Case IP_DEST_PORT_UNREACHABLE: msg = "IP Destination Port Unreachable"
Case IP_NO_RESOURCES:          msg = "IP No Resources"
Case IP_BAD_OPTION:            msg = "IP Bad Option"
Case IP_HW_ERROR:              msg = "IP Hw Error"
Case IP_PACKET_TOO_BIG:        msg = "IP Packet Too Big"
Case IP_REQ_TIMED_OUT:         msg = "IP currently timed out"
Case IP_BAD_REQ:               msg = "IP Bad Request"
Case IP_BAD_ROUTE:             msg = "IP Bad Route"
Case IP_TTL_EXPIRED_TRANSIT:   msg = "IP ttl expired transit"
Case IP_TTL_EXPIRED_REASSEM:   msg = "IP ttl expired reassem"
Case IP_PARAM_PROBLEM:         msg = "IP Parameter Problem"
Case IP_SOURCE_QUENCH:         msg = "IP Source Quench"
Case IP_OPTION_TOO_BIG:        msg = "IP Option too Big"
Case IP_BAD_DESTINATION:       msg = "IP Bad Destination"
Case IP_ADDR_DELETED:          msg = "IP addr deleted"
Case IP_SPEC_MTU_CHANGE:       msg = "IP Spec Mtu Change"
Case IP_MTU_CHANGE:            msg = "IP Mtu Change"
Case IP_UNLOAD:                msg = "IP Unload"
Case IP_ADDR_ADDED:            msg = "IP Address Added"
Case IP_GENERAL_FAILURE:       msg = "IP General Failure"
Case IP_PENDING:               msg = "IP Pending"
Case PING_TIMEOUT:             msg = "Ping Timeout"
Case Else:                     msg = "Unknown Message Returned!"
End Select

GetStatusCode = CStr(status) & "   [ " & msg & " ]"
   
End Function

Public Function HiByte(ByVal wParam As Integer)
HiByte = wParam \ &H100 And &HFF&
End Function

Public Function LoByte(ByVal wParam As Integer)
LoByte = wParam And &HFF&
End Function

Public Function Ping(szAddress As String, ECHO As ICMP_ECHO_REPLY) As Long
'Download by http://www.codefans.net
Dim hPort        As Long
Dim dwAddress    As Long
Dim sDataToSend  As String
Dim iOpt         As Long
   
sDataToSend = "Echo This"
dwAddress = AddressStringToLong(szAddress)
   
Call SocketsInitialize
hPort = IcmpCreateFile()
   
If IcmpSendEcho(hPort, _
dwAddress, _
sDataToSend, _
Len(sDataToSend), _
0, _
ECHO, _
Len(ECHO), _
PING_TIMEOUT) Then
Ping = ECHO.RoundTripTime
Else: Ping = ECHO.status * -1
End If
                       
Call IcmpCloseHandle(hPort)
Call SocketsCleanup
   
End Function
Function AddressStringToLong(ByVal tmp As String) As Long
Dim i As Integer
Dim parts(1 To 4) As String
   
i = 0

While InStr(tmp, ".") > 0
i = i + 1
parts(i) = Mid(tmp, 1, InStr(tmp, ".") - 1)
tmp = Mid(tmp, InStr(tmp, ".") + 1)
Wend
   
i = i + 1
parts(i) = tmp
   
If i <> 4 Then
AddressStringToLong = 0
Exit Function
End If

AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & _
Right("00" & Hex(parts(3)), 2) & _
Right("00" & Hex(parts(2)), 2) & _
Right("00" & Hex(parts(1)), 2))
   
End Function
Public Function SocketsCleanup() As Boolean
    Dim X As Long
    X = WSACleanup()
    If X <> 0 Then
    MsgBox "Windows Sockets error " & Trim$(Str$(X)) & _
    " occurred in Cleanup.", vbExclamation
    SocketsCleanup = False
    Else
    SocketsCleanup = True
    End If
End Function
Public Function SocketsInitialize() As Boolean
Dim WSAD As WSADATA
    Dim X As Integer
    Dim szLoByte As String, szHiByte As String, szBuf As String
    X = WSAStartup(WS_VERSION_REQD, WSAD)
        
    If X <> 0 Then
    MsgBox "Windows Sockets for 32 bit Windows " & _
    "environments is not successfully responding."
    SocketsInitialize = False
    Exit Function
    End If
        
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
    (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
    HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
    
    szHiByte = Trim$(Str$(HiByte(WSAD.wVersion)))
    szLoByte = Trim$(Str$(LoByte(WSAD.wVersion)))
    szBuf = "Windows Sockets Version " & szLoByte & "." & szHiByte
    szBuf = szBuf & " is not supported by Windows " & _
    "Sockets for 32 bit Windows environments."
    MsgBox szBuf, vbExclamation
    SocketsInitialize = False
    Exit Function
    End If
        
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
    szBuf = "This application requires a minimum of " & _
    Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
    MsgBox szBuf, vbExclamation
    SocketsInitialize = False
    Exit Function
    End If
        
    SocketsInitialize = True
End Function
Public Function connect8(strSql As String) As ADODB.Recordset
'�������ݿ�����'
Set cnn = Nothing
Set rs = Nothing
'�뱾�ص�Access���ݿ�����
cnn.ConnectionString = "Provider=microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ping_data.mdb"
cnn.Open
rs.CursorLocation = adUseClient
rs.Open strSql, cnn, 3, 3
Set connect8 = rs
 End Function

Public Function connect9(strSql_1 As String) As ADODB.Recordset
'�뱾�ص�Access���ݿ�����
cnn_1.ConnectionString = "Provider=microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ping_data.mdb"
cnn_1.Open
rs_1.CursorLocation = adUseClient
rs_1.Open strSql_1, cnn_1, 3, 3
Set connect9 = rs_1
 End Function

