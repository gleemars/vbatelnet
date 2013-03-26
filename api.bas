Attribute VB_Name = "api"
'*********************************************
'Project:vbatelnet
'File Name:api.bas
'Copyright (C) 2013 Yan Xingjian(gleemars@gmail.com or yanxingjian@139.com)
'All rights reserved.
'The contents of this file may be used under the terms of the
'LGPL license (the "GNU LIBRARY GENERAL PUBLIC LICENSE").
'*********************************************
'*********************************************
' 项目名：vbatelnet 1.0
' 文件名：api.bas
' 版权所有 (C) 2013 颜兴建
' EMail:yanxingjian@139.com or gleemars@gmail.com
' 版权所有
' 此文件内容可在LGPL授权下使用。
'*********************************************

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Public Const hostent_size = 16
Public Const INADDR_NONE = &HFFFFFFFF
Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type
Public Const SOCKADDR_SIZE = 16
Public Const WSA_DESCRIPTIONLEN = 256
Public Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1
Public Const WSA_SYS_STATUS_LEN = 128
Public Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1
Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Public Const IPPROTO_TCP = 6
Public Const SOCKET_ERROR = -1
Public Const SOCK_STREAM = 1
Public Const AF_INET = 2
Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Any, ByVal Source As Any, ByVal Length As Integer)
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Public Const FD_READ = &H1&
Public Const FD_ACCEPT = &H8&
Public Const FD_CONNECT = &H10&
Public Const FD_CLOSE = &H20&
Public Declare Function accept Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, addrlen As Long) As Long
Public Declare Function bind Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
Public Declare Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long
Public Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
Public Declare Function listen Lib "wsock32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
Public Declare Function recv Lib "wsock32.dll" (ByVal s As Long, ByVal buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function send Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function socket Lib "wsock32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal host_name As String) As Long
Public Declare Function connect Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
Public Declare Function gethostname Lib "wsock32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long

Public Const WSA_NoName = "Unknown"
Public Function AddrToIP(ByVal AddrOrIP$) As String
    AddrToIP$ = GetAscIP(GetHostByNameAlias(AddrOrIP$))
End Function
Public Function GetAscIP(ByVal inn As Long) As String
        Dim nStr&
    Dim lpStr&
    Dim retString$
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr Then
        nStr = lstrlen(lpStr)
        If nStr > 32 Then nStr = 32
        MemCopy ByVal retString, ByVal lpStr, nStr
        retString = Left(retString, nStr)
        GetAscIP = retString
    Else
        GetAscIP = "255.255.255.255"
    End If
End Function
Public Function GetHostByNameAlias(ByVal HostName$) As Long
    Dim phe&
    Dim heDestHost As HostEnt
    Dim addrList&
    Dim retIP&
    retIP = inet_addr(HostName$)
    If retIP = INADDR_NONE Then
        phe = gethostbyname(HostName$)
        If phe <> 0 Then
            MemCopy heDestHost, ByVal phe, hostent_size
            MemCopy addrList, ByVal heDestHost.h_addr_list, 4
            MemCopy retIP, ByVal addrList, heDestHost.h_length
        Else
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
End Function
Public Function GetLocalHostName() As String
    Dim sName$
    sName = String(256, 0)
    If gethostname(sName, 256) Then
        sName = WSA_NoName
    Else
        If InStr(sName, Chr(0)) Then
            sName = Left(sName, InStr(sName, Chr(0)) - 1)
        End If
    End If
    GetLocalHostName = sName
End Function
