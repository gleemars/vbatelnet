Attribute VB_Name = "tlnt"
'*********************************************
'Project:vbatelnet 1.0
'File Name:tlnt.bas
'Home Page:http://gleemars.com
'Copyright (C) 2013 Yan Xingjian(gleemars@gmail.com or yanxingjian@139.com)
'All rights reserved.
'The contents of this file may be used under the terms of the
'LGPL license (the "GNU LIBRARY GENERAL PUBLIC LICENSE").
'*********************************************
'*********************************************
' 项目名：vbatelnet
' 文件名：tlnt.bas
' 主页：http://gleemars.com
' 版权所有 (C) 2013 颜兴建
' EMail:yanxingjian@139.com or gleemars@gmail.com
' 版权所有
' 此文件内容可在LGPL授权下使用。
'*********************************************

'Const A As String = Chr$(10)
Dim CR  As String
Dim LF  As String


Public Function tlnt_open(ByVal hst As String, ByVal prt As Long) As Long
    Dim wsa_data As WSADataType
    Dim sck_addr As sockaddr
    Dim rtn As Long
    Dim sck_nm As Long
    rtn = WSAStartup(&H101, wsa_data)
    sck_nm = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    sck_addr.sin_family = AF_INET
    'SocketBuffer.sin_port = 80
    'SocketBuffer.sin_port = htons(511)
    sck_addr.sin_port = htons(prt)
    sck_addr.sin_addr = inet_addr(hst)
    sck_addr.sin_zero = String$(8, 0)
    rtn = connect(sck_nm, sck_addr, 16)
    tlnt_open = sck_nm
End Function



Public Function tlnt_close(ByVal sck_nm As Long) As Long
    Dim rtn As Long
    rtn = closesocket(sck_nm)
    rtn = WSACleanup()
    tlnt_close = rtn
End Function



Public Function tlnt_puts(ByVal sck_nm As Long, ByVal snd As Variant) As Long
    Dim lft As Long, idx As Long, rtn As Long, i As Long
    Dim snd_byte_lft() As Byte
    
    'Dim snd_byte() As Byte
    'snd_byte = StrConv(snd, vbFromUnicode)
    idx = 0
    lft = UBound(snd) - LBound(snd) + 1 'Len(snd)
    
    Do While lft > 0
        ReDim snd_byte_lft(lft - 1) As Byte
        'CopyMemory VarPtr(snd_byte_lft(0)), VarPtr(snd(idx)), lft
        For i = 0 To (lft - 1)
            snd_byte_lft(i) = snd(idx + i)
        Next
        
        rtn = send(sck_nm, snd_byte_lft(0), lft, 0)
        If rtn = SOCKET_ERROR Then
            tlnt_puts = 1
            Exit Function
        End If
        lft = lft - rtn
        idx = idx + rtn
    Loop
    tlnt_puts = 0
End Function






Public Function tlnt_gets(ByVal sck_nm As Long) As String
Dim IAC As Byte
Dim DONT As Byte
Dim DOO As Byte
Dim WONT As Byte
Dim WILL  As Byte
Dim SB  As Byte
Dim AYT  As Byte
Dim SE  As Byte

Dim BGN  As Byte
Dim ENDD  As Byte


Dim OPT_BINARY  As Byte
Dim OPT_ECHO  As Byte
Dim OPT_SGA  As Byte


Dim rcv_sz As Long, idx As Long, tmp As Long, i As Long
Dim rtn_str As String, rcv_bff(2047) As Byte, snd_bff() As Byte, rcv_str As String, snd_str As String, stt As String
    
IAC = 255 'Chr(255)
DONT = 254 'Chr(254)
DOO = 253 'Chr(253)
WONT = 252 'Chr(252)
WILL = 251 'Chr(251)
SB = 250 'Chr(250)
AYT = 246 'Chr(246)
SE = 240 'Chr(240)

BGN = 0 'Chr(0)
ENDD = 1 'Chr(1)
CR = 13 'Chr(13)
LF = 10 'Chr(10)

OPT_BINARY = 0 'Chr(0)
OPT_ECHO = 1 'Chr(1)
OPT_SGA = 3 'Chr(3)
    
    ReDim snd_bff(0) As Byte
    rtn_str = ""
    snd_str = ""
    rcv_sz = recv(sck_nm, VarPtr(rcv_bff(0)), 2048, 0)
    'snd_bff = Array()
    'rcv_str = StrConv(rcv_bff, vbUnicode)
    'Debug.Print rcv_sz
    'Debug.Print rtn_str
    idx = 0
    stt = BGN
    Do While rcv_sz > 0
        Dim c As Byte
        c = rcv_bff(idx) 'Mid$(rcv_str, idx, 1)
        'Debug.Print rcv_sz
        'Debug.Print c = IAC
        Select Case stt
        Case BGN, ENDD
            If c = IAC Then
                stt = IAC
            Else
                If c <> Chr(0) Then
                    rtn_str = rtn_str & Chr(c)
                End If
            End If
        Case IAC
            Select Case c
            Case IAC, DOO, DONT, WILL, WONT, SB
                stt = c
            Case SE
                stt = ENDD
            Case AYT
                'snd_str = snd_str & "Y" & CR & LF
                Call bytes_add(snd_bff, "Y")
                Call bytes_add(snd_bff, CR)
                Call bytes_add(snd_bff, LF)
            End Select
        Case DOO
            If c = OPT_BINARY Then
                stt = ENDD
            Else
                'snd_str = snd_str & IAC & WONT & c
                Call bytes_add(snd_bff, IAC)
                Call bytes_add(snd_bff, WONT)
                Call bytes_add(snd_bff, c)
                stt = ENDD
            End If
        Case DONT
            'snd_str = snd_str & IAC & WONT & c
            Call bytes_add(snd_bff, IAC)
            Call bytes_add(snd_bff, WONT)
            Call bytes_add(snd_bff, c)
            stt = ENDD
        Case WILL
            Select Case c
            Case OPT_ECHO, OPT_SGA
                'snd_str = snd_str & IAC & DOO & c
                Call bytes_add(snd_bff, IAC)
                Call bytes_add(snd_bff, DOO)
                Call bytes_add(snd_bff, c)
                stt = ENDD
            Case Else
                'snd_str = snd_str & IAC & WONT & c
                Call bytes_add(snd_bff, IAC)
                Call bytes_add(snd_bff, WONT)
                Call bytes_add(snd_bff, c)
                stt = ENDD
            End Select
        Case WONT
            'snd_str = snd_str & IAC & DONT & c
            Call bytes_add(snd_bff, IAC)
            Call bytes_add(snd_bff, DONT)
            Call bytes_add(snd_bff, c)
            stt = ENDD
        Case SB
            Select Case c
            Case IAC, SB
                stt = c
            End Select
        
        End Select
        rcv_sz = rcv_sz - 1
        idx = idx + 1
    Loop
    'ReDim Preserve snd_bff(UBound(snd_bff) - 1)
    'snd_bff(UBound(snd_bff)) = 0
    If UBound(snd_bff) > 0 Then
        Dim snd_bff_nw() As Byte
        ReDim snd_bff_nw(UBound(snd_bff) - 1) As Byte
        For i = 0 To UBound(snd_bff) - 1
            snd_bff_nw(i) = snd_bff(i)
        Next
        tmp = tlnt_puts(sck_nm, snd_bff_nw)
    End If
    tlnt_gets = rtn_str
End Function


Public Function tlnt_wait(ByVal sck_nm As Long, ByVal str As String) As String
    Dim rtn As String
    Dim fnded As String
    fnded = ""
    rtn = tlnt_gets(sck_nm)
    
    Do While 1
        'Debug.Print rtn
        If Len(rtn) > 20 Then
            fnded = Mid$(rtn, Len(rtn) - 20, 20)
        Else
            fnded = rtn
        End If
        If InStr(fnded, str) > 0 Then
            tlnt_wait = rtn
            Exit Function
        End If
        rtn = rtn & tlnt_gets(sck_nm)
    Loop
    
End Function

Public Sub bytes_add(a As Variant, ByVal b As Byte)
   
    'Dim r() As Byte
    'Dim i As Long
    ' ...assign data to ab1 and ab2...
    'r = a
    ReDim Preserve a(UBound(a) + 1)
    'For i = 0 To UBound(b)
    '    r(i + UBound(a)) = b(i)
    'Next
    a(UBound(a) - 1) = b
    'bytes_add = r
End Sub



Public Function tlnt_command(ByVal sck_nm As Long, ByVal cmd_str As String)
    Dim cmd() As Byte
    Dim rtn As Long
    cmd = StrConv(cmd_str, vbFromUnicode)
    rtn = tlnt_puts(sck_nm, cmd)
   
End Function

Public Sub main()
    Dim tlnt As Long, rtn As Long
    Dim s As String
    Dim cmd() As Byte
    
    tlnt = tlnt_open("127.0.0.1", 23)
    's = tlnt_gets(tlnt)
    'Debug.Print s
    s = tlnt_wait(tlnt, "Username:")
    Debug.Print s
    cmd = StrConv("xxxxxx" & vbCr, vbFromUnicode)
    rtn = tlnt_puts(tlnt, cmd)
    s = tlnt_wait(tlnt, "Password:")
    Debug.Print s
    cmd = StrConv("xxxxxx" & vbCr, vbFromUnicode)
    rtn = tlnt_puts(tlnt, cmd)
    s = tlnt_wait(tlnt, "rt1>")
    Debug.Print s
    's = tlnt_gets(tlnt)
    'Debug.Print s
    cmd = StrConv("xxxxxx" & vbCr, vbFromUnicode)
    rtn = tlnt_puts(tlnt, cmd)
    s = tlnt_wait(tlnt, "ENTER USERNAME < ")
    Debug.Print s
    cmd = StrConv("xxxxxx" & vbCr, vbFromUnicode)
    rtn = tlnt_puts(tlnt, cmd)
    s = tlnt_wait(tlnt, "ENTER PASSWORD < ")
    Debug.Print s
    rtn = tlnt_close(tlnt)
End Sub

