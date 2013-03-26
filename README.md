vbatelnet
=========

a tiny telnet library for vba.

WEBSITE
==========
[gleemars](http://gleemars.com)

API
=========
* Public Function tlnt_open(ByVal hst As String, ByVal prt As Long) As Long
* Public Function tlnt_gets(ByVal sck_nm As Long) As String
* Public Function tlnt_puts(ByVal sck_nm As Long, ByVal snd As Variant) As Long
* Public Function tlnt_close(ByVal sck_nm As Long) As Long
* Public Function tlnt_command(ByVal sck_nm As Long, ByVal cmd_str As String)
* Public Function tlnt_wait(ByVal sck_nm As Long, ByVal str As String) As String

EXAMPLE
==========
```vbnet
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
```