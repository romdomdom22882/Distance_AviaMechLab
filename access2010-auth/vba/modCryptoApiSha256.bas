Attribute VB_Name = "modCryptoApiSha256"
Option Compare Database
Option Explicit

' SHA-256 через Windows CryptoAPI (advapi32)
' Работает в Access 2010 (VBA7) для 32/64-bit.

#If VBA7 Then
    Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
        ByRef phProv As LongPtr, ByVal pszContainer As String, ByVal pszProvider As String, _
        ByVal dwProvType As Long, ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" ( _
        ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" ( _
        ByVal hProv As LongPtr, ByVal Algid As Long, ByVal hKey As LongPtr, ByVal dwFlags As Long, _
        ByRef phHash As LongPtr) As Long

    Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" ( _
        ByVal hHash As LongPtr, ByRef pbData As Byte, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" ( _
        ByVal hHash As LongPtr, ByVal dwParam As Long, ByRef pbData As Byte, ByRef pdwDataLen As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" ( _
        ByVal hHash As LongPtr) As Long

    Private Declare PtrSafe Function CryptGenRandom Lib "advapi32.dll" ( _
        ByVal hProv As LongPtr, ByVal dwLen As Long, ByRef pbBuffer As Byte) As Long

    Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef Destination As Any, ByVal Source As LongPtr, ByVal Length As LongPtr)
#Else
    Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
        ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, _
        ByVal dwProvType As Long, ByVal dwFlags As Long) As Long

    Private Declare Function CryptReleaseContext Lib "advapi32.dll" ( _
        ByVal hProv As Long, ByVal dwFlags As Long) As Long

    Private Declare Function CryptCreateHash Lib "advapi32.dll" ( _
        ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, _
        ByRef phHash As Long) As Long

    Private Declare Function CryptHashData Lib "advapi32.dll" ( _
        ByVal hHash As Long, ByRef pbData As Byte, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long

    Private Declare Function CryptGetHashParam Lib "advapi32.dll" ( _
        ByVal hHash As Long, ByVal dwParam As Long, ByRef pbData As Byte, ByRef pdwDataLen As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function CryptDestroyHash Lib "advapi32.dll" ( _
        ByVal hHash As Long) As Long

    Private Declare Function CryptGenRandom Lib "advapi32.dll" ( _
        ByVal hProv As Long, ByVal dwLen As Long, ByRef pbBuffer As Byte) As Long

    Private Declare Sub RtlMoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef Destination As Any, ByVal Source As Long, ByVal Length As Long)
#End If

Private Const PROV_RSA_AES As Long = 24
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000

Private Const CALG_SHA_256 As Long = &H800C
Private Const HP_HASHVAL As Long = &H2

Public Function Crypto_RandomSaltHex(Optional ByVal byteCount As Long = 16) As String
    Dim hProv As LongPtr
    Dim ok As Long
    Dim b() As Byte

    If byteCount <= 0 Then byteCount = 16
    ReDim b(0 To byteCount - 1) As Byte

    ok = CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT)
    If ok = 0 Then
        Crypto_RandomSaltHex = vbNullString
        Exit Function
    End If

    ok = CryptGenRandom(hProv, byteCount, b(0))
    CryptReleaseContext hProv, 0

    If ok = 0 Then
        Crypto_RandomSaltHex = vbNullString
        Exit Function
    End If

    Crypto_RandomSaltHex = BytesToHex(b)
End Function

Public Function Crypto_Sha256Hex(ByVal inputText As String) As String
    Dim hProv As LongPtr
    Dim hHash As LongPtr
    Dim ok As Long
    Dim data() As Byte
    Dim hash() As Byte
    Dim hashLen As Long

    data = StringToUtf16LeBytes(inputText)

    ok = CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT)
    If ok = 0 Then GoTo EH

    ok = CryptCreateHash(hProv, CALG_SHA_256, 0, 0, hHash)
    If ok = 0 Then GoTo EH

    If (UBound(data) >= 0) Then
        ok = CryptHashData(hHash, data(0), UBound(data) + 1, 0)
        If ok = 0 Then GoTo EH
    End If

    hashLen = 32
    ReDim hash(0 To hashLen - 1) As Byte
    ok = CryptGetHashParam(hHash, HP_HASHVAL, hash(0), hashLen, 0)
    If ok = 0 Then GoTo EH

    Crypto_Sha256Hex = BytesToHex(hash)

CleanUp:
    On Error Resume Next
    If hHash <> 0 Then CryptDestroyHash hHash
    If hProv <> 0 Then CryptReleaseContext hProv, 0
    Exit Function

EH:
    Crypto_Sha256Hex = vbNullString
    Resume CleanUp
End Function

Private Function BytesToHex(ByRef b() As Byte) As String
    Dim i As Long
    Dim s As String

    If (Not Not b) = 0 Then
        BytesToHex = vbNullString
        Exit Function
    End If

    s = String$((UBound(b) - LBound(b) + 1) * 2, "0")
    For i = LBound(b) To UBound(b)
        Mid$(s, (i - LBound(b)) * 2 + 1, 2) = Right$("0" & Hex$(b(i)), 2)
    Next i
    BytesToHex = LCase$(s)
End Function

Private Function StringToUtf16LeBytes(ByVal s As String) As Byte()
    Dim cb As Long
    Dim b() As Byte

    cb = LenB(s) ' bytes в UTF-16LE (VBA string)
    If cb = 0 Then
        ReDim b(0 To -1) As Byte
        StringToUtf16LeBytes = b
        Exit Function
    End If

    ReDim b(0 To cb - 1) As Byte
    RtlMoveMemory b(0), StrPtr(s), cb
    StringToUtf16LeBytes = b
End Function

