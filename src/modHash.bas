Attribute VB_Name = "modHash"
'==============================================================================
' modHash
' ----------------------------------------------------------------------------
' File-content fingerprinting for the .fmap audit trail. Used to detect when
' a saved mapping configuration no longer matches the source files on disk.
'
' Hash function: FNV-1a 32-bit. Chosen over SHA-1 to avoid fragile CryptoAPI
' Declares across 32/64-bit Office. FNV-1a is not cryptographic but is plenty
' for drift detection -- combined with file size and modtime, the chance of a
' false "unchanged" verdict on a modified .mpp is negligible.
'
' VBA's Long is signed 32-bit; FNV-1a needs unsigned 32-bit arithmetic. We
' use the Decimal Variant subtype for intermediate values, which carries 96
' bits of integer precision -- plenty for the 32x32->64-bit multiplications
' before the modulo-2^32 reduction.
'==============================================================================
Option Explicit

Private Const FNV_OFFSET_BASIS As String = "2166136261"   ' 0x811C9DC5
Private Const FNV_PRIME        As Long = 16777619         ' 0x01000193

'==============================================================================
' Fnv1a32
' ----------------------------------------------------------------------------
' Computes FNV-1a 32-bit over a Byte array. Returns an 8-character uppercase
' hex string (e.g. "9F2C8B11").
'
' Algorithm:
'   hash = 0x811C9DC5
'   for each byte b:
'       hash = hash XOR b
'       hash = (hash * 0x01000193) mod 2^32
'==============================================================================
Public Function Fnv1a32(ByRef bytes() As Byte) As String
    Dim h As Variant, pow32 As Variant
    h = CDec(FNV_OFFSET_BASIS)
    pow32 = CDec("4294967296")           ' 2^32

    Dim lo As Long, hi As Variant
    Dim i As Long
    For i = LBound(bytes) To UBound(bytes)
        ' XOR the low byte of h with the next input byte
        hi = h \ 256
        lo = CLng(h - hi * 256)          ' h mod 256, guaranteed in 0..255
        lo = lo Xor CLng(bytes(i))
        h = hi * 256 + lo

        ' Multiply by prime and reduce mod 2^32
        h = h * FNV_PRIME
        h = h - (h \ pow32) * pow32
    Next i

    Fnv1a32 = ToHex8(h)
End Function

'==============================================================================
' Fnv1a32OfFileHead
' ----------------------------------------------------------------------------
' Reads the first nBytes of a file and returns its FNV-1a 32-bit hash. Files
' smaller than nBytes are hashed in their entirety. Empty files return the
' FNV-1a offset basis ("811C9DC5"), which is the canonical empty hash.
'
' Errors are swallowed and reported as "ERR<number>" so a single
' inaccessible file doesn't crash a fingerprint pass.
'==============================================================================
Public Function Fnv1a32OfFileHead(ByVal path As String, _
                                  Optional ByVal nBytes As Long = FINGERPRINT_HEAD_BYTES) _
                                  As String
    Dim fnum As Integer
    fnum = 0
    On Error GoTo Fail

    Dim actualLen As Long
    actualLen = FileLen(path)
    If nBytes > actualLen Then nBytes = actualLen
    If nBytes <= 0 Then
        Fnv1a32OfFileHead = "811C9DC5"   ' FNV-1a offset basis
        Exit Function
    End If

    Dim b() As Byte
    ReDim b(0 To nBytes - 1)

    fnum = FreeFile
    Open path For Binary Access Read As #fnum
    Get #fnum, 1, b
    Close #fnum
    fnum = 0

    Fnv1a32OfFileHead = Fnv1a32(b)
    Exit Function

Fail:
    Dim n As Long
    n = Err.Number
    On Error Resume Next
    If fnum <> 0 Then Close #fnum
    On Error GoTo 0
    Err.Clear
    Fnv1a32OfFileHead = "ERR" & n
End Function

'==============================================================================
' FileFingerprint
' ----------------------------------------------------------------------------
' Drift-detection fingerprint for a file. Format:
'     "<size>_<modtime>_<headhash>"
' Example:
'     "4827392_20260422091400_9F2C8B11"
'
' Used in .fmap persistence so we can warn the user when a saved mapping
' configuration was built against a different version of a source file.
'==============================================================================
Public Function FileFingerprint(ByVal path As String) As String
    On Error GoTo Fail
    Dim sz As Long, mt As Date
    sz = FileLen(path)
    mt = FileDateTime(path)
    FileFingerprint = sz & "_" & _
                      Format$(mt, "yyyymmddhhnnss") & "_" & _
                      Fnv1a32OfFileHead(path)
    Exit Function
Fail:
    FileFingerprint = "FAIL_" & Err.Number
    Err.Clear
End Function

'==============================================================================
' ToHex8 (private)
' ----------------------------------------------------------------------------
' Render a Decimal in [0, 2^32) as 8-character uppercase hex.
'==============================================================================
Private Function ToHex8(ByVal v As Variant) As String
    Static digits As String
    If LenB(digits) = 0 Then digits = "0123456789ABCDEF"

    Dim s As String, i As Long, d As Long
    s = "00000000"
    Dim n As Variant
    n = v
    For i = 8 To 1 Step -1
        d = CLng(n - (n \ 16) * 16)
        Mid$(s, i, 1) = Mid$(digits, d + 1, 1)
        n = n \ 16
    Next i
    ToHex8 = s
End Function
