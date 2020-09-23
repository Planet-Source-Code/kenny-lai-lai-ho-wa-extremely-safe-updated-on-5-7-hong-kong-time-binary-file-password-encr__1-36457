VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBinaryEncryptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event EncryptProgress(Progress As Long, ProgressMax As Long)
Public Event DecryptProgress(Progress As Long, ProgressMax As Long)

Sub SaveBinaryArray(ByVal Filename As String, WriteData() As Byte)

    Dim t As Integer
    t = FreeFile
    Open Filename For Binary Access Write As #t
        
            Put #t, , WriteData()
        
    Close #t
    
End Sub

Function ReadBinaryArray(ByVal Source As String)

    Dim bytBuf() As Byte
    Dim intN As Long
    
    Dim t As Integer
    t = FreeFile
    
    Open Source For Binary Access Read As #t
    
    Dim n As Long
    
    ReDim bytBuf(1 To LOF(t)) As Byte
    Get #t, , bytBuf()
    
    ReadBinaryArray = bytBuf()
    
    Close #t
    
End Function

Function EncryptFile(Source As String, Destination As String, Password As String) As Boolean

On Error GoTo 1
    Dim ByteIn() As Byte, ByteOut() As Byte
    
    ByteIn() = ReadBinaryArray(Source)
    
    ReDim ByteOut(LBound(ByteIn) To UBound(ByteIn)) As Byte
    
    Dim i As Long, j As Long
    
    Dim PL As Integer
    PL = Len(Password)
    
    Dim ChrBNow As Integer
    Dim PosNow As Integer
    
    Dim TempByte As Integer
    
    'Decrypt
    Dim TempDByte As Integer, ByteFinal As Integer
    
    For i = LBound(ByteIn) To UBound(ByteIn)
        
        PosNow = i Mod PL
        
        ChrBNow = AscB(Mid(Password, PosNow + 1, 1)) Xor 17
        'ChrBNow = 1
        
        TempByte = (ByteIn(i) + ChrBNow) Mod 256
        
        TempDByte = (TempByte - ChrBNow)
                
                If TempDByte < 0 Then
                
                    ByteFinal = 256 - Abs(TempDByte)
                    
                Else
                    
                    ByteFinal = TempDByte
                
                End If
                
        ByteOut(i) = TempByte
        
        If i Mod 500 = 0 Then
            DoEvents
            RaiseEvent EncryptProgress(i, UBound(ByteIn))
        End If
        
    Next i

SaveBinaryArray Destination, ByteOut

EncryptFile = True
Exit Function
1
EncryptFile = False

End Function

Function DecryptFile(Source As String, Destination As String, Password As String) As Boolean

On Error GoTo 1
    Dim ByteIn() As Byte, ByteOut() As Byte
    
    ByteIn() = ReadBinaryArray(Source)
    
    ReDim ByteOut(LBound(ByteIn) To UBound(ByteIn)) As Byte
    
    Dim i As Long, j As Long
    
    Dim PL As Integer
    PL = Len(Password)
    
    Dim ChrBNow As Integer
    Dim PosNow As Integer
    
    Dim TempDByte As Integer, ByteFinal As Integer
    
    For i = LBound(ByteIn) To UBound(ByteIn)
        
        PosNow = i Mod PL
        
        ChrBNow = AscB(Mid(Password, PosNow + 1, 1)) Xor 17
        
        TempDByte = (ByteIn(i) - ChrBNow)
                
                If TempDByte < 0 Then
                
                    ByteFinal = 256 - Abs(TempDByte)
                    
                Else
                    
                    ByteFinal = TempDByte
                
                End If
        
        ByteOut(i) = ByteFinal
        
        If i Mod 500 = 0 Then
        DoEvents
        RaiseEvent DecryptProgress(i, UBound(ByteIn))
        End If
      
    Next i

SaveBinaryArray Destination, ByteOut

DecryptFile = True
Exit Function
1
DecryptFile = False

End Function
