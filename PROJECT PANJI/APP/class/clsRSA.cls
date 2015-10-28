VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0
'################################################################################'
'#  RSA Class .01                                                               #
'#  Copyright (C) 2005  Jim Banasiak     <itsjimbo@yahoo.com>                   #
'#                                                                              #
'#  This program is free software; you can redistribute it and/or               #
'#  modify it under the terms of the GNU General Public License                 #
'#  as published by the Free Software Foundation; either version 2              #
'#  of the License, or (at your option) any later version.                      #
'#                                                                              #
'#  This program is distributed in the hope that it will be useful,             #
'#  but WITHOUT ANY WARRANTY; without even the implied warranty of              #
'#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the               #
'#  GNU General Public License for more details.                                #
'#                                                                              #
'#  You should have received a copy of the GNU General Public License           #
'#  along with this program; if not, write to the Free Software                 #
'#  Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA. #
'#                                                                              #
'################################################################################
'/****************************************************************************/'
'Translation of API function calls to rsa_dll.dll (export list) by Hariyanto Lim
'This is an addition to the original code posted by Hariyanto Lim Submitted on: 11/9/2001 3:25:08 AM at planetsoucecode.com
'You can download and compile the source of rsa_dll.dll by getting it from planetsourcecode.com

Private Declare Function RSA_generate_keys Lib "rsa_dll.dll" (ByVal lpPrivKey As Long, ByVal lpPubKey As Long, ByRef lpPrivKeyLength As Long, ByRef lpPubKeyLength As Long) As Long 'returns 0 on success
'/*  RSA_encrypt:: result is an allocated buffer, the result_size is for comparison
    'whether the buffer size is enough or not
    'the return value is the REAL result size required */
Private Declare Function RSA_encrypt Lib "rsa_dll.dll" (ByVal lpSource As Long, ByVal source_size As Long, ByVal lpResult As Long, ByRef result_size As Long, ByVal pubkey_content As Long, ByVal pubkey_size As Long) As Long
'/*  RSA_decrypt:: result is an allocated buffer, the result_size is for comparison
    'whether the buffer size is enough or not
    'the return value is the REAL result size required
    'result_size' must be at least equal to 'source_size' to prevent overlap */
Private Declare Function RSA_decrypt Lib "rsa_dll.dll" (ByVal lpSource As Long, ByVal source_size As Long, ByVal lpResult As Long, ByVal result_size As Long, ByVal prikey_content As Long, ByVal prikey_size As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function RSA_sign Lib "rsa_dll.dll" (ByVal lpSource As Long, ByVal source_size As Long, ByVal lpResult As Long, ByVal result_size As Long, ByVal prikey_content As Long, ByVal prikey_size As Long) As Long

'/*  result is an allocated buffer, the result_size is for comparison
'    whether the buffer size is enough or not
'    the return value is the REAL result size required
'    result_size MUST be at least equal with source_size, to prevent error and memory overlap */
Private Declare Function RSA_verify Lib "rsa_dll.dll" (ByVal lpSource As Long, ByVal source_size As Long, ByVal lpResult As Long, ByVal result_size As Long, ByVal pubkey_content As Long, ByVal pubkey_size As Long) As Long
Public m_PublicKey As clsCert
Public m_PrivateKey As clsCert

Private Sub Class_Initialize()
    Set m_PublicKey = New clsCert
    Set m_PrivateKey = New clsCert
End Sub

Private Sub Class_Terminate()
'
End Sub
Property Get PublicKey() As clsCert
Set PublicKey = m_PublicKey
End Property
Property Get PrivateKey() As clsCert
Set PrivateKey = m_PrivateKey
End Property
Public Sub GenerateKeys()
On Error GoTo generr
Dim ret As Long
'pad the strings
Dim privkey As String
Dim pubkey As String
Dim Prikeylength As Long
Dim Pubkeylength As Long
privkey = String$(5000, vbNullChar)
pubkey = String$(5000, vbNullChar)

ret = RSA_generate_keys(StrPtr(privkey), StrPtr(pubkey), Prikeylength, Pubkeylength)
If (ret = 0) Then
        m_PrivateKey.Key = StrConv(privkey, vbUnicode)
        m_PublicKey.Key = StrConv(pubkey, vbUnicode)
        m_PrivateKey.Length = Prikeylength
        m_PublicKey.Length = Pubkeylength
        
End If
Exit Sub
generr:
If err.Number = 53 Then
  'no dll found
  MsgBox "1) Rename rsa_dll.dl_ to rsa_dll.dll " & vbCrLf & _
         "2)Copy rsa_dll.dll to the " & App.Path & " directory or the windows\sytem32 directory"
Else
MsgBox err.Description
End If


End Sub
Public Function Sign(Private_key As clsCert, ByVal content As String) As String
Dim ret As Long
Dim result As String
Dim result_sz As Long
Dim i As Long
Dim Length As Long
Dim Conv As String

Dim ss() As Byte
If Private_key.Key <> "" Then
            i = 1
            Length = Len(content)
            Conv = StrConv(content, vbFromUnicode)
            i = ((Length * 2) \ 256) + 1
            result = String$(256 * i, vbNullChar)
            result = StrConv(result, vbFromUnicode)
            result_sz = 256 * i
            ret = RSA_sign(StrPtr(Conv), Length, StrPtr(result), result_sz, StrPtr(StrConv(Private_key.Key, vbFromUnicode)), Private_key.Length)
        If (ret > result_sz) Then
                result_sz = ret
                result = String$(result_sz, vbNullChar)
                result = StrConv(result, vbFromUnicode)
                ret = RSA_sign(StrPtr(Conv), Length, StrPtr(result), result_sz, StrPtr(StrConv(Private_key.Key, vbFromUnicode)), Private_key.Length)
        End If
        
        If (ret > 0) Then
            result = StrConv(result, vbUnicode)
            result = Mid(result, 1, ret)
            Sign = HexFromString(result)
          
         Else
          err.Raise vbObjectError + 1111, , "Encryption Failed!"
        End If
 
End If
End Function
Public Function Verify(Public_key As clsCert, ByVal content As String) As String
Dim ret As Long
Dim result As String
Dim result_sz As Long
Dim i As Long
Dim ss() As Byte
If Public_key.Key <> "" Then
            i = 1
               Do While (ret = 0)
                        ss = BytesFromHex(content)
                        result = String$(UBound(ss) * i, vbNullChar)
                        result = StrConv(result, vbFromUnicode)
                        result_sz = UBound(ss) * i
                        ret = RSA_verify(VarPtr(ss(LBound(ss))), UBound(ss), StrPtr(result), result_sz, StrPtr(StrConv(Public_key.Key, vbFromUnicode)), Public_key.Length)
                        i = i + 1
                Loop
                Erase ss()
                Verify = Mid(StrConv(result, vbUnicode), 1, ret)
End If
    
End Function

Public Function Decrypt(Private_key As clsCert, ByVal content As String) As String
'decrypt
On Error Resume Next
Dim Conv As String
Dim Length As Long
Dim ret As Long
Dim result As String
Dim result_sz As Long
Dim srcsize As Long
Dim i As Integer
Dim ss() As Byte
    i = 1
     'we should not pass the vbstring using vbFromUnicode because vb might trim it off with weird chars
   Do While (ret = 0)
            ss = BytesFromHex(content)
            result = String$(UBound(ss) + 1 * i, vbNullChar)
            result = StrConv(result, vbFromUnicode)
            Debug.Print ss(UBound(ss))
            result_sz = UBound(ss) + 1 * i
            ret = RSA_decrypt(VarPtr(ss(LBound(ss))), UBound(ss), StrPtr(result), result_sz, StrPtr(StrConv(Private_key.Key, vbFromUnicode)), Private_key.Length)
            i = i + 1
    Loop
  Erase ss()
  Decrypt = Mid(StrConv(result, vbUnicode), 1, ret)
End Function
'return a byte array with the encrypted info
Public Function Encrypt(Public_key As clsCert, ByVal content As String) As String
Dim Conv As String
Dim Length As Long
Dim ret As Long
Dim result As String
Dim result_sz As Long
Dim public_key_size As Long
Dim i As Long
If Public_key.Key <> "" Then
            Length = Len(content)
            Conv = StrConv(content, vbFromUnicode)
            i = ((Length * 2) \ 256) + 1
            result = String$(256 * i, vbNullChar)
            result = StrConv(result, vbFromUnicode)
            result_sz = 256 * i
            ret = RSA_encrypt(StrPtr(Conv), Length, StrPtr(result), result_sz, StrPtr(StrConv(Public_key.Key, vbFromUnicode)), Public_key.Length)
        If (ret > result_sz) Then
                result_sz = ret
                result = String$(result_sz, vbNullChar)
                result = StrConv(result, vbFromUnicode)
                ret = RSA_encrypt(StrPtr(Conv), Length, StrPtr(result), result_sz, StrPtr(StrConv(Public_key.Key, vbFromUnicode)), Public_key.Length)
        End If
        If (ret > 0) Then
            result = StrConv(result, vbUnicode)
            result = Mid(result, 1, ret)
            Encrypt = HexFromString(result)
          Debug.Print "okay"
          
         Else
          err.Raise vbObjectError + 1111, , "Encryption Failed!"
        End If
End If

End Function

'-----------------------------------------------------------------
' Misc functions
'-----------------------------------------------------------------
Public Function BytesFromHex(ByVal sInputHex As String) As Variant
'returns array of bytes from the hex
' E.g. sHex="E32F3C" will return array {&HE3, &H2F, &H3C}
    Dim i As Long
    Dim M As Long
    Dim aBytes() As Byte
    If Len(sInputHex) Mod 2 <> 0 Then
        sInputHex = "0" & sInputHex
    End If
    
    M = Len(sInputHex) \ 2
    If M <= 0 Then
        BytesFromHex = aBytes
        Exit Function
    End If
    
    ReDim aBytes(1 To M)
    
    For i = 1 To M
        aBytes(i) = Val("&H" & Mid$(sInputHex, i * 2 - 1, 2))
    Next
    
    BytesFromHex = aBytes

End Function

Public Function HexFromString(str As String) As String
'takes a string and converts it to the hex code
    Dim byt As Byte
    Dim i As Long
    Dim n As Long
    Dim iIndex As Long
    Dim sHex As String
    
    n = Len(str)
    sHex = String(n * 2, " ")
    iIndex = 0
    For i = 1 To n
        byt = CByte(Asc(Mid$(str, i, 1)) And &HFF)
        Mid$(sHex, iIndex + 1, 2) = HexFromByte(byt)
        iIndex = iIndex + 2
    Next
    HexFromString = sHex
    
End Function

Public Function HexFromByte(ByVal x) As String
'helper to hexfrom bytes
    x = x And &HFF
    If x < 16 Then
        HexFromByte = "0" & Hex(x)
    Else
        HexFromByte = Hex(x)
    End If
End Function


Public Function HexFromBytes(aBytes() As Byte) As String
'takes the bytes and represents as hex
    Dim i As Long
    Dim iIndex As Long
    Dim nLen As Long
    On Error GoTo ArrayIsEmpty

    nLen = UBound(aBytes) - LBound(aBytes) + 1

    HexFromBytes = String(nLen * 2, " ")
    iIndex = 0
    For i = LBound(aBytes) To UBound(aBytes)
        Mid$(HexFromBytes, iIndex + 1, 2) = HexFromByte(aBytes(i))
        iIndex = iIndex + 2
    Next
    
ArrayIsEmpty:
    
End Function
