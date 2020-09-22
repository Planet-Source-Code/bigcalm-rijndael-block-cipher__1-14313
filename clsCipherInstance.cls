VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCipherInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Mode As CipherModes
Private IV(MAXIVSIZE) As Long

Public Function CipherInit(pMode As CipherModes, pIV As String) As Long
Dim i As Long
    Mode = pMode
    If Len(pIV) > 0 Then
        For i = 0 To (MAXIVSIZE - 8) Step 8
            IV(i) = HexStrToLong(Mid(pIV, 1 + i, 8))
        Next
    Else
        Erase IV
    End If
End Function
