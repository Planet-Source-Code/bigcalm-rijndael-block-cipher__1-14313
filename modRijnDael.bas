Attribute VB_Name = "modRijnDael"
Option Explicit

' This module contains all the constants and enums required
' (Until I find a better home for them)...
Public Const MAXKEYSIZE = 64
Public Const MAXNR = 14
Public Const MAXKB = (256 / 8)
Public Const MAXIVSIZE = 4

Public Enum RijnDaelEncDirections
    Encrypt = 0
    Decrypt = 1
End Enum
Public Enum RijnDaelCipherModes
    ECB = 1
    CBC = 2
    CFB1 = 3
End Enum

