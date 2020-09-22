Attribute VB_Name = "modBinDllDeclares"
Option Explicit

' Note that this is not my dll or declaration code.
' Credit should go to Peter Raddatz for this excellent work.
' http://www.planetsourcecode.com/xq/ASP/txtCodeId.14026/lngWId.1/qx/vb/scripts/ShowCode.htm

' The only bits I have used (so far) are the bit-shifting operations which cannot be done efficiently from
' Visual Basic, but can be done very easily from C.

' If you are having problems accessing it, try copying the Bindll.dll file to
' your Windows system directory.  This should cure the problem.

'--------BYTE Subs & Functions
'---Bin
'   ---Variables
Declare Sub byte2bin Lib "bindll" (ByVal BVar As Byte, BinArray As Byte)
Declare Function bin2byte Lib "bindll" (BinArray As Byte) As Byte
'   ---Arrays
Declare Sub byteA2binA Lib "bindll" (ByteArray As Byte, BinArray As Byte, ByVal Num As Integer)
Declare Sub binA2byteA Lib "bindll" (BinArray As Byte, ByteArray As Byte, ByVal Num As Integer)
'-
'---Hex
'   ---Variables
Declare Sub byte2hex Lib "bindll" (ByVal BVal As Byte, HexArray As Byte)
Declare Function hex2byte Lib "bindll" (HexArray As Byte) As Byte
'   ---Arrays
Declare Sub byteA2hexA Lib "bindll" (ByteArray As Byte, HexArray As Byte, ByVal Num As Integer)
Declare Sub hexA2byteA Lib "bindll" (HexArray As Byte, ByteArray As Byte, ByVal Num As Integer)
'-
'---Rotate
Declare Function BrotL Lib "bindll" (ByVal BVar As Byte, ByVal bits As Byte) As Byte
Declare Function BrotR Lib "bindll" (ByVal BVar As Byte, ByVal bits As Byte) As Byte
'-
'---Shift
Declare Function BshL Lib "bindll" (ByVal BVar As Byte, ByVal bits As Byte) As Byte
Declare Function BshR Lib "bindll" (ByVal BVar As Byte, ByVal bits As Byte) As Byte
'-
'---Toggle
Declare Function Btoggle Lib "bindll" (ByVal BVar As Byte, ByVal bits As Byte) As Byte
'---------------------------------End of Bytes

'--------INT Subs & Functions
'---Bin
'   ---Variables
Declare Sub int2bin Lib "bindll" (ByVal IVar As Integer, BinArray As Byte)
Declare Function bin2int Lib "bindll" (BinArray As Byte) As Integer
'   ---Arrays
Declare Sub intA2binA Lib "bindll" (IntArray As Integer, BinArray As Byte, ByVal Num As Integer)
Declare Sub binA2intA Lib "bindll" (BinArray As Byte, IntArray As Integer, ByVal Num As Integer)
'-
'---Hex
'   ---Variables
Declare Sub int2hex Lib "bindll" (ByVal IVal As Integer, HexArray As Byte)
Declare Function hex2int Lib "bindll" (HexArray As Byte) As Integer
'   ---Arrays
Declare Sub intA2hexA Lib "bindll" (IntArray As Integer, HexArray As Byte, ByVal Num As Integer)
Declare Sub hexA2intA Lib "bindll" (HexArray As Byte, IntArray As Integer, ByVal Num As Integer)
'-
'---Rotate
Declare Function IrotL Lib "bindll" (ByVal IVar As Integer, ByVal bits As Byte) As Integer
Declare Function IrotR Lib "bindll" (ByVal IVar As Integer, ByVal bits As Byte) As Integer
'-
'---Shift
Declare Function IshL Lib "bindll" (ByVal IVar As Integer, ByVal bits As Byte) As Integer
Declare Function IshR Lib "bindll" (ByVal IVar As Integer, ByVal bits As Byte) As Integer
'-
'---Toggle
Declare Function Itoggle Lib "bindll" (ByVal IVar As Integer, ByVal bits As Byte) As Integer
'---------------------------------End of Ints

'--------LONG Subs & Functions
'---Bin
'   ---Variables
Declare Sub lng2bin Lib "bindll" (ByVal LVar As Long, BinArray As Byte)
Declare Function bin2lng Lib "bindll" (BinArray As Byte) As Long
'   ---Arrays
Declare Sub lngA2binA Lib "bindll" (LngArray As Long, BinArray As Byte, ByVal Num As Integer)
Declare Sub binA2lngA Lib "bindll" (BinArray As Byte, LngArray As Long, ByVal Num As Integer)
'-
'---Hex
'   ---Variables
Declare Sub lng2hex Lib "bindll" (ByVal LVal As Long, HexArray As Byte)
Declare Function hex2lng Lib "bindll" (HexArray As Byte) As Long
'   ---Arrays
Declare Sub lngA2hexA Lib "bindll" (LngArray As Long, HexArray As Byte, ByVal Num As Integer)
Declare Sub hexA2lngA Lib "bindll" (HexArray As Byte, LngArray As Long, ByVal Num As Integer)
'-
'---Rotate
Declare Function LrotL Lib "bindll" (ByVal LVar As Long, ByVal bits As Byte) As Long
Declare Function LrotR Lib "bindll" (ByVal LVar As Long, ByVal bits As Byte) As Long
'-
'---Shift
Declare Function LshL Lib "bindll" (ByVal LVar As Long, ByVal bits As Byte) As Long
Declare Function LshR Lib "bindll" (ByVal LVar As Long, ByVal bits As Byte) As Long
'-
'---Toggle
Declare Function Ltoggle Lib "bindll" (ByVal LVar As Long, ByVal bits As Byte) As Long
'---------------------------------End of Longs
'---------------------------------End of bindll declares
'phew!

'------------------------General declare
Declare Function GetTickCount Lib "kernel32" () As Long

