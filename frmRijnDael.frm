VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRijnDael 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The RijnDael Block Cipher"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmRijnDael.frx":0000
   ScaleHeight     =   347
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   432
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "..."
      Height          =   255
      Left            =   6120
      TabIndex        =   15
      Top             =   3960
      Width           =   255
   End
   Begin VB.TextBox txtDecrypted 
      Height          =   285
      Left            =   3600
      TabIndex        =   14
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CommandButton Command7 
      Caption         =   "..."
      Height          =   255
      Left            =   6120
      TabIndex        =   12
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txtEncryptedFile 
      Height          =   285
      Left            =   3600
      TabIndex        =   11
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      Caption         =   "..."
      Height          =   255
      Left            =   6120
      TabIndex        =   9
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox txtRawFile 
      Height          =   285
      Left            =   3600
      TabIndex        =   8
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Decrypt File"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3600
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "password"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Encrypt File"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generate Known Answer Tests (KATS)"
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate FIPS Test Vectors"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cdb1 
      Left            =   5400
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   975
      Left            =   480
      TabIndex        =   16
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Decrypted File"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Encrypted File"
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "File To Be Encrypted"
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " File Encryption Demonstration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "frmRijnDael"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Most of the code here concerns production of the "known answer tests"
' I have checked the output against the original
' C source code, so I know these tests are correct.
' That isn't to say there aren't bugs elsewhere...

' Please read included ReadMe.rtf file for credits, explanations
' and links.
'   -   Jonathan Daniel, bigcalm@hotmail.com

#Const TRACE_KAT_MCT = True

Private Sub Command1_Click()
    makeFIPSTestVectors App.Path & "\fipstest2.txt"
End Sub

Public Sub makeFIPSTestVectors(FileName As String)
Dim i As Long, KeyLength As Long, r As Long, w As Long
Dim KeyInst As New rijndaelKeyInstance
Dim KeyMateria As String
Dim PT() As Long
Dim CT() As Long
Dim FormatStr As String
Dim FileNumber As Long
#If TRACE_KAT_MCT Then
    Debug.Print "Making FIPS test vectors"
#End If

    FileNumber = FreeFile
    Open FileName For Output As #FileNumber
    
    FPrintF FileNumber, "\n" & "================================\n\n" & "FILENAME:  '%s'\n\n" & "FIPS Test Vectors\n", FileName
    '    /* 128-bit key: 00010103...0e0f: */
    KeyLength = 128
    ReDim PT(KeyLength / 32)
    ReDim CT(KeyLength / 32)
    KeyMateria = ""
    For i = 0 To (KeyLength / 8) - 1
        KeyMateria = KeyMateria & PadHexStr(Hex(i))
    Next
    FPrintF FileNumber, "\n================================\n\n"
    FPrintF FileNumber, "KEYSIZE = 128\n\n"
    FPrintF FileNumber, "KEY=%s\n\n", KeyMateria
    '    /* plaintext is always 00112233...eeff: */
    PT(0) = &H112233
    PT(1) = &H44556677
    PT(2) = &H8899AABB
    PT(3) = &HCCDDEEFF
    '    /* encryption: */
    KeyInst.makeKey KeyLength, Encrypt, KeyMateria
    KeyInst.CipherInit ECB
    FPrintF FileNumber, "Round Subkey Values (Encryptions)\n\n"
    For r = 0 To KeyInst.mNr
        FPrintF FileNumber, "RK%s=", r
        For i = 0 To 3
            w = KeyInst.rk((4 * r) + i)
            FPrintF FileNumber, "%s", PadHexStr(Hex(w), 8)
        Next
        FPrintF FileNumber, "\n"
    Next
    FPrintF FileNumber, "\nIntermediate Ciphertext Values (Encryptions)\n\n"
    FPrintF FileNumber, "PT="
    For i = 0 To 3
        FPrintF FileNumber, "%s", PadHexStr(Hex(PT(i)), 8)
    Next
    FPrintF FileNumber, "\n"
    For r = 1 To KeyInst.mNr
        KeyInst.cipherUpdateRounds PT, 4, CT, r
        FPrintF FileNumber, "CT%s=", r
        For i = 0 To 3
            FPrintF FileNumber, "%s", PadHexStr(Hex(CT(i)), 8)
        Next
        FPrintF FileNumber, "\n"
    Next
    '    /* decryption: */
    KeyInst.makeKey KeyLength, Decrypt, KeyMateria
    KeyInst.CipherInit ECB
    FPrintF FileNumber, "\nRound Subkey Values (Decryption)\n\n"
    For r = 0 To KeyInst.mNr
        FPrintF FileNumber, "RK%s=", r
        For i = 0 To 3
            w = KeyInst.rk((4 * r) + i)
            FPrintF FileNumber, "%s", PadHexStr(Hex(w), 8)
        Next
        FPrintF FileNumber, "\n"
    Next
    FPrintF FileNumber, "\nIntermediate Ciphertext Values (Decryption)\n\n"
    FPrintF FileNumber, "CT="
    For i = 0 To 3
        FPrintF FileNumber, "%s", PadHexStr(Hex(CT(i)), 8)
    Next
    FPrintF FileNumber, "\n"
    For r = 1 To KeyInst.mNr
        KeyInst.cipherUpdateRounds CT, 4, PT, r
        FPrintF FileNumber, "PT%s=", r
        For i = 0 To 3
            FPrintF FileNumber, "%s", PadHexStr(Hex(PT(i)), 8)
        Next
        FPrintF FileNumber, "\n"
    Next

    
    '    /* 192-bit key: 00010103...1617: */
    KeyLength = 192
    ReDim PT(KeyLength / 32)
    ReDim CT(KeyLength / 32)
    KeyMateria = ""
    For i = 0 To (KeyLength / 8) - 1
        KeyMateria = KeyMateria & PadHexStr(Hex(i))
    Next
    FPrintF FileNumber, "\n" & "================================\n\n" & "FILENAME:  '%s'\n\n" & "FIPS Test Vectors\n", FileName
    FPrintF FileNumber, "KEYSIZE = %s\n\n", KeyLength
    FPrintF FileNumber, "KEY=%s\n\n", KeyMateria
    '    /* plaintext is always 00112233...eeff: */
    PT(0) = &H112233
    PT(1) = &H44556677
    PT(2) = &H8899AABB
    PT(3) = &HCCDDEEFF
    '    /* encryption: */
    KeyInst.makeKey KeyLength, Encrypt, KeyMateria
    KeyInst.CipherInit ECB
    FPrintF FileNumber, "Round Subkey Values (Encryptions)\n\n"
    For r = 0 To KeyInst.mNr
        FPrintF FileNumber, "RK%s=", r
        For i = 0 To 3
            w = KeyInst.rk((4 * r) + i)
            FPrintF FileNumber, "%s", PadHexStr(Hex(w), 8)
        Next
        FPrintF FileNumber, "\n"
    Next
    FPrintF FileNumber, "\nIntermediate Ciphertext Values (Encryptions)\n\n"
    FPrintF FileNumber, "PT="
    For i = 0 To 3
        FPrintF FileNumber, "%s", PadHexStr(Hex(PT(i)), 8)
    Next
    FPrintF FileNumber, "\n"
    For r = 1 To KeyInst.mNr
        KeyInst.cipherUpdateRounds PT, 4, CT, r
        FPrintF FileNumber, "CT%s=", r
        For i = 0 To 3
            FPrintF FileNumber, "%s", PadHexStr(Hex(CT(i)), 8)
        Next
        FPrintF FileNumber, "\n"
    Next
    '    /* decryption: */
    KeyInst.makeKey KeyLength, Decrypt, KeyMateria
    KeyInst.CipherInit ECB
    FPrintF FileNumber, "\nRound Subkey Values (Decryption)\n\n"
    For r = 0 To KeyInst.mNr
        FPrintF FileNumber, "RK%s=", r
        For i = 0 To 3
            w = KeyInst.rk((4 * r) + i)
            FPrintF FileNumber, "%s", PadHexStr(Hex(w), 8)
        Next
        FPrintF FileNumber, "\n"
    Next
    FPrintF FileNumber, "\nIntermediate Ciphertext Values (Decryption)\n\n"
    FPrintF FileNumber, "CT="
    For i = 0 To 3
        FPrintF FileNumber, "%s", PadHexStr(Hex(CT(i)), 8)
    Next
    FPrintF FileNumber, "\n"
    For r = 1 To KeyInst.mNr
        KeyInst.cipherUpdateRounds CT, 4, PT, r
        FPrintF FileNumber, "PT%s=", r
        For i = 0 To 3
            FPrintF FileNumber, "%s", PadHexStr(Hex(PT(i)), 8)
        Next
        FPrintF FileNumber, "\n"
    Next
    
    '    /* 256-bit key: 00010103...1e1f: */
    KeyLength = 256
    ReDim PT(KeyLength / 32)
    ReDim CT(KeyLength / 32)
    KeyMateria = ""
    For i = 0 To (KeyLength / 8) - 1
        KeyMateria = KeyMateria & PadHexStr(Hex(i))
    Next
    FPrintF FileNumber, "\n================================\n\n"
    FPrintF FileNumber, "KEYSIZE = %s\n\n", KeyLength
    FPrintF FileNumber, "KEY=%s\n\n", KeyMateria
    '    /* plaintext is always 00112233...eeff: */
    PT(0) = &H112233
    PT(1) = &H44556677
    PT(2) = &H8899AABB
    PT(3) = &HCCDDEEFF
    '    /* encryption: */
    KeyInst.makeKey KeyLength, Encrypt, KeyMateria
    KeyInst.CipherInit ECB
    FPrintF FileNumber, "Round Subkey Values (Encryptions)\n\n"
    For r = 0 To KeyInst.mNr
        FPrintF FileNumber, "RK%s=", r
        For i = 0 To 3
            w = KeyInst.rk((4 * r) + i)
            FPrintF FileNumber, "%s", PadHexStr(Hex(w), 8)
        Next
        FPrintF FileNumber, "\n"
    Next
    FPrintF FileNumber, "\nIntermediate Ciphertext Values (Encryptions)\n\n"
    FPrintF FileNumber, "PT="
    For i = 0 To 3
        FPrintF FileNumber, "%s", PadHexStr(Hex(PT(i)), 8)
    Next
    FPrintF FileNumber, "\n"
    For r = 1 To KeyInst.mNr
        KeyInst.cipherUpdateRounds PT, 4, CT, r
        FPrintF FileNumber, "CT%s=", r
        For i = 0 To 3
            FPrintF FileNumber, "%s", PadHexStr(Hex(CT(i)), 8)
        Next
        FPrintF FileNumber, "\n"
    Next
    '    /* decryption: */
    KeyInst.makeKey KeyLength, Decrypt, KeyMateria
    KeyInst.CipherInit ECB
    FPrintF FileNumber, "\nRound Subkey Values (Decryption)\n\n"
    For r = 0 To KeyInst.mNr
        FPrintF FileNumber, "RK%s=", r
        For i = 0 To 3
            w = KeyInst.rk((4 * r) + i)
            FPrintF FileNumber, "%s", PadHexStr(Hex(w), 8)
        Next
        FPrintF FileNumber, "\n"
    Next
    FPrintF FileNumber, "\nIntermediate Ciphertext Values (Decryption)\n\n"
    FPrintF FileNumber, "CT="
    For i = 0 To 3
        FPrintF FileNumber, "%s", PadHexStr(Hex(CT(i)), 8)
    Next
    FPrintF FileNumber, "\n"
    For r = 1 To KeyInst.mNr
        KeyInst.cipherUpdateRounds CT, 4, PT, r
        FPrintF FileNumber, "PT%s=", r
        For i = 0 To 3
            FPrintF FileNumber, "%s", PadHexStr(Hex(PT(i)), 8)
        Next
        FPrintF FileNumber, "\n"
    Next
    ' end of 256
    
    Close #FileNumber
#If TRACE_KAT_MCT Then
    Debug.Print "Done"
#End If
End Sub

' Original C-Source for MakeFIPSTestVectors()
'static void makeFIPSTestVectors(const char *fipsFile) {
'    int i, keyLength, r;
'    keyInstance keyInst;
'    cipherInstance cipherInst;
'    BYTE keyMaterial[320];
'    u8 pt[16], ct[16];
'    char format[64];
'    FILE *fp;
'
'#ifdef TRACE_KAT_MCT
'    printf("Generating FIPS test vectors...");
'#endif /* ?TRACE_KAT_MCT */
'
'    fp = fopen(fipsFile, "w");
'    fprintf(fp,
'        "\n"
'        "================================\n\n"
'        "FILENAME:  \"%s\"\n\n"
'        "FIPS Test Vectors\n",
'        fipsFile);
'
'    /* 128-bit key: 00010103...0e0f: */
'    keyLength = 128;
'    memset(keyMaterial, 0, sizeof (keyMaterial));
'    for (i = 0; i < keyLength/8; i++) {
'        sprintf(&keyMaterial[2*i], "%02X", i);
'    }
'
'    fprintf(fp, "\n================================\n\n");
'    fprintf(fp, "KEYSIZE=128\n\n");
'    fprintf(fp, "KEY=%s\n\n", keyMaterial);
'
'    /* plaintext is always 00112233...eeff: */
'    for (i = 0; i < 16; i++) {
'        pt[i] = (i << 4) | i;
'    }
'
'    /* encryption: */
'    makeKey(&keyInst, DIR_ENCRYPT, keyLength, keyMaterial);
'    cipherInit(&cipherInst, MODE_ECB, NULL);
'    fprintf(fp, "Round Subkey Values (Encryption)\n\n");
'    for (r = 0; r <= keyInst.Nr; r++) {
'        fprintf(fp, "RK%d=", r);
'        for (i = 0; i < 4; i++) {
'            u32 w = keyInst.rk[4*r + i];
'            fprintf(fp, "%02X%02X%02X%02X", w >> 24, (w >> 16) & 0xff, (w >> 8) & 0xff, w & 0xff);
'        }
'        fprintf(fp, "\n");
'    }
'    fprintf(fp, "\nIntermediate Ciphertext Values (Encryption)\n\n");
'    blockPrint(fp, pt, "PT");
'    for (i = 1; i < keyInst.Nr; i++) {
'        cipherUpdateRounds(&cipherInst, &keyInst, pt, 16, ct, i);
'        sprintf(format, "CT%d", i);
'        blockPrint(fp, ct, format);
'    }
'    cipherUpdateRounds(&cipherInst, &keyInst, pt, 16, ct, keyInst.Nr);
'    blockPrint(fp, ct, "CT");
'
'    /* decryption: */
'    makeKey(&keyInst, DIR_DECRYPT, keyLength, keyMaterial);
'    cipherInit(&cipherInst, MODE_ECB, NULL);
'    fprintf(fp, "\nRound Subkey Values (Decryption)\n\n");
'    for (r = 0; r <= keyInst.Nr; r++) {
'        fprintf(fp, "RK%d=", r);
'        for (i = 0; i < 4; i++) {
'            u32 w = keyInst.rk[4*r + i];
'            fprintf(fp, "%02X%02X%02X%02X", w >> 24, (w >> 16) & 0xff, (w >> 8) & 0xff, w & 0xff);
'        }
'        fprintf(fp, "\n");
'    }
'    fprintf(fp, "\nIntermediate Ciphertext Values (Decryption)\n\n");
'    blockPrint(fp, ct, "CT");
'    for (i = 1; i < keyInst.Nr; i++) {
'        cipherUpdateRounds(&cipherInst, &keyInst, ct, 16, pt, i);
'        sprintf(format, "PT%d", i);
'        blockPrint(fp, pt, format);
'    }
'    cipherUpdateRounds(&cipherInst, &keyInst, ct, 16, pt, keyInst.Nr);
'    blockPrint(fp, pt, "PT");
'
'    /* 192-bit key: 00010103...1617: */
'    keyLength = 192;
'    memset(keyMaterial, 0, sizeof (keyMaterial));
'    for (i = 0; i < keyLength/8; i++) {
'        sprintf(&keyMaterial[2*i], "%02X", i);
'    }
'
'    fprintf(fp, "\n================================\n\n");
'    fprintf(fp, "KEYSIZE=192\n\n");
'    fprintf(fp, "KEY=%s\n\n", keyMaterial);
'
'    /* plaintext is always 00112233...eeff: */
'    for (i = 0; i < 16; i++) {
'        pt[i] = (i << 4) | i;
'    }
'
'    /* encryption: */
'    makeKey(&keyInst, DIR_ENCRYPT, keyLength, keyMaterial);
'    cipherInit(&cipherInst, MODE_ECB, NULL);
'    fprintf(fp, "\nRound Subkey Values (Encryption)\n\n");
'    for (r = 0; r <= keyInst.Nr; r++) {
'        fprintf(fp, "RK%d=", r);
'        for (i = 0; i < 4; i++) {
'            u32 w = keyInst.rk[4*r + i];
'            fprintf(fp, "%02X%02X%02X%02X", w >> 24, (w >> 16) & 0xff, (w >> 8) & 0xff, w & 0xff);
'        }
'        fprintf(fp, "\n");
'    }
'    fprintf(fp, "\nIntermediate Ciphertext Values (Encryption)\n\n");
'    blockPrint(fp, pt, "PT");
'    for (i = 1; i < keyInst.Nr; i++) {
'        cipherUpdateRounds(&cipherInst, &keyInst, pt, 16, ct, i);
'        sprintf(format, "CT%d", i);
'        blockPrint(fp, ct, format);
'    }
'    cipherUpdateRounds(&cipherInst, &keyInst, pt, 16, ct, keyInst.Nr);
'    blockPrint(fp, ct, "CT");
'
'    /* decryption: */
'    makeKey(&keyInst, DIR_DECRYPT, keyLength, keyMaterial);
'    cipherInit(&cipherInst, MODE_ECB, NULL);
'    fprintf(fp, "\nRound Subkey Values (Decryption)\n\n");
'    for (r = 0; r <= keyInst.Nr; r++) {
'        fprintf(fp, "RK%d=", r);
'        for (i = 0; i < 4; i++) {
'            u32 w = keyInst.rk[4*r + i];
'            fprintf(fp, "%02X%02X%02X%02X", w >> 24, (w >> 16) & 0xff, (w >> 8) & 0xff, w & 0xff);
'        }
'        fprintf(fp, "\n");
'    }
'    fprintf(fp, "\nIntermediate Ciphertext Values (Decryption)\n\n");
'    blockPrint(fp, ct, "CT");
'    for(i = 1; i < keyInst.Nr; i++) {
'        cipherUpdateRounds(&cipherInst, &keyInst, ct, 16, pt, i);
'        sprintf(format, "PT%d", i);
'        blockPrint(fp, pt, format);
'    }
'    cipherUpdateRounds(&cipherInst, &keyInst, ct, 16, pt, keyInst.Nr);
'    blockPrint(fp, pt, "PT");
'
'    /* 256-bit key: 00010103...1e1f: */
'    keyLength = 256;
'    memset(keyMaterial, 0, sizeof (keyMaterial));
'    for (i = 0; i < keyLength/8; i++) {
'        sprintf(&keyMaterial[2*i], "%02X", i);
'    }
'
'    fprintf(fp, "\n================================\n\n");
'    fprintf(fp, "KEYSIZE=256\n\n");
'    fprintf(fp, "KEY=%s\n\n", keyMaterial);
'
'    /* plaintext is always 00112233...eeff: */
'    for (i = 0; i < 16; i++) {
'        pt[i] = (i << 4) | i;
'    }
'
'    /* encryption: */
'    makeKey(&keyInst, DIR_ENCRYPT, keyLength, keyMaterial);
'    cipherInit(&cipherInst, MODE_ECB, NULL);
'    fprintf(fp, "\nRound Subkey Values (Encryption)\n\n");
'    for (r = 0; r <= keyInst.Nr; r++) {
'        fprintf(fp, "RK%d=", r);
'        for (i = 0; i < 4; i++) {
'            u32 w = keyInst.rk[4*r + i];
'            fprintf(fp, "%02X%02X%02X%02X", w >> 24, (w >> 16) & 0xff, (w >> 8) & 0xff, w & 0xff);
'        }
'        fprintf(fp, "\n");
'    }
'    fprintf(fp, "\nIntermediate Ciphertext Values (Encryption)\n\n");
'    blockPrint(fp, pt, "PT");
'    for(i = 1; i < keyInst.Nr; i++) {
'        cipherUpdateRounds(&cipherInst, &keyInst, pt, 16, ct, i);
'        sprintf(format, "CT%d", i);
'        blockPrint(fp, ct, format);
'    }
'    cipherUpdateRounds(&cipherInst, &keyInst, pt, 16, ct, keyInst.Nr);
'    blockPrint(fp, ct, "CT");
'
'    /* decryption: */
'    makeKey(&keyInst, DIR_DECRYPT, keyLength, keyMaterial);
'    cipherInit(&cipherInst, MODE_ECB, NULL);
'    fprintf(fp, "\nRound Subkey Values (Decryption)\n\n");
'    for (r = 0; r <= keyInst.Nr; r++) {
'        fprintf(fp, "RK%d=", r);
'        for (i = 0; i < 4; i++) {
'            u32 w = keyInst.rk[4*r + i];
'            fprintf(fp, "%02X%02X%02X%02X", w >> 24, (w >> 16) & 0xff, (w >> 8) & 0xff, w & 0xff);
'        }
'        fprintf(fp, "\n");
'    }
'    fprintf(fp, "\nIntermediate Ciphertext Values (Decryption)\n\n");
'    blockPrint(fp, ct, "CT");
'    for(i = 1; i < keyInst.Nr; i++) {
'        cipherUpdateRounds(&cipherInst, &keyInst, ct, 16, pt, i);
'        sprintf(format, "PT%d", i);
'        blockPrint(fp, pt, format);
'    }
'    cipherUpdateRounds(&cipherInst, &keyInst, ct, 16, pt, keyInst.Nr);
'    blockPrint(fp, pt, "PT");
'
'    fprintf(fp, "\n");
'    fclose(fp);
'#ifdef TRACE_KAT_MCT
'    printf(" done.\n");
'#endif /* ?TRACE_KAT_MCT */
'}

Private Sub Command2_Click()
    MakeKats "ecb_vk2.txt", "ecb_vt2.txt", "ecb_tbl2.txt", "ecb_iv2.txt"
End Sub

Public Sub rijndaelVKKAT(FileNumber As Long, KeyLength As Long)
Dim i As Long, j As Long, r As Long
Dim Block(4) As Long
Dim KeyMaterial As String
Dim byteVal As Byte
Dim KeyInst As rijndaelKeyInstance
#If TRACE_KAT_MCT Then
    PrintF "Executing Variable-Key KAT(Key %s): ", KeyLength
#End If
    byteVal = 8
    FPrintF FileNumber, "\n============\n\nKEYSIZE=%s\n\n", KeyLength
    FPrintF FileNumber, "PT="
    For i = 0 To 3
        Block(i) = Block(0)
        FPrintF FileNumber, "%s", PadHexStr(Hex(Block(i)), 8)
    Next
    FPrintF FileNumber, "\n"
    Set KeyInst = New rijndaelKeyInstance
    KeyMaterial = RepeatChar("0", KeyLength \ 4)
    For i = 0 To KeyLength - 1
        KeyMaterial = RepeatChar("0", i \ 4) & Hex(byteVal) & RepeatChar("0", (KeyLength \ 4) - (i \ 4) - 1)
        r = KeyInst.makeKey(KeyLength, Encrypt, KeyMaterial)
        If r <> True Then
            PrintF "makeKey error %s\n", r
            End
        End If
        FPrintF FileNumber, "\nI=%s\n", i + 1
        FPrintF FileNumber, "KEY=%s\n", KeyMaterial
        Block(0) = 0: Block(1) = 0: Block(2) = 0: Block(3) = 0
        r = KeyInst.CipherInit(ECB)
        If r <> 0 Then
            PrintF "cipherInit error %s\n", r
            End
        End If
        r = KeyInst.blockEncrypt(Block, 128, Block)
        If r <> 128 Then
            PrintF "blockEncrypt error %s\n", r
            End
        End If
        FPrintF FileNumber, "CT="
        For j = 0 To 3
            FPrintF FileNumber, "%s", PadHexStr(Hex(Block(j)), 8)
        Next
        FPrintF FileNumber, "\n"
        '        /* now check decryption: */
        KeyInst.makeKey KeyLength, Decrypt, KeyMaterial
        KeyInst.BlockDecrypt Block, 128, Block
        For j = 0 To 3
            If Block(j) <> 0 Then
                PrintF "Assert!  Encrypt/Decrypt Failed! %s\n", j
                End
            End If
        Next
        '        /* undo changes for the next iteration: */
        KeyMaterial = ""
        Select Case byteVal
            Case 8
                byteVal = 4
            Case 4
                byteVal = 2
            Case 2
                byteVal = 1
            Case 1
                byteVal = 8
        End Select
    Next
#If TRACE_KAT_MCT Then
    Debug.Print "Done"
#End If
End Sub

Private Sub rijndaelVTKAT(FileNumber As Long, KeyLength As Long)
Dim i As Long, j As Long
Dim Block(4) As Long
Dim KeyMaterial As String
Dim KeyInst As rijndaelKeyInstance
Dim tmpStr As String
#If TRACE_KAT_MCT Then
    PrintF "Executing Variable-Text KAT (Key %s): ", KeyLength
#End If
    FPrintF FileNumber, "\n===========\n\nKEYSIZE=%s\n\n", KeyLength
    KeyMaterial = RepeatChar("0", KeyLength / 4)
    Set KeyInst = New rijndaelKeyInstance
    KeyInst.makeKey KeyLength, Encrypt, KeyMaterial
    FPrintF FileNumber, "KEY=%s\n", KeyMaterial
    For i = 0 To 127
        Block(0) = 0: Block(1) = 0: Block(2) = 0: Block(3) = 0
        ' I hate clever C programmers sometimes... i.e.
        '        block[i/8] |= 1 << (7 - i%8); /* set only the i-th bit of the i-th test block */
        Block(i \ 32) = HexStrToLong(RepeatChar("0", ((i Mod 32) \ 4)) & (2 ^ (3 - (i Mod 4))) & RepeatChar("0", 7 - ((i Mod 32) \ 4)))
        ' Revenge is sweet.  (Don't ask me to explain this.  I won't know by the time I look at it again!).  It works ok?
        ' Who says only C programmers can write incomprehensible code eh?
        FPrintF FileNumber, "\nI=%s\n", i + 1
        FPrintF FileNumber, "PT="
        For j = 0 To 3
            FPrintF FileNumber, "%s", PadHexStr(Hex(Block(j)), 8)
        Next
        FPrintF FileNumber, "\n"
        KeyInst.CipherInit ECB
        KeyInst.blockEncrypt Block, 128, Block
        FPrintF FileNumber, "CT="
        For j = 0 To 3
            FPrintF FileNumber, "%s", PadHexStr(Hex(Block(j)), 8)
        Next
        FPrintF FileNumber, "\n"
    Next
#If TRACE_KAT_MCT Then
    Debug.Print "Done"
#End If
End Sub

Private Sub rijndaelTKAT(FileNumber As Long, KeyLength As Long, FileNumber2 As Long)
Dim i As Long, j As Long
Dim s As Long
Dim Block(4) As Long
Dim Block2(4) As Long
Dim KeyMaterial As String
Dim KeyInst As rijndaelKeyInstance
Dim tmpStr As String, tmpStr2 As String
Dim LineInpStr As String
Dim StrArr() As String

#If TRACE_KAT_MCT Then
    PrintF "Executing Tables KAT (key %s): ", KeyLength
#End If
    Set KeyInst = New rijndaelKeyInstance
    FPrintF FileNumber, "\n==========\n\nKEYSIZE=%s\n\n", KeyLength
    For i = 0 To 63
        FPrintF FileNumber, "\nI=%s\n", i + 1
        Line Input #FileNumber2, LineInpStr
        StrArr = Split(LineInpStr, " ")
        KeyMaterial = StrArr(0)
        KeyInst.makeKey KeyLength, Encrypt, KeyMaterial
        FPrintF FileNumber, "KEY=%s\n", KeyMaterial
        tmpStr2 = ""
        For j = 0 To 15
            tmpStr2 = tmpStr2 & StrArr(j + 1)
            If (j + 1) Mod 4 = 0 Then
                Block(j \ 4) = HexStrToLong(tmpStr2)
                tmpStr2 = ""
            End If
        Next
        FPrintF FileNumber, "PT="
        For j = 0 To 3
            FPrintF FileNumber, "%s", PadHexStr(Hex(Block(j)), 8)
        Next
        FPrintF FileNumber, "\n"
        KeyInst.CipherInit ECB
        KeyInst.blockEncrypt Block, 128, Block2
        FPrintF FileNumber, "CT="
        For j = 0 To 3
            FPrintF FileNumber, "%s", PadHexStr(Hex(Block2(j)), 8)
        Next
        FPrintF FileNumber, "\n"
    Next
    For i = 64 To 127
        FPrintF FileNumber, "\nI=%s\n", i + 1
        Line Input #FileNumber2, LineInpStr
        StrArr = Split(LineInpStr, " ")
        KeyMaterial = StrArr(0)
        KeyInst.makeKey KeyLength, Decrypt, KeyMaterial
        FPrintF FileNumber, "KEY=%s\n", KeyMaterial
        tmpStr2 = ""
        For j = 0 To 15
            tmpStr2 = tmpStr2 & StrArr(j + 1)
            If (j + 1) Mod 4 = 0 Then
                Block(j \ 4) = HexStrToLong(tmpStr2)
                tmpStr2 = ""
            End If
        Next
        KeyInst.CipherInit ECB
        KeyInst.BlockDecrypt Block, 128, Block2
        FPrintF FileNumber, "PT="
        For j = 0 To 3
            FPrintF FileNumber, "%s", PadHexStr(Hex(Block2(j)), 8)
        Next
        FPrintF FileNumber, "\n"
        FPrintF FileNumber, "CT="
        For j = 0 To 3
            FPrintF FileNumber, "%s", PadHexStr(Hex(Block(j)), 8)
        Next
        FPrintF FileNumber, "\n"
    Next
#If TRACE_KAT_MCT Then
    Debug.Print "Done"
#End If
End Sub

Private Sub MakeKats(vkFile As String, vtFile As String, tblFile As String, ivFile As String)
Dim FileNumber As Long, FileNumber2 As Long
    FileNumber = FreeFile
    Open App.Path & "\" & vkFile For Output As #FileNumber
    FPrintF FileNumber, "\n=====================\n\nFILENAME:  %s\n\nElectronic Codebook (ECB) Mode\nVariable Key Known Answer Tests\nAlgorithm Name: RijnDael", vkFile
    rijndaelVKKAT FileNumber, 128
    rijndaelVKKAT FileNumber, 192
    rijndaelVKKAT FileNumber, 256
    FPrintF FileNumber, "\n============\n"
    Close #FileNumber
    '    /* prepare Variable Text Known Answer Tests: */
    FileNumber = FreeFile
    Open App.Path & "\" & vtFile For Output As #FileNumber
    FPrintF FileNumber, "\n=================\n\nFILENAME: %s\n\nElectronic Codebook (ECB) Mode\nVariable Known Answer Tests\n\nAlgorithm Name: RijnDael\n", vtFile
    rijndaelVTKAT FileNumber, 128
    rijndaelVTKAT FileNumber, 192
    rijndaelVTKAT FileNumber, 256
    FPrintF FileNumber, "\n=============\n"
    Close #FileNumber
    '    /* prepare Tables Known Answer Tests: */
    FileNumber = FreeFile
    Open App.Path & "\" & tblFile For Output As #FileNumber
    FPrintF FileNumber, "/* Description of what tables are tested:\nThe provided implementations each use a different set of tables\n" & _
        "    - Java implementation: uses no tables\n" & _
        "    - reference C implementation: uses Logtable, Alogtable, S, Si, rcon\n" & _
        "    - fast C implementation: uses rcon and additionally\n" & _
        "        Te0, Te1, Te2, Te3, Te4, Td0, Td1, Td2, Td3, Td4.\n" & _
        "    - VB implementation: uses rcon and additionally\n" & _
        "        Te0, Te1, Te2, Te3, Te4, Td0, Td1, Td2, Td3, Td4. - as optimised C\n" & _
        "   All these tables are tested.\n" & _
        "\n" & _
        "=========================\n" & _
        "\n" & _
        "FILENAME:  %s\n" & _
        "\n" & _
        "Electronic Codebook (ECB) Mode\n" & _
        "Tables Known Answer Tests\n" & _
        "\n" & _
        "Algorithm Name: Rijndael\n\n", tblFile
    FileNumber2 = FreeFile
    Open App.Path & "\table.128.txt" For Input As #FileNumber2
    rijndaelTKAT FileNumber, 128, FileNumber2
    Close #FileNumber2
    FileNumber2 = FreeFile
    Open App.Path & "\table.192.txt" For Input As #FileNumber2
    rijndaelTKAT FileNumber, 192, FileNumber2
    Close #FileNumber2
    FileNumber2 = FreeFile
    Open App.Path & "\table.256.txt" For Input As #FileNumber2
    rijndaelTKAT FileNumber, 256, FileNumber2
    Close #FileNumber2
    FPrintF FileNumber, "\n===========\n"
    Close #FileNumber
End Sub

'#ifdef INTERMEDIATE_VALUE_KAT
'    /* prepare Intermediate Values Known Answer Tests: */
'    fp = fopen(ivFile, "w");
'    fprintf(fp,
'        "\n"
'        "=========================\n"
'        "\n"
'        "FILENAME:  \"%s\"\n"
'        "\n"
'        "Electronic Codebook (ECB) Mode\n"
'        "Intermediate Value Known Answer Tests\n"
'        "\n"
'        "Algorithm Name: Rijndael\n"
'        "Principal Submitter: %s\n",
'        ivFile, SUBMITTER);
'    fflush(fp);
'
'    rijndaelIVKAT(fp, 128);
'    rijndaelIVKAT(fp, 192);
'    rijndaelIVKAT(fp, 256);
'
'    fprintf(fp,
'        "\n"
'        "==========");
'    fclose(fp);
'#endif /* INTERMEDIATE_VALUE_KAT */
'}

Private Sub Command3_Click()
    MsgBox "None as yet"
End Sub

Private Sub DecryptFile(InputFile As String, OutputFile As String, PassWord As String)
Dim KeyInst As rijndaelKeyInstance
Dim KeyMaterial As String
Dim KeyLength As Long
Dim FileNumber As Long, FileNumber2 As Long
Dim Block(4) As Long
Dim oBlock(4) As Long
Dim CompareMaterial As String
Dim i As Long, j As Long
Dim StartTime As Long
Dim EndTime As Long
    StartTime = timeGetTime
    ' Setup key...
    KeyLength = 128
    Set KeyInst = New rijndaelKeyInstance
    KeyMaterial = KeyInst.ConvertPassWordStringToMakeKeyAcceptableFormat(PassWord, KeyLength)
    KeyInst.makeKey KeyLength, Decrypt, KeyMaterial, CBC
    ' Open files...
    FileNumber = FreeFile
    Open InputFile For Binary Access Read As #FileNumber
    
    ' Read back encryption key to check that the password
    ' was correct
    GetBlock FileNumber, Block
    KeyInst.BlockDecrypt Block, 128, oBlock
    CompareMaterial = ""
    For i = 0 To 3
        CompareMaterial = CompareMaterial & PadHexStr(Hex(oBlock(i)), 8)
    Next
    If KeyLength > 128 Then
        GetBlock FileNumber, Block
        KeyInst.BlockDecrypt Block, 128, oBlock
        If KeyLength = 192 Then
            j = 1
        Else
            j = 3
        End If
        For i = 0 To j
            CompareMaterial = CompareMaterial & PadHexStr(Hex(oBlock(i)), 8)
        Next
    End If
    If CompareMaterial <> KeyMaterial Then
        MsgBox "Invalid Password!!!"
        Close #FileNumber
        Close #FileNumber2
        Exit Sub
    End If
    On Error Resume Next
    Kill OutputFile
    On Error GoTo 0
    FileNumber2 = FreeFile
    Open OutputFile For Binary Access Write As #FileNumber2
    ' If we're ok, we just need to decrypt now
    Do While EOF(FileNumber) = False
        GetBlock FileNumber, Block
        KeyInst.BlockDecrypt Block, 128, oBlock
        PutBlock FileNumber2, oBlock
    Loop
    Close #FileNumber
    Close #FileNumber2
    EndTime = timeGetTime
    Label6.Caption = "File of size " & FileLen(InputFile) & " decrypted in " & EndTime - StartTime & " milliseconds"
End Sub

Private Sub EncryptFile(InputFile As String, OutputFile As String, PassWord As String)
Dim KeyInst As rijndaelKeyInstance
Dim KeyMaterial As String
Dim KeyLength As Long
Dim FileNumber As Long, FileNumber2 As Long
Dim Block(4) As Long
Dim oBlock(4) As Long
Dim StartTime As Long
Dim EndTime As Long
    StartTime = timeGetTime
    ' Setup key...
    KeyLength = 128
    Set KeyInst = New rijndaelKeyInstance
    KeyMaterial = KeyInst.ConvertPassWordStringToMakeKeyAcceptableFormat(PassWord, KeyLength)
    KeyInst.makeKey KeyLength, Encrypt, KeyMaterial, CBC
    
    ' Open files...
    FileNumber = FreeFile
    Open InputFile For Binary Access Read As #FileNumber
    On Error Resume Next
    Kill OutputFile
    On Error GoTo 0
    FileNumber2 = FreeFile
    Open OutputFile For Binary Access Write As #FileNumber2
    
    ' The first thing we need to do is to encrypt the key material (or some known value)
    ' This means that when we decrypt we have a check to see if
    ' decryption works OK.
    Block(0) = HexStrToLong(Mid(KeyMaterial, 1, 8))
    Block(1) = HexStrToLong(Mid(KeyMaterial, 9, 8))
    Block(2) = HexStrToLong(Mid(KeyMaterial, 17, 8))
    Block(3) = HexStrToLong(Mid(KeyMaterial, 25, 8))
    KeyInst.blockEncrypt Block, 128, oBlock
    PutBlock FileNumber2, oBlock
    If KeyLength > 128 Then
        Block(0) = 0: Block(1) = 0: Block(2) = 0: Block(3) = 0
        Block(0) = HexStrToLong(Mid(KeyMaterial, 33, 8))
        Block(1) = HexStrToLong(Mid(KeyMaterial, 41, 8))
        If KeyLength > 192 Then
            Block(2) = HexStrToLong(Mid(KeyMaterial, 49, 8))
            Block(3) = HexStrToLong(Mid(KeyMaterial, 57, 8))
        End If
        KeyInst.blockEncrypt Block, 128, oBlock
        PutBlock FileNumber2, oBlock
    End If
    ' Put "Expected number of bytes?"
    ' Now that we have output the key, I want to encrypt the file
    Do While EOF(FileNumber) = False
        GetBlock FileNumber, Block
        KeyInst.blockEncrypt Block, 128, oBlock
        PutBlock FileNumber2, oBlock
    Loop
    Close #FileNumber
    Close #FileNumber2
    EndTime = timeGetTime
    Label6.Caption = "File of size " & FileLen(InputFile) & " encrypted in " & EndTime - StartTime & " milliseconds"
End Sub

Private Sub PutBlock(FileNumber As Long, ByRef Block() As Long)
Dim i As Long
    For i = 0 To 3
        Put #FileNumber, , Block(i)
    Next
End Sub
Private Sub GetBlock(FileNumber As Long, ByRef Block() As Long)
Dim i As Long
    For i = 0 To 3
        Get #FileNumber, , Block(i)
    Next
End Sub

Private Sub Command4_Click()
Dim tmpPass As String
    tmpPass = frmPassWord.GetPassWord("Please enter password to Encrypt File", "RijnDael Block Cipher", "")
    If Len(tmpPass) = 0 Then
        Exit Sub
    End If
    EncryptFile txtRawFile.Text, txtEncryptedFile.Text, tmpPass
End Sub

Private Sub Command5_Click()
Dim tmpPass As String
    tmpPass = frmPassWord.GetPassWord("Please enter password to Decrypt File", "RijnDael Block Cipher", "")
    If Len(tmpPass) = 0 Then
        Exit Sub
    End If
    DecryptFile txtEncryptedFile.Text, txtDecrypted.Text, tmpPass
End Sub

Private Sub Command6_Click()
    On Error GoTo ErrHandler
    cdb1.CancelError = True
    cdb1.DialogTitle = "Raw File to be encrypted..."
    cdb1.Filter = "All Files (*.*)|*.*"
    cdb1.FilterIndex = 0
    cdb1.FileName = ""
    cdb1.Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist
    cdb1.MaxFileSize = 32000
    cdb1.ShowOpen
    txtRawFile.Text = cdb1.FileName
    Exit Sub
ErrHandler:
    If Err.Number = 32755 Then
       ' cancel was selected
    Else
        MsgBox Err.Description
    End If
    Exit Sub
End Sub

Private Sub Command7_Click()
    On Error GoTo ErrHandler
    cdb1.CancelError = True
    cdb1.DefaultExt = ".rji"
    cdb1.DialogTitle = "Output file for encryption..."
    cdb1.Filter = "RijnDael encrypted files (*.rji)|*.rji"
    cdb1.FilterIndex = 0
    cdb1.FileName = ""
    cdb1.Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist
    cdb1.MaxFileSize = 32000
    cdb1.ShowOpen
    txtEncryptedFile.Text = cdb1.FileName
    Exit Sub
ErrHandler:
    If Err.Number = 32755 Then
       ' cancel was selected
    Else
        MsgBox Err.Description
    End If
    Exit Sub
End Sub

Private Sub Command8_Click()
    On Error GoTo ErrHandler
    cdb1.CancelError = True
    cdb1.DialogTitle = "Output file for encryption..."
    cdb1.DefaultExt = ""
    cdb1.Filter = "All files (*.*)|*.*"
    cdb1.FilterIndex = 0
    cdb1.FileName = ""
    cdb1.Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist
    cdb1.MaxFileSize = 32000
    cdb1.ShowSave
    txtEncryptedFile.Text = cdb1.FileName
    Exit Sub
ErrHandler:
    If Err.Number = 32755 Then
       ' cancel was selected
    Else
        MsgBox Err.Description
    End If
    Exit Sub
End Sub

Private Sub Form_Load()
    txtRawFile.Text = App.Path & "\Image3.gif"
    txtEncryptedFile.Text = App.Path & "\Image3.rji"
    txtDecrypted.Text = App.Path & "\Test.gif"
End Sub

