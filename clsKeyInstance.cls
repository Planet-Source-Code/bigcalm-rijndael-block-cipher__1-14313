VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rijndaelKeyInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' Where I've thought it appropriate, I've left the original
' source code commented out here.
' This is for my reference so that I can trace bugs more
' easily.
' See Readme.rtf file (included) for credits.
' This implementation by Jonathan Daniel, January 2001

Public mDirection As RijnDaelEncDirections
Public mkeyLength As Long
Public mkeyMaterial As String
Public mNr As Long
Private mRk(4 * (MAXNR + 1)) As Long
Private mEk(4 * (MAXNR + 1)) As Long
Public Mode As RijnDaelCipherModes
Private IV(MAXIVSIZE) As Long


' Next two functions are intended for use with two-way encrypted conversations
' As I don't know too much about encryption, these are probably
' totally wrong - if they are, tell me, and I'll change them.
' VB won't let me compile this until I put this class into
' an OCX.  Therefore I've left them commented until next release.
'Type RijnDaelDistributableKey
'    Algorithm As Long ' Shrug. 1 = DES, 2 = RijnDael ?
'    CipherMode As RijnDaelCipherModes
'    KeySize As Long
'    rk(4 * (MAXNR + 1)) As Long
'    Nr As Long
'    IV(MAXIVSIZE) As Long
'End Type
'Public Function CreateDistributableKey() As RijnDaelDistributableKey
'Dim tmpRDDK As RijnDaelDistributableKey
'Dim i As Long
'    tmpRDDK.Algorithm = 2
'    tmpRDDK.CipherMode = Mode
'    tmpRDDK.KeySize = mkeyLength
'    tmpRDDK.Nr = mNr
'    For i = 0 To UBound(mek)
'        tmpRDDK.rk(i) = mek(i)
'    Next
'    CreateDistributableKey = tmpRDDK
'    For i = 0 To UBound(IV)
'        tmpRDDK.IV(i) = IV(i)
'    Next
'End Function
'
'Public Function MakeEncryptKeyFromDistributedType(pRDDK As RijnDaelDistributableKey) As Long
'Dim i As Long
'    mDirection = Encrypt
'    If pRDDK.KeySize = 128 Or pRDDK.KeySize = 192 Or pRDDK.KeySize = 256 Then
'        mkeyLength = KeyLength
'    Else
'        MakeEncryptKeyFromDistributedType = 0
'        Exit Function
'    End If
'    mNr = pRDDK.Nr
'    For i = 0 To UBound(pRDDK.rk)
'        mRk(i) = pRDDK.rk(i)
'        mEk(i) = pRDDK.rk(i)
'    Next
'    CipherInit pRDDK.CipherMode, pRDDK.IV(0), pRDDK.IV(1), pRDDK.IV(2), pRDDK.IV(3)
'End Function

Public Function ConvertPassWordStringToMakeKeyAcceptableFormat(PassWord As String, KeyLength As Long) As String
Dim i As Long, j As Long
Dim tmpStr As String
    If Not (KeyLength = 128 Or KeyLength = 192 Or KeyLength = 256) Then
        Err.Raise 1, "rijndaelKeyInstance.ConvertPassWordStringToMakeKeyAcceptableFormat", "KeyLength must be 128, 192 or 256"
    End If
    tmpStr = ""
    If Len(PassWord) < (KeyLength / 8) Then
        j = Len(PassWord)
    Else
        j = KeyLength / 8
    End If
    For i = 1 To j
        tmpStr = tmpStr & PadHexStr(Hex(Asc(Mid(PassWord, i, 1)) And &HFF&), 2)
    Next
    ' Now need to pad the string with nulls if necessary.
    For i = 1 To ((KeyLength / 8) - j)
        tmpStr = tmpStr & PadHexStr(Hex(0), 2)
    Next
    ConvertPassWordStringToMakeKeyAcceptableFormat = tmpStr
End Function

Public Function CipherInit(pMode As RijnDaelCipherModes, Optional pIV0 As Long = 0, Optional pIV1 As Long = 0, _
        Optional pIV2 As Long = 0, Optional pIV3 As Long = 0) As Long
Dim i As Long
    Mode = pMode
    IV(0) = pIV0
    IV(1) = pIV1
    IV(2) = pIV2
    IV(3) = pIV3
End Function

Public Property Get rk(ByVal vIndex As Long) As Long
    rk = mRk(vIndex)
End Property

Public Function makeKey(KeyLength As Long, Direction As RijnDaelEncDirections, KeyMaterial As String, Optional pMode As RijnDaelCipherModes = -1, Optional pIV0 As Long = 0, Optional pIV1 As Long = 0, _
        Optional pIV2 As Long = 0, Optional pIV3 As Long = 0) As Long
Dim i As Long
Dim keyMat As String
Dim cipherKey(MAXKB) As Long

    mDirection = Direction
    If KeyLength = 128 Or KeyLength = 192 Or KeyLength = 256 Then
        mkeyLength = KeyLength
    Else
        makeKey = 0
        Exit Function
    End If
    mkeyMaterial = Mid(KeyMaterial, 1, KeyLength / 4)
'    /* initialize key schedule: */
    keyMat = mkeyMaterial
    For i = 0 To (mkeyLength / 32) - 1
        cipherKey(i) = HexStrToLong(Mid(keyMat, 1 + i * 8, 8))
    Next
    If Direction = Encrypt Then
        mNr = RijnDaelKeySetupEnc(mRk, cipherKey, KeyLength)
    Else
        mNr = rijndaelKeySetupDec(mRk, cipherKey, KeyLength)
    End If
    RijnDaelKeySetupEnc mEk, cipherKey, KeyLength
    If pMode <> -1 Then
        CipherInit pMode, pIV0, pIV1, pIV2, pIV3
    End If
    makeKey = True
    Exit Function
End Function

' Original C code for makeKey()
'int makeKey(keyInstance *key, BYTE direction, int keyLen, char *keyMaterial) {
'    int i;
'    char *keyMat;
'    u8 cipherKey[MAXKB];
'
'    if (key == NULL) {
'        return BAD_KEY_INSTANCE;
'    }
'
'    if ((direction == DIR_ENCRYPT) || (direction == DIR_DECRYPT)) {
'        key->direction = direction;
'    } else {
'        return BAD_KEY_DIR;
'    }
'
'    if ((keyLen == 128) || (keyLen == 192) || (keyLen == 256)) {
'        key->keyLen = keyLen;
'    } else {
'        return BAD_KEY_MAT;
'    }
'
'    if (keyMaterial != NULL) {
'        strncpy(key->keyMaterial, keyMaterial, keyLen/4);
'    }
'
'    /* initialize key schedule: */
'    keyMat = key->keyMaterial;
'    for (i = 0; i < key->keyLen/8; i++) {
'        int t, v;
'
'        t = *keyMat++;
'        if ((t >= '0') && (t <= '9')) v = (t - '0') << 4;
'        else if ((t >= 'a') && (t <= 'f')) v = (t - 'a' + 10) << 4;
'        else if ((t >= 'A') && (t <= 'F')) v = (t - 'A' + 10) << 4;
'        else return BAD_KEY_MAT;
'
'        t = *keyMat++;
'        if ((t >= '0') && (t <= '9')) v ^= (t - '0');
'        else if ((t >= 'a') && (t <= 'f')) v ^= (t - 'a' + 10);
'        else if ((t >= 'A') && (t <= 'F')) v ^= (t - 'A' + 10);
'        else return BAD_KEY_MAT;
'
'        cipherKey[i] = (u8)v;
'    }
'    if (direction == DIR_ENCRYPT) {
'        key->Nr = rijndaelKeySetupEnc(key->rk, cipherKey, keyLen);
'    } else {
'        key->Nr = rijndaelKeySetupDec(key->rk, cipherKey, keyLen);
'    }
'    rijndaelKeySetupEnc(key->ek, cipherKey, keyLen);
'    return TRUE;
'}

Private Sub Class_Initialize()
InitialiseRijnDaelArrays
End Sub

Public Function cipherUpdateRounds(ByRef InputArr() As Long, InputLen As Long, _
    ByRef OutArr() As Long, Rounds As Long) As Boolean
Dim i As Long, j As Long
Dim Var As Variant

    j = 0
    For Each Var In InputArr
        OutArr(j) = Var
        j = j + 1
    Next
    If mDirection = Decrypt Then
    End If
    Select Case mDirection
        Case RijnDaelEncDirections.Encrypt
            rijndaelEncryptRound mRk, mNr, OutArr, Rounds
        Case RijnDaelEncDirections.Decrypt
            rijndaelDecryptRound mRk, mNr, OutArr, Rounds
    End Select
    cipherUpdateRounds = True
End Function

' Original C Source for cipherUpdateRounds()
'int cipherUpdateRounds(cipherInstance *cipher, keyInstance *key,
'        BYTE *input, int inputLen, BYTE *outBuffer, int rounds) {
'    u8 block[16];
'
'    if (cipher == NULL || key == NULL) {
'        return BAD_CIPHER_STATE;
'    }
'
'    memcpy(block, input, 16);
'
'    switch (key->direction) {
'    Case DIR_ENCRYPT:
'        rijndaelEncryptRound(key->rk, key->Nr, block, rounds);
'        break;
'
'    Case DIR_DECRYPT:
'        rijndaelDecryptRound(key->rk, key->Nr, block, rounds);
'        break;
'
'default:
'        return BAD_KEY_DIR;
'    }
'
'    memcpy(outBuffer, block, 16);
'
'    return TRUE;
'}

Public Function blockEncrypt(InputBuffer() As Long, InputLen As Long, _
                OutBuffer() As Long) As Long
Dim i As Long, j As Long, k As Long, t As Long, numBlocks As Long
Dim Block(4) As Long
Dim oBlock(4) As Long
Dim tmpIV(MAXIVSIZE) As Long

    If mDirection = Decrypt Then
        Err.Raise 1, "rijnDaelKeyInstance.blockEncrypt", "You cannot try to encrypt when you only have the decrypt key"
    End If
    If InputLen <= 0 Then
        Err.Raise 2, "rijnDaelKeyInstance.blockEncrypt", "Cannot encrypt a block of zero size"
    End If
    numBlocks = InputLen / 128
    Select Case Mode
        Case RijnDaelCipherModes.ECB
            j = 0
            For i = numBlocks To 1 Step -1
                Block(0) = InputBuffer(0 + j)
                Block(1) = InputBuffer(1 + j)
                Block(2) = InputBuffer(2 + j)
                Block(3) = InputBuffer(3 + j)
                rijndaelEncrypt mRk, mNr, Block, oBlock
                OutBuffer(0 + j) = oBlock(0)
                OutBuffer(1 + j) = oBlock(1)
                OutBuffer(2 + j) = oBlock(2)
                OutBuffer(3 + j) = oBlock(3)
                j = j + 4
            Next
        Case RijnDaelCipherModes.CBC
            j = 0
            tmpIV(0) = IV(0)
            tmpIV(1) = IV(1)
            tmpIV(2) = IV(2)
            tmpIV(3) = IV(3)
            For i = numBlocks To 1 Step -1
                Block(0) = InputBuffer(0 + j) Xor tmpIV(0)
                Block(1) = InputBuffer(1 + j) Xor tmpIV(1)
                Block(2) = InputBuffer(2 + j) Xor tmpIV(2)
                Block(3) = InputBuffer(3 + j) Xor tmpIV(4)
                rijndaelEncrypt mRk, mNr, Block, oBlock
                OutBuffer(0 + j) = oBlock(0)
                OutBuffer(1 + j) = oBlock(1)
                OutBuffer(2 + j) = oBlock(2)
                OutBuffer(3 + j) = oBlock(3)
                tmpIV(0) = oBlock(0)
                tmpIV(1) = oBlock(1)
                tmpIV(2) = oBlock(2)
                tmpIV(3) = oBlock(3)
                j = j + 4
            Next
        Case RijnDaelCipherModes.CFB1
            Err.Raise 3, "rijndaelKeyInstance.BlockEncrypt", "Encryption mode CFB1 not yet supported, sorry!"
            ' This is because I cannot be arsed to piss about with individual bytes
            ' which this mode requires.
            ' Maybe later if someone asks VERY nicely.
            '    Case MODE_CFB1:
'           iv = cipher->IV;
'           for (i = numBlocks; i > 0; i--) {
'            memcpy(outBuffer, input, 16);
'            for (k = 0; k < 128; k++) {
'                rijndaelEncrypt(key->ek, key->Nr, iv, block);
'                outBuffer[k >> 3] ^= (block[0] & 0x80U) >> (k & 7);
'                for (t = 0; t < 15; t++) {
'                    iv[t] = (iv[t] << 1) | (iv[t + 1] >> 7);
'                }
'                iv[15] = (iv[15] << 1) | ((outBuffer[k >> 3] >> (7 - (k & 7))) & 1);
'            }
'            outBuffer += 16;
'            input += 16;
'        }
'        break;
    End Select
    blockEncrypt = 128 * numBlocks
End Function

'Original C-Source for blockEncrypt
'int blockEncrypt(cipherInstance *cipher, keyInstance *key,
'        BYTE *input, int inputLen, BYTE *outBuffer) {
'    int i, k, t, numBlocks;
'    u8 block[16], *iv;
'
'    if (cipher == NULL ||
'        key == NULL ||
'        key->direction == DIR_DECRYPT) {
'        return BAD_CIPHER_STATE;
'    }
'    if (input == NULL || inputLen <= 0) {
'        return 0; /* nothing to do */
'    }
'
'    numBlocks = inputLen/128;
'
'    switch (cipher->mode) {
'    Case MODE_ECB:
'        for (i = numBlocks; i > 0; i--) {
'            rijndaelEncrypt(key->rk, key->Nr, input, outBuffer);
'            input += 16;
'            outBuffer += 16;
'        }
'        break;
'
'    Case MODE_CBC:
'        iv = cipher->IV;
'        for (i = numBlocks; i > 0; i--) {
'            ((u32*)block)[0] = ((u32*)input)[0] ^ ((u32*)iv)[0];
'            ((u32*)block)[1] = ((u32*)input)[1] ^ ((u32*)iv)[1];
'            ((u32*)block)[2] = ((u32*)input)[2] ^ ((u32*)iv)[2];
'            ((u32*)block)[3] = ((u32*)input)[3] ^ ((u32*)iv)[3];
'            rijndaelEncrypt(key->rk, key->Nr, block, outBuffer);
'            iv = outBuffer;
'            input += 16;
'            outBuffer += 16;
'        }
'        break;
'
'    Case MODE_CFB1:
'        iv = cipher->IV;
'        for (i = numBlocks; i > 0; i--) {
'            memcpy(outBuffer, input, 16);
'            for (k = 0; k < 128; k++) {
'                rijndaelEncrypt(key->ek, key->Nr, iv, block);
'                outBuffer[k >> 3] ^= (block[0] & 0x80U) >> (k & 7);
'                for (t = 0; t < 15; t++) {
'                    iv[t] = (iv[t] << 1) | (iv[t + 1] >> 7);
'                }
'                iv[15] = (iv[15] << 1) | ((outBuffer[k >> 3] >> (7 - (k & 7))) & 1);
'            }
'            outBuffer += 16;
'            input += 16;
'        }
'        break;
'
'default:
'        return BAD_CIPHER_STATE;
'    }
'
'    return 128*numBlocks;
'}

Public Function BlockDecrypt(InputBuffer() As Long, InputLen As Long, _
        OutputBuffer() As Long)
Dim i As Long, j As Long, k As Long, t As Long, numBlocks As Long
Dim Block(4) As Long
Dim oBlock(4) As Long
Dim tmpIV(4) As Long

    If mDirection = Encrypt Then
        Err.Raise 1, "rijnDaelKeyInstance.blockDecrypt", "You cannot try to decrypt when you only have the encrypt key"
    End If
    If InputLen <= 0 Then
        Err.Raise 2, "rijnDaelKeyInstance.blockDecrypt", "Cannot encrypt a block of zero size"
    End If
    numBlocks = InputLen / 128
    Select Case Mode
        Case RijnDaelCipherModes.ECB
            j = 0
            For i = numBlocks To 1 Step -1
                Block(0) = InputBuffer(0 + j)
                Block(1) = InputBuffer(1 + j)
                Block(2) = InputBuffer(2 + j)
                Block(3) = InputBuffer(3 + j)
                rijndaelDecrypt mRk, mNr, Block, oBlock
                OutputBuffer(0 + j) = oBlock(0)
                OutputBuffer(1 + j) = oBlock(1)
                OutputBuffer(2 + j) = oBlock(2)
                OutputBuffer(3 + j) = oBlock(3)
                j = j + 4
            Next
        Case RijnDaelCipherModes.CBC
            tmpIV(0) = IV(0)
            tmpIV(1) = IV(1)
            tmpIV(2) = IV(2)
            tmpIV(3) = IV(3)
            For i = numBlocks To 1 Step -1
                Block(0) = InputBuffer(0 + j)
                Block(1) = InputBuffer(1 + j)
                Block(2) = InputBuffer(2 + j)
                Block(3) = InputBuffer(3 + j)
                rijndaelDecrypt mRk, mNr, Block, oBlock
                oBlock(0) = oBlock(0) Xor tmpIV(0)
                oBlock(1) = oBlock(1) Xor tmpIV(1)
                oBlock(2) = oBlock(2) Xor tmpIV(2)
                oBlock(3) = oBlock(3) Xor IV(3)
                tmpIV(0) = Block(0)
                tmpIV(1) = Block(1)
                tmpIV(2) = Block(2)
                tmpIV(3) = Block(3)
                OutputBuffer(0 + j) = oBlock(0)
                OutputBuffer(1 + j) = oBlock(1)
                OutputBuffer(2 + j) = oBlock(2)
                OutputBuffer(3 + j) = oBlock(3)
                j = j + 4
            Next
        Case RijnDaelCipherModes.CFB1
        Err.Raise 3, "rijndaelKeyInstance.BlockDecrypt", "Encryption mode CFB1 not yet supported, sorry!"
        ' i.e. Sod off
'        iv = cipher->IV;
'        for (i = numBlocks; i > 0; i--) {
'            memcpy(outBuffer, input, 16);
'            for (k = 0; k < 128; k++) {
'                rijndaelEncrypt(key->ek, key->Nr, iv, block);
'                for (t = 0; t < 15; t++) {
'                    iv[t] = (iv[t] << 1) | (iv[t + 1] >> 7);
'                }
'                iv[15] = (iv[15] << 1) | ((input[k >> 3] >> (7 - (k & 7))) & 1);
'                outBuffer[k >> 3] ^= (block[0] & 0x80U) >> (k & 7);
'            }
'            outBuffer += 16;
'            input += 16;
'        }
'        break;
        
    End Select
    BlockDecrypt = 128 * numBlocks
End Function

' Original C-Source for blockDecrypt()
'int blockDecrypt(cipherInstance *cipher, keyInstance *key,
'        BYTE *input, int inputLen, BYTE *outBuffer) {
'    int i, k, t, numBlocks;
'    u8 block[16], *iv;
'
'    if (cipher == NULL ||
'        key == NULL ||
'        cipher->mode != MODE_CFB1 && key->direction == DIR_ENCRYPT) {
'        return BAD_CIPHER_STATE;
'    }
'    if (input == NULL || inputLen <= 0) {
'        return 0; /* nothing to do */
'    }
'
'    numBlocks = inputLen/128;
'
'    switch (cipher->mode) {
'    Case MODE_ECB:
'        for (i = numBlocks; i > 0; i--) {
'            rijndaelDecrypt(key->rk, key->Nr, input, outBuffer);
'            input += 16;
'            outBuffer += 16;
'        }
'        break;
'
'    Case MODE_CBC:
'        iv = cipher->IV;
'        for (i = numBlocks; i > 0; i--) {
'            rijndaelDecrypt(key->rk, key->Nr, input, block);
'            ((u32*)block)[0] ^= ((u32*)iv)[0];
'            ((u32*)block)[1] ^= ((u32*)iv)[1];
'            ((u32*)block)[2] ^= ((u32*)iv)[2];
'            ((u32*)block)[3] ^= ((u32*)iv)[3];
'            memcpy(cipher->IV, input, 16);
'            memcpy(outBuffer, block, 16);
'            input += 16;
'            outBuffer += 16;
'        }
'        break;
'
'    Case MODE_CFB1:
'        iv = cipher->IV;
'        for (i = numBlocks; i > 0; i--) {
'            memcpy(outBuffer, input, 16);
'            for (k = 0; k < 128; k++) {
'                rijndaelEncrypt(key->ek, key->Nr, iv, block);
'                for (t = 0; t < 15; t++) {
'                    iv[t] = (iv[t] << 1) | (iv[t + 1] >> 7);
'                }
'                iv[15] = (iv[15] << 1) | ((input[k >> 3] >> (7 - (k & 7))) & 1);
'                outBuffer[k >> 3] ^= (block[0] & 0x80U) >> (k & 7);
'            }
'            outBuffer += 16;
'            input += 16;
'        }
'        break;
'
'default:
'        return BAD_CIPHER_STATE;
'    }
'
'    return 128*numBlocks;
'}


' These two commented functions I will implement in the next
' release.  For now, use your own padding.
'/**
' * Encrypt data partitioned in octets, using RFC 2040-like padding.
' *
' * @param   input           data to be encrypted (octet sequence)
' * @param   inputOctets     input length in octets (not bits)
' * @param   outBuffer       encrypted output data
' *
' * @return  length in octets (not bits) of the encrypted output buffer.
' */
'Public Function PadEncrypt(InputBuffer() As Long, Octets As Long, OutputBuffer() As Long)
'Dim i As Long, numBlocks As Long, padLength As Long
'Dim Block(4) As Long
'
'    If mDirection = Decrypt Then
'        Err.Raise 1, "rijnDaelKeyInstance.PadEncrypt", "You cannot try to encrypt when you only have the decrypt key"
'    End If
'    If InputLen <= 0 Then
'        Err.Raise 2, "rijnDaelKeyInstance.PadEncrypt", "Cannot encrypt a block of zero size"
'    End If
'    numBlocks = Octets \ 16
'    Select Case Mode
'        Case RijnDaelCipherModes.ECB
'            j = 0
'            For i = numBlocks To 1 Step -1
'                Block(0) = InputBuffer(0 + j)
'                Block(1) = InputBuffer(1 + j)
'                Block(2) = InputBuffer(2 + j)
'                Block(3) = InputBuffer(3 + j)
'                rijndaelEncrypt mRk, mNr, Block, oBlock
'                OutBuffer(0 + j) = oBlock(0)
'                OutBuffer(1 + j) = oBlock(1)
'                OutBuffer(2 + j) = oBlock(2)
'                OutBuffer(3 + j) = oBlock(3)
'                j = j + 4
'            Next
'        Case RijnDaelCipherModes.CBC
'            j = 0
'            tmpIV(0) = IV(0)
'            tmpIV(1) = IV(1)
'            tmpIV(2) = IV(2)
'            tmpIV(3) = IV(3)
'            For i = numBlocks To 1 Step -1
'                Block(0) = InputBuffer(0 + j) Xor tmpIV(0)
'                Block(1) = InputBuffer(1 + j) Xor tmpIV(1)
'                Block(2) = InputBuffer(2 + j) Xor tmpIV(2)
'                Block(3) = InputBuffer(3 + j) Xor tmpIV(4)
'                rijndaelEncrypt mRk, mNr, Block, oBlock
'                OutBuffer(0 + j) = oBlock(0)
'                OutBuffer(1 + j) = oBlock(1)
'                OutBuffer(2 + j) = oBlock(2)
'                OutBuffer(3 + j) = oBlock(3)
'                tmpIV(0) = oBlock(0)
'                tmpIV(1) = oBlock(1)
'                tmpIV(2) = oBlock(2)
'                tmpIV(3) = oBlock(3)
'                j = j + 4
'            Next
'        Case RijnDaelCipherModes.CFB1
'            Err.Raise 3, "rijndaelKeyInstance.BlockEncrypt", "Encryption mode CFB1 not yet supported, sorry!"
'    End Select
'End Function
'    Case MODE_ECB:
'        for (i = numBlocks; i > 0; i--) {
'            rijndaelEncrypt(key->rk, key->Nr, input, outBuffer);
'            input += 16;
'            outBuffer += 16;
'        }
'        padLen = 16 - (inputOctets - 16*numBlocks);
'        assert(padLen > 0 && padLen <= 16);
'        memcpy(block, input, 16 - padLen);
'        memset(block + 16 - padLen, padLen, padLen);
'        rijndaelEncrypt(key->rk, key->Nr, block, outBuffer);
'        break;
'
'    Case MODE_CBC:
'        iv = cipher->IV;
'        for (i = numBlocks; i > 0; i--) {
'            ((u32*)block)[0] = ((u32*)input)[0] ^ ((u32*)iv)[0];
'            ((u32*)block)[1] = ((u32*)input)[1] ^ ((u32*)iv)[1];
'            ((u32*)block)[2] = ((u32*)input)[2] ^ ((u32*)iv)[2];
'            ((u32*)block)[3] = ((u32*)input)[3] ^ ((u32*)iv)[3];
'            rijndaelEncrypt(key->rk, key->Nr, block, outBuffer);
'            iv = outBuffer;
'            input += 16;
'            outBuffer += 16;
'        }
'        padLen = 16 - (inputOctets - 16*numBlocks);
'        assert(padLen > 0 && padLen <= 16);
'        for (i = 0; i < 16 - padLen; i++) {
'            block[i] = input[i] ^ iv[i];
'        }
'        for (i = 16 - padLen; i < 16; i++) {
'            block[i] = (BYTE)padLen ^ iv[i];
'        }
'        rijndaelEncrypt(key->rk, key->Nr, block, outBuffer);
'        break;
'
'default:
'        return BAD_CIPHER_STATE;
'    }
'
'    return 16*(numBlocks + 1);
'}
'
'int padDecrypt(cipherInstance *cipher, keyInstance *key,
'        BYTE *input, int inputOctets, BYTE *outBuffer) {
'    int i, numBlocks, padLen;
'    u8 block[16];
'
'    if (cipher == NULL ||
'        key == NULL ||
'        key->direction == DIR_ENCRYPT) {
'        return BAD_CIPHER_STATE;
'    }
'    if (input == NULL || inputOctets <= 0) {
'        return 0; /* nothing to do */
'    }
'    if (inputOctets % 16 != 0) {
'        return BAD_DATA;
'    }
'
'    numBlocks = inputOctets/16;
'
'    switch (cipher->mode) {
'    Case MODE_ECB:
'        /* all blocks but last */
'        for (i = numBlocks - 1; i > 0; i--) {
'            rijndaelDecrypt(key->rk, key->Nr, input, outBuffer);
'            input += 16;
'            outBuffer += 16;
'        }
'        /* last block */
'        rijndaelDecrypt(key->rk, key->Nr, input, block);
'        padLen = block[15];
'        if (padLen >= 16) {
'            return BAD_DATA;
'        }
'        for (i = 16 - padLen; i < 16; i++) {
'            if (block[i] != padLen) {
'                return BAD_DATA;
'            }
'        }
'        memcpy(outBuffer, block, 16 - padLen);
'        break;
'
'    Case MODE_CBC:
'        /* all blocks but last */
'        for (i = numBlocks - 1; i > 0; i--) {
'            rijndaelDecrypt(key->rk, key->Nr, input, block);
'            ((u32*)block)[0] ^= ((u32*)cipher->IV)[0];
'            ((u32*)block)[1] ^= ((u32*)cipher->IV)[1];
'            ((u32*)block)[2] ^= ((u32*)cipher->IV)[2];
'            ((u32*)block)[3] ^= ((u32*)cipher->IV)[3];
'            memcpy(cipher->IV, input, 16);
'            memcpy(outBuffer, block, 16);
'            input += 16;
'            outBuffer += 16;
'        }
'        /* last block */
'        rijndaelDecrypt(key->rk, key->Nr, input, block);
'        ((u32*)block)[0] ^= ((u32*)cipher->IV)[0];
'        ((u32*)block)[1] ^= ((u32*)cipher->IV)[1];
'        ((u32*)block)[2] ^= ((u32*)cipher->IV)[2];
'        ((u32*)block)[3] ^= ((u32*)cipher->IV)[3];
'        padLen = block[15];
'        if (padLen <= 0 || padLen > 16) {
'            return BAD_DATA;
'        }
'        for (i = 16 - padLen; i < 16; i++) {
'            if (block[i] != padLen) {
'                return BAD_DATA;
'            }
'        }
'        memcpy(outBuffer, block, 16 - padLen);
'        break;
'
'default:
'        return BAD_CIPHER_STATE;
'    }
'
'    return 16*numBlocks - padLen;
'}

