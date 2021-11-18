Option Explicit

Const kCountries3166 = "AD;AE;AF;AG;AI;AL;AM;AN;AO;AQ;AR;AS;AT;AU;AW;AX;AZ;BA;BB;BD;BE;BF;BG;BH;BI;BJ;BL;BM;BN;BO;BR;BS;BT;BV;BW;BY;BZ;CA;CC;CD;CF;CG;CH;CI;CK;CL;CM;CN;CO;CR;CU;CV;CX;CY;CZ;DE;DJ;DK;DM;DO;DZ;EC;EE;EG;EH;ER;ES;ET;FI;FJ;FK;FM;FO;FR;GA;GB;GD;GE;GF;GG;GH;GI;GL;GM;GN;GP;GQ;GR;GS;GT;GU;GW;GY;HK;HM;HN;HR;HT;HU;ID;IE;IL;IM;IN;IO;IQ;IR;IS;IT;JE;JM;JO;JP;KE;KG;KH;KI;KM;KN;KP;KR;KW;KY;KZ;LA;LB;LC;LI;LK;LR;LS;LT;LU;LV;LY;MA;MC;MD;ME;MF;MG;MH;MK;ML;MM;MN;MO;MP;MQ;MR;MS;MT;MU;MV;MW;MX;MY;MZ;NA;NC;NE;NF;NG;NI;NL;NO;NP;NR;NU;NZ;OM;PA;PE;PF;PG;PH;PK;PL;PM;PN;PR;PS;PT;PW;PY;QA;RE;RO;RS;RU;RW;SA;SB;SC;SD;SE;SG;SH;SI;SJ;SK;SL;SM;SN;SO;SR;ST;SV;SY;SZ;TC;TD;TF;TG;TH;TJ;TK;TL;TM;TN;TO;TR;TT;TV;TW;TZ;UA;UG;UM;US;UY;UZ;VA;VC;VE;VG;VI;VN;VU;WF;WS;XS;YE;YT;ZA;ZM;ZW"

' Given the first eleven characters of an ISIN, this calculates the twelfth character, the checksum.
' © Julian D. A. Wiseman 2006 to 2009; parts, including some de-Excel-isation, by and © Patrick Honorez of www.idevlop.com.
' Believed correct. If it doesn’t always work then tough—it is free.
' Latest version available via http://www.jdawiseman.com/papers/trivia/isin.html


Public Function IsIsin(ByVal strIsin As Variant, Optional strCountries As String = kCountries3166) As Boolean
' Added by Patrick Honorez of www.idevlop.com
' Returns True if string looks like a valid ISIN, False otherwise
' Parameters: strIsin     :   ISIN to check, as a string (Null accepted)
'             strCountries:   optional list of countries. If not provided, default list will be used.
'                             if provided with empty string, this check will be bypassed
' Note: some checks are redundant, but are there for speed.
    Const kIsinLike = "[A-Z][A-Z]?????????[0-9]"
    Dim strCheck As String
    If IsNull(strIsin) Then Exit Function             ' Null values
    If Len(strIsin) <> 12 Then Exit Function          ' Will return False
    If Not strIsin Like kIsinLike Then Exit Function  ' Will return False
    If Len(strCountries) > 0 Then                     ' Test country code ?
        If InStr(1, strCountries, Left(strIsin, 2)) = 0 Then Exit Function
    End If  ' Len(strCountries) > 0
    strCheck = LastDigitISIN(Left(strIsin, 11))       ' Check digit
    If Not strCheck Like "[0-9]" Then Exit Function   ' LastDigitIsin returned an error
    If Right(strIsin, 1) = strCheck Then IsIsin = True
End Function  ' IsIsin


Public Function LastDigitISIN(ElevenChars As String) As String
    Dim i As Integer, CheckSumDigits As String, TotalScore As Integer, Char As String
    If Len(ElevenChars) <> 11 Then
        LastDigitISIN = "L"  ' Length error
        Exit Function
    End If  ' Len(ElevenChars) <> 11
    CheckSumDigits = ""
    For i = 1 To 11
        Char = UCase(Mid(ElevenChars, i, 1))
        If Char >= "0" And Char <= "9" Then
            CheckSumDigits = CheckSumDigits & Char
        ElseIf Char >= "A" And Char <= "Z" Then
            CheckSumDigits = CheckSumDigits & (10 + Asc(Char) - Asc("A"))
        Else
            LastDigitISIN = "C"  ' Character error
            Exit Function
        End If
    Next i
    TotalScore = 0
    For i = 1 To Len(CheckSumDigits)
        If (i + Len(CheckSumDigits)) Mod 2 Then
            TotalScore = TotalScore + Val(Mid(CheckSumDigits, i, 1))
        Else
            TotalScore = TotalScore + Choose(1 + Val(Mid(CheckSumDigits, i, 1)), 0, 2, 4, 6, 8, 1, 3, 5, 7, 9)
        End If  ' 0 = (i + Len(CheckSumDigits)) Mod 2
    Next i
    LastDigitISIN = Format((130 - TotalScore) Mod 10, "0")
End Function  ' LastDigitISIN(ElevenChars As String) As String


Public Function LastDigitSEDOL(SixChars As String) As String  ' SixChars front padded with zeroes
    Dim i As Integer, Char As String, Multiplier As Integer, TotalScore As Integer
    If Len(SixChars) > 6 Then
        LastDigitSEDOL = "L"  ' Length error
        Exit Function
    End If  ' Len(SixChars) > 6
    For i = 1 To Len(SixChars)
        Multiplier = Choose(i + 6 - Len(SixChars), 1, 3, 1, 7, 3, 9)
        Char = UCase(Mid(SixChars, i, 1))
        If Char >= "0" And Char <= "9" Then
            TotalScore = TotalScore + Char * Multiplier
        ElseIf Char >= "A" And Char <= "Z" Then
            TotalScore = TotalScore + (10 + Asc(Char) - Asc("A")) * Multiplier
        Else
            LastDigitSEDOL = "C"  ' Character error
            Exit Function
        End If
    Next i
    LastDigitSEDOL = (870 - TotalScore) Mod 10
End Function  ' LastDigitSEDOL(SixChars As String) As String


Public Function ISINfromSEDOL6(CountryCode As String, SixChars As String) As String
    ISINfromSEDOL6 = CountryCode & "00" _
        & String("0", 6 - Len(SixChars)) & SixChars & LastDigitSEDOL(SixChars)
    ISINfromSEDOL6 = ISINfromSEDOL6 & LastDigitISIN(ISINfromSEDOL6)
End Function  ' ISINfromSEDOL6(CountryCode3166 As String, SixChars As String) As String



