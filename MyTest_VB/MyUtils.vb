Imports System.Security.Cryptography
Imports System.Text
Imports System.Runtime.CompilerServices

Namespace mytest_vb
    Public Delegate Function ParamsAction(<[ParamArray]()> args As Object()) As Int32

    Public Class MyUtils

        Public maxPrime As Int32
        Public testing As Boolean = True
        Public primesList As Int32()
        Private rndNum As Byte() = New Byte(4 - 1) {}
        Private rngCSP As RNGCryptoServiceProvider = New RNGCryptoServiceProvider()
        Protected Overrides Sub Finalize()
            rngCSP.Dispose()
        End Sub

        Public Sub New()
        End Sub
        Public Sub New(inTesting As Boolean)
            testing = inTesting
            maxPrime = Primes.GetPrimes().Last()
            primesList = Primes.GetPrimes()
        End Sub

        ' ---------- Helper functions ----------- '
        Public Function GetR(x As Int32, y As Int32) As Int32
            If (x = y) Then Return x

            Dim rslt As Int32 = x - 1
            Dim width = Math.Abs(x - y)

            rngCSP.GetBytes(rndNum)

            rslt = (BitConverter.ToUInt32(rndNum, 0) Mod width)
            rslt += Math.Min(x, y)

            Return rslt + 1

        End Function

        Public Function GetF() As Double
            Dim b As Byte() = New Byte(12) {}
            Dim rngCrypto As RNGCryptoServiceProvider = New RNGCryptoServiceProvider()
            rngCrypto.GetBytes(b)

            Dim sb As StringBuilder = New StringBuilder("0.")
            Dim numbers As Int32() = b.Select(Function(i) Convert.ToInt32((i * 100 / 255) / 10)).ToArray()

            For Each number As Int32 In numbers
                sb.Append(number)
            Next

            Return Convert.ToDouble(sb.ToString())
        End Function

        Public Function myLog(msg As String, <CallerMemberName> Optional name As String = "")
            If (testing) Then
                Console.WriteLine("\nError in " + name + ": " + msg)
            End If
        End Function


        ' ----------  Method caches ------------ '
        Private IsPrimeCache As Dictionary(Of Int32, Boolean) = New Dictionary(Of Int32, Boolean)()
        Private primeFactorsCache As Dictionary(Of Int32, Int32()) = New Dictionary(Of Int32, Int32())()
        Private LcmCache As Dictionary(Of Int32(), Long) = New Dictionary(Of Int32(), Long)()
        Private HcfCache As Dictionary(Of Int32(), Int32) = New Dictionary(Of Int32(), Int32)()
        Private SignificantFiguresCache As Dictionary(Of Double, Int32) = New Dictionary(Of Double, Int32)()
        Private StandardFormCache As Dictionary(Of Double, String) = New Dictionary(Of Double, String)()

        ' ----------  Utils functions ----------- '
        Public Function myTestRi(x As Int32, y As Int32, Optional nonZero As Boolean = True) As Int32
            Dim rtn As Int32 = 0
            Try
                If x = y Then Return x
                If x < 0 And y < 0 Then
                    x = Math.Abs(x)
                    y = Math.Abs(y)
                    If x > y Then
                        Dim tmp As Int32 = y
                        y = x
                        x = tmp
                    End If
                    While nonZero And rtn = 0
                        rtn = GetR(x, y) * -1
                    End While
                ElseIf x < 0 Then
                    If y = 0 Then
                        While nonZero And rtn = 0
                            rtn = GetR(0, Math.Abs(x)) * -1
                        End While
                    Else
                        While nonZero And rtn = 0
                            rtn = GetR(x, y)
                        End While
                    End If
                ElseIf y < 0 Then
                    If x = 0 Then
                        While nonZero And rtn = 0
                            rtn = GetR(0, Math.Abs(y)) * -1
                        End While
                    Else
                        While nonZero And rtn = 0
                            rtn = GetR(y, x)
                        End While
                    End If
                Else
                    If x > y Then
                        Dim tmp As Int32 = y
                        y = x
                        x = tmp
                    End If
                    Dim nudge As Int32 = 0
                    If x = 0 Then
                        nudge += 1
                    End If
                    If y + nudge = 0 Then
                        nudge += 1
                    End If
                    While nonZero And rtn = 0
                        rtn = GetR(x + nudge, y + nudge) - nudge
                    End While
                End If
            Catch ex As Exception
                myLog("myTestRi - " + ex.Message)
                Throw
            End Try
            Return rtn
        End Function

        Private Function GetDec(z As Int32, power As Double) As Double
            Dim rtn As Double = 0.0
            Dim n As String = ""
            While (Not width(rtn) - 2.0 = z Or Math.Round(rtn * power - 0.5, z) Mod 10 = 0 Or rtn = 0.0)
                Dim i As Int32 = width(rtn)
                n = (GetF()).ToString()
                Dim j As Int32 = Math.Round(rtn * power - 0.5, z) Mod 10
                Double.TryParse(n.Substring(0, 2 + z), rtn)
            End While
            Return rtn
        End Function

        Public Function myTestRd(x As Int32, y As Int32, z As Double, Optional nonZero As Boolean = True) As Double
            Dim rtn As Double = 0.0
            Try
                If z = 0 Then
                    Return CDbl(myTestRi(x, y))
                End If
                z = Math.Abs(z)
                Dim power As Double = Math.Pow(10.0, z)
                Dim dec As Double = 0.0
                dec = GetDec(z, power)

                If x = y Then
                    If x < 0 Then
                        dec -= 1
                    End If
                    Return ((x + dec) * power) / power
                End If
                If x < 0 And y < 0 Then
                    x = Math.Abs(x)
                    y = Math.Abs(y)
                    If (x > y) Then
                        Dim tmp As Int32 = y
                        y = x
                        x = tmp
                    End If
                    If y - x = 1 Then
                        Return (x + dec) * -1
                    End If
                    While nonZero And rtn = 0.0
                        rtn = (myTestRi(x, (y - 1)) + dec) * -1
                    End While
                ElseIf (x < 0) Then
                    If y - x = 1 Then
                        rtn = dec * -1
                    ElseIf y = 0 Then
                        While (nonZero And rtn = 0)
                            rtn = (myTestRi(0, Math.Abs(x) - 1) + dec) * -1
                        End While
                    Else
                        While (nonZero And rtn = 0.0)
                            rtn = CDbl(myTestRi(x, y - 1)) + (1.0 - dec)
                        End While
                    End If
                ElseIf y < 0 Then
                    If x - y = 1 Then
                        rtn = dec * -1
                    ElseIf x = 0 Then
                        While nonZero And rtn = 0
                            rtn = (myTestRi(0, Math.Abs(y) - 1) + dec) * -1
                        End While
                    Else
                        While nonZero And rtn = 0
                            rtn = CDbl(myTestRi(y, x - 1)) + dec
                        End While
                    End If
                Else
                    If x > y Then
                        Dim tmp As Int32 = y
                        y = x
                        x = tmp
                    End If
                    y -= 1
                    Dim nudge As Int32 = 0
                    If x = 0 Then
                        nudge += 1
                    End If
                    If y + nudge = 0 Then
                        nudge += 1
                    End If
                    While nonZero And rtn = 0
                        rtn = (myTestRi(x + nudge, y + nudge) - nudge) + dec
                    End While
                End If
            Catch ex As Exception
                myLog("myTestRi - " + ex.Message)
                Throw
            End Try
            Return rtn
        End Function

        Public Function width(x As Double) As Int32
            Return CDbl(Math.Abs(x)).ToString().Length
        End Function

        Public Function myTestRfr(numWidth As Int32, denomWidth As Int32, Optional proper As Boolean = True) As Int32()
            Dim num As Int32 = 0
            Dim denom As Int32 = 0
            Try
                If (numWidth <= 0 Or denomWidth <= 0 Or numWidth > 9 Or denomWidth > 9) Then
                    Throw New Exception("Argument(s) exception: rfr(" + numWidth + ", " + denomWidth + "})")
                End If
                num = myTestRi(CInt(Math.Pow(10, numWidth - 1)), CInt(Math.Pow(10, numWidth) - 1))
                denom = myTestRi(CInt(Math.Pow(10, denomWidth - 1)), CInt(Math.Pow(10, denomWidth) - 1))
                If numWidth > denomWidth Then
                    proper = False
                End If
                While (num >= denom And proper)
                    num = myTestRi(CInt(Math.Pow(10, numWidth - 1)), CInt(Math.Pow(10, numWidth) - 1))
                    denom = myTestRi(CInt(Math.Pow(10, denomWidth - 1)), CInt(Math.Pow(10, denomWidth) - 1))
                End While
            Catch ex As Exception
                myLog("myTestRfr - " + ex.Message)
                Throw
            End Try
            Return New Int32() {num, denom}
        End Function

        Public Function myTestIsPrime(x As Int32) As Boolean
            If IsPrimeCache.ContainsKey(x) Then
                Return IsPrimeCache(x)
            End If
            Dim rtn As Boolean = True
            Try
                If x = 2 Then
                    Return True
                ElseIf (x < 2) Then
                    Return False
                ElseIf x Mod 2 = 0 And x > 2 Then
                    Return False
                ElseIf x < maxPrime Then
                    rtn = primesList.Contains(x)
                Else
                    For y = 3 To x
                        If (y Mod x = 0) Then
                            rtn = False
                            Exit For
                        End If
                    Next
                End If
                IsPrimeCache(x) = rtn
                Return rtn
            Catch ex As Exception
                myLog("myTestIsPrime - " + ex.Message)
                Throw
            End Try
            Return True
        End Function

        Public Function myTestPrimeFactors(x As Int32) As Int32()
            If primeFactorsCache.ContainsKey(x) Then
                Return primeFactorsCache(x)
            End If
            Dim orig As Int32 = x
            Dim rtn As List(Of Int32) = New List(Of Int32)()
            Try
                'if (x > 1048577)
                If (x > 1000000) Then
                    Throw New Exception("Argument exception: myTestPrimefactors(" + x + ")")
                End If
                If x < 2 Then
                    Return New Int32() {0}
                End If
                If x < 4 Then
                    Return New Int32() {x}
                End If
                Dim i As Int32 = 0
                Try
                    i = Math.Max(x / Math.Max(CInt(Math.Abs(x)).ToString().Length * 2, 1), 1) + 5
                Catch ex As Exception
                    i = x
                End Try
                Dim prime As Int32 = 0
                While i >= 0
                    prime = primesList(i)
                    If x Mod prime = 0 Then
                        rtn.Add(prime)
                        x /= prime
                        Try
                            i = Math.Max(x / Math.Max(CInt(Math.Abs(x)).ToString().Length * 2, 1), 1) + 5
                        Catch ex As Exception
                            i = x
                        End Try
                    Else
                        i -= 1
                    End If
                End While
            Catch ex As Exception
                myLog("myTestPrimeFactors - " + ex.Message)
                Throw
            End Try
            rtn.Reverse()
            primeFactorsCache(orig) = rtn.ToArray()
            Return primeFactorsCache(orig)
        End Function

        Public Function myTestLcm(arr As Int32()) As Long
            If (LcmCache.ContainsKey(arr)) Then
                Return LcmCache(arr)
            End If
            Dim rtn As Long = 0
            Try
                Dim mult As Long = 1
                arr = arr.Select(Function(a) Math.Abs(a)).ToArray()
                mult = arr.Aggregate(1, Function(a, b) a * b)
                If mult > 150000000 Then
                    Throw New Exception("\nInput array too large: " + mult + " - mytestLcm")
                End If
                rtn = mult
                Dim skip As Boolean = False
                Dim count As Int32 = arr.Length
                For i = mult To 1 Step -1
                    skip = False
                    For j = 0 To count - 1
                        If Not i Mod arr(j) = 0 Then
                            skip = True
                            Exit For
                        End If
                    Next
                    If Not skip Then
                        rtn = i
                    End If
                Next
            Catch ex As Exception
                myLog("myTestLcm - " + ex.Message)
                Throw
            End Try
            LcmCache(arr) = rtn
            Return rtn
        End Function

        Public Function myTestHcf(arr As Int32())
            If HcfCache.ContainsKey(arr) Then
                Return HcfCache(arr)
            End If
            Dim rtn As Int32 = 0
            Try
                arr = arr.Select(Function(a) Math.Abs(a)).ToArray()
                Dim bound As Int32 = arr.Min()
                Dim skip As Boolean = False
                Dim count As Int32 = arr.Length
                For i = 1 To bound
                    For j = 0 To count - 1
                        If Not arr(j) Mod i = 0 Then
                            skip = True
                            Exit For
                        End If
                    Next
                    If Not skip Then
                        rtn = i
                    End If
                    skip = False
                Next
            Catch ex As Exception
                myLog("myTestHcf - " + ex.Message)
                Throw
            End Try
            HcfCache(arr) = rtn
            Return rtn
        End Function

        Public Function myLibSquare(x As Int32, y As Int32) As Int32
            Dim ran As Int32 = 0
            Try
                If y < 0 Then
                    Throw New Exception("\nArgument(s) exception: " + x + ", " + y + " in myLibSquare")
                End If
                If x = 0 Then
                    Return 0
                End If
                ran = myTestRi(x, y)
            Catch ex As Exception
                myLog("myLibSquare - " + ex.Message)
                Throw
            End Try
            Return CInt(Math.Pow(ran, 2))
        End Function

        Public Function myLibIsSquare(x As Int32) As Boolean
            Dim rtn As Boolean = False
            Try
                If x < 4 Then
                    Return False
                End If
                rtn = myLibIsInt(Math.Sqrt(CDbl(x)))
            Catch ex As Exception
                myLog("myLibIsSquare - " + ex.Message)
                Throw
            End Try
            Return rtn
        End Function

        Public Function myLibChoose(Of T)(inSet As T())
            Return inSet(GetR(0, inSet.Length - 1))
        End Function

        Public Function myLibIsInt(x As Double)
            Return x = Math.Ceiling(x)
        End Function

        Public Function myLibNotSquare(x As Int32, y As Int32) As Int32
            Dim rtn As Int32 = 4
            Try
                While myLibIsSquare(rtn) 'Could run forever...
                    rtn = myTestRi(x, y)
                    If 1 = Math.Abs(rtn) Then
                        Exit While
                    End If
                End While
            Catch ex As Exception
                myLog("myLibNotSquare - " + ex.Message)
                Throw
            End Try
            Return rtn
        End Function

        Public Function myLibRunFuncUntilNot(notUs As Int32(), func As Func(Of Int32, Int32, Boolean, Int32), maxTries As Int32, ParamArray args As Object())
            Dim rslt As Int32 ' TODO: Pass in a failure return
            Try
                For i = 0 To maxTries
                    rslt = func.DynamicInvoke(args)
                    If Not notUs.Contains(rslt) Then
                        Return True
                    End If
                Next
                Return False
            Catch ex As Exception
                myLog("myLibRunFuncUntilNot - " + ex.Message)
                Throw
            End Try
            Return False
        End Function

        Public Function myLibSignificantFigures(x As Double, figures As Int32) As Double
            'If SignificantFiguresCache.ContainsKey(x, figures) Then
            '    Return SignificantFiguresCache(x, figures)
            'End If
            Dim rtn As Double = Math.Abs(x)
            Try
                Dim shifts As Int32 = 0
                If figures < 1 Then
                    Throw New Exception("\nArgument(s) exception: " + x + ", " + figures + " in myLibSignificantFigures")
                End If
                If x = 0 Then
                    Return 0
                End If
                If rtn < 1.0 Then
                    While rtn < 1.0
                        rtn *= 10
                        shifts -= 1
                    End While
                    rtn *= 0.1
                    shifts += 1
                Else
                    While rtn > 1.0
                        rtn *= 0.1
                        shifts += 1
                    End While
                End If
                rtn *= Math.Pow(10, figures)
                rtn = Math.Round(rtn)
                rtn *= Math.Pow(10, shifts - figures)
                If x < 0 Then
                    rtn *= -1
                End If
            Catch ex As Exception
                myLog("myLibSignificantFigures - " + ex.Message)
                Throw
            End Try
            'SignificantFiguresCache((x, figures))=rtn
            Return rtn
        End Function

        Public Function myLibToSup(str As String) As String
            Dim sbRtn As StringBuilder = New StringBuilder()
            Try
                For Each s In str
                    Select Case s
                        Case "1"
                            sbRtn.Append("¹")
                        Case "2"
                            sbRtn.Append("²")
                        Case "3"
                            sbRtn.Append("³")
                        Case "4"
                            sbRtn.Append("⁴")
                        Case "5"
                            sbRtn.Append("⁵")
                        Case "6"
                            sbRtn.Append("⁶")
                        Case "7"
                            sbRtn.Append("⁷")
                        Case "8"
                            sbRtn.Append("⁸")
                        Case "9"
                            sbRtn.Append("⁹")
                        Case "0"
                            sbRtn.Append("⁰")
                        Case "-"
                            sbRtn.Append("⁻")
                        Case "+"
                            sbRtn.Append("⁺")
                        Case "*"
                            sbRtn.Append("*")
                        Case "/"
                            sbRtn.Append("ᐟ")
                        Case "\\"
                            sbRtn.Append("ᐠ")
                        Case "a"
                            sbRtn.Append("ᵃ")
                        Case "b"
                            sbRtn.Append("ᵇ")
                        Case "c"
                            sbRtn.Append("ᶜ")
                        Case "d"
                            sbRtn.Append("ᵈ")
                        Case "e"
                            sbRtn.Append("ᵉ")
                        Case "f"
                            sbRtn.Append("ᶠ")
                        Case "g"
                            sbRtn.Append("ᵍ")
                        Case "h"
                            sbRtn.Append("ʰ")
                        Case "i"
                            sbRtn.Append("ⁱ")
                        Case "j"
                            sbRtn.Append("ʲ")
                        Case "k"
                            sbRtn.Append("ᵏ")
                        Case "l"
                            sbRtn.Append("ˡ")
                        Case "m"
                            sbRtn.Append("ᵐ")
                        Case "n"
                            sbRtn.Append("ⁿ")
                        Case "o"
                            sbRtn.Append("ᵒ")
                        Case "p"
                            sbRtn.Append("ᵖ")
                        Case "q"
                            sbRtn.Append("۹")
                        Case "r"
                            sbRtn.Append("ʳ")
                        Case "s"
                            sbRtn.Append("ˢ")
                        Case "t"
                            sbRtn.Append("ᵗ")
                        Case "u"
                            sbRtn.Append("ᵘ")
                        Case "v"
                            sbRtn.Append("ᵛ")
                        Case "w"
                            sbRtn.Append("ʷ")
                        Case "x"
                            sbRtn.Append("ˣ")
                        Case "y"
                            sbRtn.Append("ʸ")
                        Case "z"
                            sbRtn.Append("ᶻ")
                        Case "A"
                            sbRtn.Append("ᴬ")
                        Case "B"
                            sbRtn.Append("ᴮ")
                        Case "C"
                            sbRtn.Append("ᶜ")
                        Case "D"
                            sbRtn.Append("ᴰ")
                        Case "E"
                            sbRtn.Append("ᴱ")
                        Case "F"
                            sbRtn.Append("ᶠ")
                        Case "G"
                            sbRtn.Append("ᴳ")
                        Case "H"
                            sbRtn.Append("ᴴ")
                        Case "I"
                            sbRtn.Append("ᴵ")
                        Case "J"
                            sbRtn.Append("ᴶ")
                        Case "K"
                            sbRtn.Append("ᴷ")
                        Case "L"
                            sbRtn.Append("ᴸ")
                        Case "M"
                            sbRtn.Append("ᴹ")
                        Case "N"
                            sbRtn.Append("ᴺ")
                        Case "O"
                            sbRtn.Append("ᴼ")
                        Case "P"
                            sbRtn.Append("ᴾ")
                        Case "R"
                            sbRtn.Append("ᴿ")
                        Case "T"
                            sbRtn.Append("ᵀ")
                        Case "S"
                            sbRtn.Append("ˢ")
                        Case "U"
                            sbRtn.Append("ᵁ")
                        Case "V"
                            sbRtn.Append("ⱽ")
                        Case "W"
                            sbRtn.Append("ᵂ")
                        Case "X"
                            sbRtn.Append("ˣ")
                        Case "Y"
                            sbRtn.Append("ʸ")
                        Case "Z"
                            sbRtn.Append("ᶻ")
                        Case "="
                            sbRtn.Append("⁼")
                        Case "("
                            sbRtn.Append("⁽")
                        Case ")"
                            sbRtn.Append("⁾")
                        Case Else
                            sbRtn.Append(s)
                    End Select
                Next
            Catch ex As Exception
                myLog("myLibToSuper - " + ex.Message)
                Throw
            End Try
            Return sbRtn.ToString()
        End Function

        Public Function myLibToSub(str As String) As String
            Dim sbRtn As StringBuilder = New StringBuilder()
            Try
                For Each s In str
                    Select Case s
                        Case "1"
                            sbRtn.Append("₁")
                        Case "2"
                            sbRtn.Append("₂")
                        Case "3"
                            sbRtn.Append("₃")
                        Case "4"
                            sbRtn.Append("₄")
                        Case "5"
                            sbRtn.Append("₅")
                        Case "6"
                            sbRtn.Append("₆")
                        Case "7"
                            sbRtn.Append("₇")
                        Case "8"
                            sbRtn.Append("₈")
                        Case "9"
                            sbRtn.Append("₉")
                        Case "0"
                            sbRtn.Append("₀")
                        Case "+"
                            sbRtn.Append("₊")
                        Case "-"
                            sbRtn.Append("₋")
                        Case "="
                            sbRtn.Append("₌")
                        Case "("
                            sbRtn.Append("₍")
                        Case ")"
                            sbRtn.Append("₎")
                        Case "a"
                            sbRtn.Append("ₐ")
                        Case "b"
                            sbRtn.Append("♭")
                        Case "e"
                            sbRtn.Append("ₑ")
                        Case "g"
                            sbRtn.Append("₉")
                        Case "h"
                            sbRtn.Append("ₕ")
                        Case "i"
                            sbRtn.Append("ᵢ")
                        Case "j"
                            sbRtn.Append("ⱼ")
                        Case "k"
                            sbRtn.Append("ₖ")
                        Case "l"
                            sbRtn.Append("ₗ")
                        Case "m"
                            sbRtn.Append("ₘ")
                        Case "n"
                            sbRtn.Append("ₙ")
                        Case "o"
                            sbRtn.Append("ₒ")
                        Case "p"
                            sbRtn.Append("ₚ")
                        Case "r"
                            sbRtn.Append("ᵣ")
                        Case "s"
                            sbRtn.Append("ₛ")
                        Case "t"
                            sbRtn.Append("ₜ")
                        Case "u"
                            sbRtn.Append("ᵤ")
                        Case "v"
                            sbRtn.Append("ᵥ")
                        Case "x"
                            sbRtn.Append("ₓ")
                        Case "A"
                            sbRtn.Append("ₐ")
                        Case "B"
                            sbRtn.Append("₈")
                        Case "E"
                            sbRtn.Append("ₑ")
                        Case "F"
                            sbRtn.Append("բ")
                        Case "G"
                            sbRtn.Append("₉")
                        Case "H"
                            sbRtn.Append("ₕ")
                        Case "I"
                            sbRtn.Append("ᵢ")
                        Case "J"
                            sbRtn.Append("ⱼ")
                        Case "K"
                            sbRtn.Append("ₖ")
                        Case "L"
                            sbRtn.Append("ₗ")
                        Case "M"
                            sbRtn.Append("ₘ")
                        Case "N"
                            sbRtn.Append("ₙ")
                        Case "O"
                            sbRtn.Append("ₒ")
                        Case "P"
                            sbRtn.Append("ₚ")
                        Case "R"
                            sbRtn.Append("ᵣ")
                        Case "T"
                            sbRtn.Append("ₜ")
                        Case "S"
                            sbRtn.Append("ₛ")
                        Case "U"
                            sbRtn.Append("ᵤ")
                        Case "V"
                            sbRtn.Append("ᵥ")
                        Case "W"
                            sbRtn.Append("w")
                        Case "X"
                            sbRtn.Append("ₓ")
                        Case "Y"
                            sbRtn.Append("ᵧ")
                        Case "Z"
                            sbRtn.Append("Z")
                        Case Else
                            sbRtn.Append(s)
                    End Select
                Next
            Catch ex As Exception
                myLog("myLibToSub - " + ex.Message)
                Throw
            End Try
            Return sbRtn.ToString()
        End Function

        Public Function myLibToStandardForm(x As Double) As String
            'If StandardFormCache.ContainsKey(x) Then
            '    Return StandardFormCache(x)
            'End If
            Dim sign As Int32 = 0
            Dim expSigned As Boolean = False
            Dim shifts As Int32 = 0
            Dim floatLen As Int32 = 0
            Dim orig As Double = x
            Try
                If 0 = x Then
                    Return "0"
                End If
                If 1 = x Then
                    Return "1"
                End If
                If -1 = x Then
                    Return "-1"
                End If
                sign = IIf(x < 0, -1, 1)
                expSigned = IIf(Math.Abs(x) < 1, False, True)
                Dim splitDP As String()
                Dim EShift As Int32 = 0
                If x.ToString().Contains("E") Then
                    ' 0.00012345
                    Dim SplitOnE As String() = x.ToString().Split("E", 2) ' [1.2345], [-4]
                    Int32.TryParse(SplitOnE(1), EShift) '-4
                    EShift *= -1
                    Dim minusE As String = SplitOnE(0)            ' 1.2345
                    splitDP = minusE.Split(".", 2)
                    If sign = -1 Then
                        floatLen -= 1
                    End If
                Else
                    splitDP = x.ToString().Split(".", 2)
                End If
                floatLen += splitDP(splitDP.Length - 1).Length + EShift
                Dim i As Int32 = 0
                If myLibIsInt(x) Then
                    floatLen = 0
                End If
                If Math.Abs(x) >= 1.0 And Math.Abs(x) < 10.0 Then
                    Return x.ToString()
                End If
                x = Math.Abs(x)
                If x < 1 Then
                    While x < 1.0
                        x *= 10.0
                        shifts += 1
                        floatLen -= 1
                    End While
                Else
                    While x >= 10.0
                        x /= 10.0
                        shifts += 1
                        floatLen += 1
                    End While
                End If
            Catch ex As Exception
                myLog("myLibToStandardForm - " + ex.Message)
                Throw
            End Try

            Return (x * sign).ToString("F" + floatLen.ToString()) + "*10" + myLibToSup(IIf(expSigned, "-", "") + shifts.ToString())
            'StandardFormCache[orig] = (x * sign).ToString("F" + floatLen.ToString()) + "*10" + myLibToSup((expSigned ? "-" : "") + shifts.ToString());
            'Return StandardFormCache[orig];

        End Function
    End Class
End Namespace