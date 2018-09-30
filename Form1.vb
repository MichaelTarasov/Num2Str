Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Number As Double
        Number = CDbl(TextBox1.Text)
        TextBox2.Text = Num2Str(Number)
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        Dim Number As Double
        If e.KeyCode = Keys.Enter Then
            Number = CDbl(TextBox1.Text)
            TextBox2.Text = Num2Str(Number)
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Convert.ToChar(27) Then
            MessageBox.Show("Выходим...")
            Application.Exit()
        End If
    End Sub

    '    ПРИМЕРЫ ИСПОЛЬЗОВАНИЯ:

    '    Num2Str(3456832.71))
    '    Num2Str(3.2, "м;метр;метра;метров", "м;дециметр;дециметра;дециметров;1,0")
    '    Num2Str(3.71, "м;метр;метра;метров", "м;милиметр;милиметра;милиметров;2,0")
    '    Num2Str(32.102, "ж|тонна|тонны|тонн", "м|килограмм|киллограмма|киллограмм|3 и ещё 0")
    '    Num2Str(32, "м;человек|человека|человек")
    '    Num2Str(21, "с;окно|окна|окон")
    '    Num2Str(21, "ж;дубинка|дубинки|дубинок")
    '    Num2Str(21, "ж;бутылка молока|бутылки молока|бутылок молока")
    '    Num2Str(1277, "ж;бутылка молока|бутылки молока|бутылок молока")

    '    'Три миллиона четыреста пятьдесят шесть тысяч восемьсот тридцать два рубля 71 копейка 
    '    'Три метра, 2 дециметра 
    '    'Три метра, 71 милиметр 
    '    'Тридцать две тонны и ещё 102 киллограмма 
    '    'Тридцать два человека 
    '    'Двадцать одно окно 
    '    'Двадцать одна дубинка 
    '    'Двадцать одна бутылка молока 
    '    'Одна тысяча двести семьдесят семь бутылок молока 



    ' "ж|копейка|копейки|копеек|2,00" 
    '                           ^ - кол-во значащих знаков после запятой 
    '                            ^ - символ-разделитель (пробел если не нужен) 
    '                             ^^ - формат вывода числа дробной части 
    '  ^ - род наименования 
    '    ^^^^^^^ - именительный падеж 
    '            ^^^^^^^ - родительный падеж 
    '                    ^^^^^^ - родительный падеж множественного числа 
    Public Function Num2Str(ByVal xsu As Object,
                         Optional ByVal PString1 As String = "м|рубль|рубля|рублей",
                         Optional ByVal PString2 As String = vbNullString) As String

        Dim ssu As String, nsu As Byte, edi As Byte, des As Byte, sot As Byte, ind As Byte, i As Integer, v() As String, j As Integer, sb As New System.Text.StringBuilder
        Dim r1 As String = "м", r10 As String = vbNullString, r11 As String = vbNullString, r12 As String = vbNullString
        Dim r2 As String = vbNullString, r20 As String = vbNullString, r21 As String = vbNullString, r22 As String = vbNullString,
            r2_ As String = vbNullString, r2n As Short = 2, r2s As String = "00"

        On Error GoTo Err_
        If Not IsNumeric(xsu) Then Num2Str = vbNullString : Exit Function
        If xsu >= 10000000000000 Then Num2Str = "Слишком большое число" : Exit Function

        If PString1 Is Nothing Then
            PString2 = vbNullString
        Else
            PString1 = PString1.ToLower.Replace(";", "|")
            If Not PString2 Is Nothing Then PString2 = PString2.ToLower.Replace(";", "|")
            v = PString1.Split("|")
            If v.Length >= 4 Then
                If 0 = PString1.CompareTo("м|рубль|рубля|рублей") And PString2 Is Nothing Then
                    PString2 = "ж|копейка|копейки|копеек|2 00"
                End If
                r1 = v(0).Substring(0, 1)
                r10 = v(1)
                r11 = v(2)
                r12 = v(3)
            End If
        End If

        If Not PString2 Is Nothing Then
            v = PString2.Split("|")
            If v.Length = 4 Or v.Length = 5 Then
                r2 = v(0).Substring(0, 1)
                r20 = v(1)
                r21 = v(2)
                r22 = v(3)
                If v.Length = 5 Then
                    r2n = CShort(v(4).Substring(0, 1))
                    r2_ = v(4).Substring(1, 1).Trim
                    r2s = v(4).Substring(2).Trim
                    If r2s.Length = 0 Then r2s = "0"
                End If
            End If
        End If

        If Fix(xsu) = 0 Then
            sb.Append("ноль " & r12 & " ")
        Else
            If xsu < 0 Then sb.Append("минус ")
            ssu = Fix(System.Math.Abs(xsu)).ToString       ' строка рублей без знака 
            nsu = (ssu.Length + 2) \ 3                     ' количество троек цифр 
            ssu = Microsoft.VisualBasic.Strings.Right$("00", nsu * 3 - ssu.Length) + ssu ' добавляем нулями 
            For i = nsu To 1 Step -1
                j = (nsu - i) * 3
                sot = CByte(ssu.Substring(j, 1))     ' сотни 
                des = CByte(ssu.Substring(j + 1, 1)) ' десятки 
                edi = CByte(ssu.Substring(j + 2, 1)) ' единицы 
                If sot + des + edi > 0 Or i = 1 Then
                    If sot > 0 Then
                        sb.Append(Choose(sot, "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот") + " ")
                    End If
                    If des = 1 Then
                        sb.Append(Choose(edi + 1, "десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать") + " ")
                        ind = 3
                    Else
                        If des <> 0 Then
                            sb.Append(Choose(des - 1, "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто") + " ")
                        End If
                        If edi <> 0 Then ' вычисляем индекс для тысяч (одна,две) 
                            ind = IIf(i = 2 And (edi = 1 Or edi = 2), 9, 0)
                            Select Case r1
                                Case "м" : sb.Append(Choose(edi + ind, "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять", "одна", "две") + " ")
                                Case "ж" : sb.Append(Choose(edi + ind, "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять", "одна", "две") + " ")
                                Case Else : sb.Append(Choose(edi + ind, "одно", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять", "одна", "две") + " ")
                            End Select
                        End If
                        Select Case edi
                            Case 1 : ind = 1
                            Case 2 To 4 : ind = 2
                            Case Else : ind = 3
                        End Select
                    End If
                    sb.Append(Choose((i - 1) * 3 + ind, r10, r11, r12, "тысяча", "тысячи", "тысяч", "миллион", "миллиона", "миллионов", "миллиард", "миллиарда", "миллиардов", "триллион", "триллиона", "триллионов") & " ")
                End If
            Next i
        End If
        If Not r2 Is Nothing Then
            ssu = Microsoft.VisualBasic.Strings.Right(Format(xsu, ".000".Substring(0, r2n + 1)), r2n)
            If r2n > 1 Then des = CByte(ssu.Substring(r2n - 2, 1)) Else des = 0
            edi = CByte(ssu.Substring(r2n - 1, 1))
            xsu = CShort((xsu - Fix(xsu)) * (10 ^ r2n))
            If des = 1 Then
                ind = 3
            Else
                Select Case edi
                    Case 1 : ind = 1
                    Case 2 To 4 : ind = 2
                    Case Else : ind = 3
                End Select
            End If
            If r2.Length > 0 And Not r2_ Is Nothing Then
                If r2_.Length > 0 And sb.Length Then sb.Insert(sb.Length - 1, r2_)
            End If
            sb.Append(Format(xsu, r2s) & " " & Choose(ind, r20, r21, r22))
        End If
        Num2Str = sb.ToString.TrimEnd : sb = Nothing
        Mid(Num2Str, 1, 1) = Mid(Num2Str, 1, 1).ToUpper
        Exit Function

Err_:
        Num2Str = "Ошибка числа прописью"
    End Function


End Class
