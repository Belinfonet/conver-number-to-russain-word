Private Function CONVERT_NUMBER_TO_WORDS_FRM(Summa As Double) As String
  CONVERT_NUMBER_TO_WORDS_FRM = Format(Abs(Summa), "000000000")
End Function

Public Function CONVERT_NUMBER_TO_3WORDS(ThreeNumbers As String, Obj1 As String, Obj2 As String, Obj5 As String, Genus As Integer) As String
  CONVERT_NUMBER_TO_3WORDS = ""
  If ThreeNumbers <> "000" Then
    Dim d1 As Integer, d2 As Integer, d3 As Integer
    d1 = Val(Mid(ThreeNumbers, 1, 1))
    d2 = Val(Mid(ThreeNumbers, 2, 1))
    d3 = Val(Mid(ThreeNumbers, 3, 1))
    If d1 > 0 Then
      CONVERT_NUMBER_TO_3WORDS = Choose(d1, "сто ", "двести ", "триста ", "четыреста ", "пятьсот ", "шестьсот ", "семьсот ", "восемьсот ", "девятьсот ")
    End If
    If d2 = 1 Then
      If d3 = 0 Then
        CONVERT_NUMBER_TO_3WORDS = CONVERT_NUMBER_TO_3WORDS & "десять "
      Else
        CONVERT_NUMBER_TO_3WORDS = CONVERT_NUMBER_TO_3WORDS & Choose(d3, "один", "две", "три", "четыр", "пят", "шест", "сем", "восем", "девят") & "надцать "
      End If
      CONVERT_NUMBER_TO_3WORDS = CONVERT_NUMBER_TO_3WORDS & Obj5 & " "
    Else
      If d2 > 0 Then
        CONVERT_NUMBER_TO_3WORDS = CONVERT_NUMBER_TO_3WORDS & Choose(d2 - 1, "двадцать ", "тридцать ", "сорок ", "пятьдесят ", "шестьдесят ", "семьдесят ", "восемьдесят ", "девяносто ")
      End If
      Select Case d3
        Case 1:
          CONVERT_NUMBER_TO_3WORDS = CONVERT_NUMBER_TO_3WORDS & Choose(Genus, "один ", "одна ", "одно ") & Obj1 & " "
        Case 2 To 4:
          If d3 = 2 Then
            If Genus = 2 Then
              CONVERT_NUMBER_TO_3WORDS = CONVERT_NUMBER_TO_3WORDS & "две "
            Else
              CONVERT_NUMBER_TO_3WORDS = CONVERT_NUMBER_TO_3WORDS & "два "
            End If
          Else
            CONVERT_NUMBER_TO_3WORDS = CONVERT_NUMBER_TO_3WORDS & Choose(d3 - 2, "три ", "четыре ")
          End If
          CONVERT_NUMBER_TO_3WORDS = CONVERT_NUMBER_TO_3WORDS & Obj2 & " "
        Case Else:
          If d3 > 0 Then
            CONVERT_NUMBER_TO_3WORDS = CONVERT_NUMBER_TO_3WORDS & Choose(d3 - 4, "пять ", "шесть ", "семь ", "восемь ", "девять ")
          End If
          CONVERT_NUMBER_TO_3WORDS = CONVERT_NUMBER_TO_3WORDS & Obj5 & " "
      End Select
    End If
  End If
End Function

Public Function CONVERT_NUMBER_TO_WORDS(Number As Double) As String
  Dim Obj1 As String
  Obj1 = "рубль"
  
  Dim Obj2 As String
  Obj2 = "рубля"
  
  Dim Obj5 As String
  Obj5 = "рублей"
  
  Dim Genus As Integer
  Genus = 1
  
  Dim BigLetter As Boolean
  BigLetter = false
  
  Dim Str As String

  Dim Img As String
  Img = CONVERT_NUMBER_TO_WORDS_FRM(Number)
  Str = CONVERT_NUMBER_TO_3WORDS(Mid(Img, 1, 3), "миллион", "миллиона", "миллионов", 1) + CONVERT_NUMBER_TO_3WORDS(Mid(Img, 4, 3), "тысяча", "тысячи", "тысяч", 2) + CONVERT_NUMBER_TO_3WORDS(Mid(Img, 7, 3), Obj1, Obj2, Obj5, Genus)
      
  If Len(Str) = 0 Then
    Str = "ноль " & Obj5 & " "
  Else
    If Mid(Img, 7, 3) = "000" Then
      Str = Str & Obj5 & " "
    End If
  End If
  If (Number < 0) Then
    Str = "минус " & Str
  End If
  If (BigLetter) Then
    Str = UCase(Mid(Str, 1, 1)) & Mid(Str, 2)
  End If
  CONVERT_NUMBER_TO_WORDS = Str
End Function
