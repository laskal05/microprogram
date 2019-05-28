Sub add()

  ' 同じシート内でセルを選んで足し算をする

  ' 変数定義
  Dim x As Double, a As Double, b As Double
  a = Range("C2").Value
  b = Range("C4").Value
  x = a + b
  ' ActiveなシートのC6セルに計算結果を挿入する
  Range("C6").Value = x


  ' 複数のシートを指定してセルを選んで足し算をする

  ' 変数定義
  Dim y As Double, c As Double, d As Double
  c = WorkSheetRange("C2").Value
  d = Range("C4").Value
  y = c + d
  ' 1番目のシートのC6セルに計算結果を挿入する
  Worksheets("Sheet2").Range("C6").Value=y
  Set ws = Worksheets(1)

End Sub
