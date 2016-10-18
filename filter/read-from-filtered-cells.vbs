'表示中のセルのみを範囲とする
Sub ReadFromFilteredCellsSample
    'オートフィルタなどをかけた後で、表示しているすべてのセルの集合
    Dim rngFilteredAll As Range
    Set rngFilteredAll = Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible)
    
    'B列すべてのセルの集合
    Dim rngColB As Range
    Set rngColB = Range("B:B")
    
    'オートフィルタで表示中、かつB列で共通のセルの集合を取得
    Dim rngCross As Range
    Set rngCross = Application.Intersect(rngFilteredAll, rngColB)
    
    For Each cell In rngCross
       MsgBox (cell.Value2)
    Next

    
    'ちなみに・・・某VBA講師の解説では、ActiveSheet.AutoFilter.Range(ActiveSheet.AutoFilter.Range.Count).Row をセル下端としていた。
    'そういえば、セルが下端限界まで伸びて削除しても戻らない現象がしょっちゅう起きる。
    'こうしないと重くなるんだろうか？
End Sub