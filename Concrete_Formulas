Sub Macro_Concrete_Formulas()
' Makes all cell references in selected area concrete references e.g. A1 --> $A$1

Dim c As Range
For Each c In Selection
    c.Formula = Application.ConvertFormula(c.Formula, xlA1, , xlAbsolute)
Next
    
End Sub

