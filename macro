Sub AutoReconciler()
'
' AutoReconciler Macro
' Made by Zack, Jul. 1st 2020
' Reconciles deposits to paycheques in book-keeping
' revision 1.0.3_7-3-20
'
    Dim rng As Range
    Set rng = Application.InputBox("Select a target cell", "Obtain Range Object", Type:=8)
    
    Dim rng1 As Range
    Set rng1 = Application.InputBox("Select a variable range", "Obtain Range Object", Type:=8)
    
    Dim inputData As Range
    Set inputData = Application.InputBox("Enter a targeted value range", "Obtain number Object", Type:=8)
    
    Dim c As Range
    Dim s As Range
    Dim ctr As Integer
    Dim cellsToExclude
    Dim SelectCells As String
    
    For Each s In inputData.Cells
    
        ctr = ctr + 1
        
        
        SolverReset
    
        SolverOk SetCell:=rng.Address, MaxMinVal:=3, ValueOf:=s.Value, ByChange:=rng1.Address _
           , Engine:=2, EngineDesc:="Simplex LP"
        
        SolverAdd CellRef:=rng1.Address, Relation:=5, FormulaText:="binary"
    
        SolverOk SetCell:=rng.Address, MaxMinVal:=3, ValueOf:=s.Value, ByChange:=rng1.Address _
           , Engine:=2, EngineDesc:="Simplex LP"
        
        SolverSolve userFinish:=True
    
        SolverFinish keepFinal:=True
        
        
        SelectCells = ""
        For Each c In rng1.Cells
            If c.Value = 1 Then
                If SelectCells = "" Then
                    SelectCells = CStr(c.Row)
                Else
                    SelectCells = (SelectCells + "," + CStr(c.Row))
                End If
            End If
        Next c
        
        Dim rw As Integer
        Dim cl As Integer
        rw = (inputData.Row + ctr - 1)
        cl = (inputData.Column - 1)
        Cells(rw, cl).Value = SelectCells
        
    Next s
    
End Sub
