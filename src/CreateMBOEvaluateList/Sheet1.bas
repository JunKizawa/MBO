Private Sub cmdOutputMBOList_Click()
    
    Call CreateMboList
    
End Sub

Private Sub cmdGetPrevMBO_Click()
    
    Dim wsTarget As Worksheet
    Set wsTarget = GetLatestMboSheet(ThisWorkbook)
    
    If wsTarget Is Nothing Then
        MsgBox "MBO 一覧シートが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    Call UpdatePrevTermEvaluation(wsTarget)
    
End Sub

