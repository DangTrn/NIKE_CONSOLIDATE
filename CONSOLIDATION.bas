Attribute VB_Name = "CONSOLIDATION"
Function sumarray(arr As Variant) As Integer
    sumarray = 0
    For i = LBound(arr, 2) To UBound(arr, 2)
        sumarray = sumarray + arr(8, i)
    Next i
End Function

Function maxarray(arr As Variant) As Integer
    maxarray = 0
    For i = LBound(arr, 2) To UBound(arr, 2)
        maxarray = Application.WorksheetFunction.Max(maxarray, arr(8, i))
    Next i
End Function

Function countarray(arr As Variant) As Integer
    countarray = 0
    For i = LBound(arr, 2) To UBound(arr, 2)
        If arr(8, i) > 0 Then
            countarray = countarray + 1
        End If
    Next i
End Function

Sub sort(list As ListObject)
    With list.sort
        With .SortFields
            .Clear
            .Add2 _
                key:=Range("STOCK_DETAIL_BY_UPC[SKU]"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .Add2 _
                key:=Range("STOCK_DETAIL_BY_UPC[GROUP]"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .Add2 _
                key:=Range("STOCK_DETAIL_BY_UPC[8W_SOLD]"), _
                SortOn:=xlSortOnValues, _
                Order:=xlDescending, _
                DataOption:=xlSortNormal
            .Add2 _
                key:=Range("STOCK_DETAIL_BY_UPC[STORE RANK]"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .Add2 _
                key:=Range("STOCK_DETAIL_BY_UPC[UPC]"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
        End With
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Function check(storein As Variant, storeout As Variant) As Boolean
    Dim arr As Variant
    arr = storein
    For i = 1 To UBound(storein, 2)
        arr(8, i) = arr(8, i) + storeout(8, i)
    Next i
    
    ttl = sumarray(arr)
    MAXQTY = maxarray(arr)
    szcount = countarray(arr)
    ave = ttl / szcount
    
    If ave <= 3 And MAXQTY <= 6 Then
        check = True
    Else: check = False
    End If
End Function

Function move(storein As Variant, storeout As Variant, storecodein As Variant, storecodeout As Variant) As Collection
    Set move = New Collection
    For i = LBound(storein, 2) To UBound(storein, 2)
        storein(8, i) = storein(8, i) + storeout(8, i)
        storeout(8, i) = 0
        storeout(9, i) = storecodein
    Next i
    
    move.Add _
        Item:=storein, _
        key:=storecodein
    move.Add _
        Item:=storeout, _
        key:=storecodeout
End Function

Sub CONSOLIDATE()
    Dim size As Variant
    Dim store As New Scripting.Dictionary
    Dim group As New Scripting.Dictionary
    Dim sku As New Scripting.Dictionary
    Dim list As ListObject
    
    Dim stock As Variant, storein As Variant, storeout As Variant
    
    Dim i As Long, j As Long, t As Long, k As Long, l As Long, m As Long
    Dim lr As Long, tmp As Long
    Dim transfer As Worksheet
      
    Dim StartTime As Double
    Dim MinutesElapsed As String
    
'Remember time when macro starts
    StartTime = Timer
    
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    
    Workbooks("CONSOLIDATE version 2.1.xlsm").RefreshAll
    
    Set list = Worksheets("STOCK_DETAIL_BY_UPC").ListObjects("STOCK_DETAIL_BY_UPC")
    Call sort(list)
    stock = list.DataBodyRange.Value
    
'nap stock vao multiple levels dictionary SKU
    tmp = 1
    For i = tmp To UBound(stock) - 1
        Set group = New Scripting.Dictionary
        For j = i To UBound(stock) - 1
            Set store = New Scripting.Dictionary
            For k = j To UBound(stock) - 1
                ReDim size(9, 0)
                For l = k To UBound(stock) - 1
                    t = UBound(size, 2) + 1
                    ReDim Preserve size(9, t)
                    For m = 4 To 12
                        size(m - 3, t) = stock(l, m)
                    Next m
                    If stock(l + 1, 3) <> stock(l, 3) Or stock(l + 1, 2) <> stock(l, 2) Or stock(l + 1, 1) <> stock(l, 1) Then
                        tmp = l
                        Exit For
                    End If
                Next l
                k = tmp
                store.Add stock(k, 3), size
                If stock(k + 1, 2) <> stock(k, 2) Or stock(k + 1, 1) <> stock(k, 1) Then
                    Exit For
                End If
            Next k
            j = tmp
            group.Add stock(j, 2), store
            If stock(j + 1, 1) <> stock(j, 1) Then
                Exit For
            End If
        Next j
        i = tmp
        sku.Add stock(i, 1), group
    Next i
    
'transfering
    For Each sk In sku
        For Each gr In sku(sk)
            Set store = sku(sk)(gr)
            For i = 0 To store.Count - 2
                For j = store.Count - 1 To i + 1 Step -1
                    storein = sku(sk)(gr)(store.Keys(i))
                    storeout = sku(sk)(gr)(store.Keys(j))
                    If sumarray(storein) * sumarray(storeout) > 0 And storeout(3, 1) > 3 Then
                        If check(storein, storeout) = True Then
                             sku(sk)(gr)(store.Keys(i)) = move(storein, storeout, store.Keys(i), store.Keys(j))(1)
                             sku(sk)(gr)(store.Keys(j)) = move(storein, storeout, store.Keys(i), store.Keys(j))(2)
                        End If
                    End If
                Next j
            Next i
        Next
    Next
        

'fill dictionary SKU ra sheet TRANSFERLIST
    Set transfer = ActiveWorkbook.Worksheets("TRANSFERLIST")
    transfer.Cells.ClearContents
    
    list.HeaderRowRange.Copy
    transfer.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    
    lr = 2
    For Each sk In sku
        For Each gr In sku(sk)
            For Each st In sku(sk)(gr)
                For i = 1 To UBound(sku(sk)(gr)(st), 2)
                    transfer.Cells(lr, 1) = sk
                    transfer.Cells(lr, 2) = gr
                    transfer.Cells(lr, 3) = st
                    For m = 4 To 12
                        transfer.Cells(lr, m) = sku(sk)(gr)(st)(m - 3, i)
                    Next m
                    lr = lr + 1
                Next i
            Next st
        Next gr
    Next sk
    
    transfer.Range("A1:L1").EntireColumn.AutoFit
    
    Workbooks("CONSOLIDATE version 2.1.xlsm").Save
    
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
    
    
    MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
    MsgBox "Congratulation! It's finally done." & vbNewLine & "It's take " & MinutesElapsed & " to finish", vbInformation
End Sub

