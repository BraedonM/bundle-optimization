Sub BuildOptimizedBundles()
    Dim wsInput As Worksheet, wsOutput As Worksheet
    Dim lastRow As Long, outRow As Long
    Dim dictItems As Object, dictBundle As Object
    Dim i As Long, qty As Long, key As Variant
    Dim orderID As Variant, sku As Variant
    Dim unitWidth As Double, unitHeight As Double, unitLength As Double, unitWeight As Double
    Dim maxWidth As Double: maxWidth = 590
    Dim maxHeight As Double: maxHeight = 590
    Dim currentWidth As Double, currentHeight As Double
    Dim bundleNum As Long
    Dim parts() As String
    Dim loopOrderID As Variant

    Set wsInput = ThisWorkbook.Sheets("SO_Input")
    Set wsOutput = ThisWorkbook.Sheets("Optimized_Bundles")
    Set dictItems = CreateObject("Scripting.Dictionary")

    ' Clear output
    wsOutput.Cells.ClearContents
    wsOutput.Range("A1:L1").Value = Array("Order ID", "Bundle No.", "SKU", "Qty", "SKU Description", _
                                          "SKU Width (mm)", "SKU Height (mm)", "SKU Weight (kg)", _
                                          "Bundle Total Width", "Bundle Total Height", "Bundle Total Weight", "Note")

    ' Load data into dictionary
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        orderID = wsInput.Cells(i, 1).Value
        sku = wsInput.Cells(i, 2).Value
        qty = wsInput.Cells(i, 4).Value
        unitWidth = wsInput.Cells(i, 5).Value
        unitHeight = wsInput.Cells(i, 6).Value
        unitLength = wsInput.Cells(i, 7).Value
        unitWeight = wsInput.Cells(i, 8).Value

        If qty > 0 Then
            key = orderID & "|" & sku & "|" & unitWidth & "|" & unitHeight & "|" & unitLength & "|" & unitWeight
            If dictItems.exists(key) Then
                dictItems(key) = dictItems(key) + qty
            Else
                dictItems.Add key, qty
            End If
        End If
    Next i

    outRow = 2

    For Each loopOrderID In GetUniqueOrders(dictItems)
        bundleNum = 1
        Do
            currentWidth = 0
            currentHeight = 0
            Set dictBundle = CreateObject("Scripting.Dictionary")
            Dim totalWeight As Double: totalWeight = 0

            For Each key In dictItems.Keys
                If InStr(CStr(key), loopOrderID) = 1 And dictItems(key) > 0 Then
                    If VarType(key) = vbString And InStr(CStr(key), "|") > 0 Then
                        parts = Split(CStr(key), "|")
                        If UBound(parts) >= 5 Then
                            sku = parts(1)
                            unitWidth = CDbl(parts(2))
                            unitHeight = CDbl(parts(3))
                            unitWeight = CDbl(parts(5))
                            qty = dictItems(key)
                        Else
                            GoTo SkipKey
                        End If
                    Else
                        GoTo SkipKey
                    End If

                    ' Fill bundle with as many of the same SKU as possible
                    Do While qty > 0 And currentWidth + unitWidth <= maxWidth
                        If unitHeight + currentHeight <= maxHeight Then
                            dictBundle.Add key & "|" & dictBundle.Count, Array(loopOrderID, sku, unitWidth, unitHeight, unitWeight)
                            currentWidth = currentWidth + unitWidth
                            currentHeight = Application.WorksheetFunction.Max(currentHeight, unitHeight)
                            totalWeight = totalWeight + unitWeight
                            qty = qty - 1
                            dictItems(key) = qty
                        Else
                            Exit Do
                        End If
                    Loop

                    If qty > 0 Then Exit For ' Force exit to break row
                End If
SkipKey:
            Next key

            ' Add filler if needed
            If currentHeight < maxHeight Then
                Dim fillerHeight As Double
                fillerHeight = maxHeight - currentHeight
                If fillerHeight > 0 Then
                    Dim fillerSKU As Variant: fillerSKU = "FILLER.100x100"
                    dictBundle.Add "FILLER|" & dictBundle.Count, Array(loopOrderID, fillerSKU, 100, fillerHeight, 0.1)
                    totalWeight = totalWeight + 0.1
                    currentHeight = maxHeight
                End If
            End If

            ' Output bundle
            For Each key In dictBundle.Keys
                On Error Resume Next
                If IsArray(dictBundle(key)) Then
                    parts = dictBundle(key)
                    If Not IsEmpty(parts) And UBound(parts) >= 4 Then
                        On Error GoTo 0
                        With wsOutput
                            .Cells(outRow, 1).Value = parts(0)              ' Order ID
                            .Cells(outRow, 2).Value = bundleNum             ' Bundle No.
                            .Cells(outRow, 3).Value = parts(1)              ' SKU
                            .Cells(outRow, 4).Value = 1                     ' Qty (1 per row)
                            .Cells(outRow, 5).Value = parts(1)              ' SKU Description
                            .Cells(outRow, 6).Value = parts(2)              ' Width
                            .Cells(outRow, 7).Value = parts(3)              ' Height
                            .Cells(outRow, 8).Value = parts(4)              ' Weight
                            .Cells(outRow, 9).Value = currentWidth          ' Bundle Total Width
                            .Cells(outRow, 10).Value = currentHeight        ' Bundle Total Height
                            .Cells(outRow, 11).Value = totalWeight          ' Bundle Total Weight
                            .Cells(outRow, 12).Value = ""                   ' Note
                        End With
                        outRow = outRow + 1
                    End If
                End If
                On Error GoTo 0
            Next key
            bundleNum = bundleNum + 1
        Loop While AnyLeft(loopOrderID, dictItems)
    Next loopOrderID

    MsgBox "Bundle Optimization Complete!", vbInformation
End Sub

Function GetUniqueOrders(dict As Object) As Collection
    Dim coll As New Collection
    Dim k As Variant, id As Variant
    On Error Resume Next
    For Each k In dict.Keys
        id = Split(CStr(k), "|")(0)
        coll.Add id, id
    Next k
    On Error GoTo 0
    Set GetUniqueOrders = coll
End Function

Function AnyLeft(orderID As Variant, dict As Object) As Boolean
    Dim k As Variant
    For Each k In dict.Keys
        If InStr(k, orderID) = 1 And dict(k) > 0 Then
            AnyLeft = True
            Exit Function
        End If
    Next k
    AnyLeft = False
End Function
