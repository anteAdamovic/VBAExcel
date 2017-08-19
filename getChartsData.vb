Dim redovniStudenti As New Collection
Dim izvanredniStudenti As New Collection
Dim studentData As New Collection

Function getChartsData()
    On Error Resume Next

    Dim range As range
    Dim id As String
    Dim val As String

    Dim redovniStudentiSheet, izvanredniStudentiSheet, vjezbeSheet, blitzSheet, kolokvij1Sheet, kolokvij2Sheet, rok1Sheet, rok2Sheet As Worksheet
    Dim redovniCount, izvanredniCount As Integer
    
    ' Save all required sheets in variables
    Set redovniStudentiSheet = ActiveWorkbook.Sheets("RedovniStudenti")
    Set izvanredniStudentiSheet = ActiveWorkbook.Sheets("IzvanredniStudenti")
    Set vjezbeSheet = ActiveWorkbook.Sheets("Vjezbe")
    Set blitzSheet = ActiveWorkbook.Sheets("Blic")
    Set kolokvij1Sheet = ActiveWorkbook.Sheets("Kolokvij1")
    Set kolokvij2Sheet = ActiveWorkbook.Sheets("Kolokvij2")
    Set rok1Sheet = ActiveWorkbook.Sheets("1 ROK")
    Set rok2Sheet = ActiveWorkbook.Sheets("2 ROK")
    
    redovniCount = 0
    izvanredniCount = 0
   
    ' Calculate redovni studenti count
    For Each range In redovniStudentiSheet.Rows
        If range.Row > 2 Then
            If redovniStudentiSheet.Cells(range.Row, 1).val = "" Then
                Exit For
            Else
                redovniCount = redovniCount + 1
                id = redovniStudentiSheet.Cells(range.Row, 3).val + " " + redovniStudentiSheet.Cells(range.Row, 4).val
                redovniStudenti.Add studentData, id
            End If
        End If

    Next range

    Debug.Print "Redovni studenti: ", redovniCount

    ' Calculate izvanredni studenti count
    For Each range In izvanredniStudentiSheet.Rows
        If range.Row > 2 Then
            If izvanredniStudentiSheet.Cells(range.Row, 1).val = "" Then
                Exit For
            Else
                izvanredniCount = izvanredniCount + 1
                id = izvanredniStudentiSheet.Cells(range.Row, 3).val + " " + izvanredniStudentiSheet.Cells(range.Row, 4).val
                izvanredniStudenti.Add studentData, id
            End If
        End If

    Next range

    Debug.Print "Izvanredni studenti: ", izvanredniCount
    
    Dim excerciseCount As Integer
    excerciseCount = 0

    ' Fetch excercise marks
    For Each range In vjezbeSheet.Rows

        If range.Row > 1 Then
            If vjezbeSheet.Cells(range.Row, 1).Value = "" Then
                Exit For
            End If
            excerciseCount = excerciseCount + 1
            id = vjezbeSheet.Cells(range.Row, 3) + " " + vjezbeSheet.Cells(range.Row, 2)
            val = vjezbeSheet.Cells(range.Row, 16).Value
            If val = Empty Then
                val = "0"
            End If

            If collectionContains(redovniStudenti, id) Then
                ' Debug.Print "Found " + id + " in redovniStudenti, val is ", val
                Set studentData = redovniStudenti.Item(id)
                studentData.Remove ("vjezbe")
                studentData.Add val, "vjezbe"
                redovniStudenti.Remove (id)
                redovniStudenti.Add studentData, id
                ' Debug.Print redovniStudenti.Item(id)("vjezbe")
                  
            ElseIf collectionContains(izvanredniStudenti, id) Then
                ' Debug.Print "Found " + id + " in izvanredniStudenti, val is ", val
                Set studentData = izvanredniStudenti.Item(id)
                studentData.Remove ("vjezbe")
                studentData.Add val, "vjezbe"
                izvanredniStudenti.Remove (id)
                izvanredniStudenti.Add studentData, id
                ' Debug.Print izvanredniStudenti.Item(id)("vjezbe")

            End If
        End If

    Next range

    Debug.Print "Got data about", excerciseCount, "excercises"

    Dim blitzCount As Integer
    blitzCount = 0

    ' Fetch excercise marks
    For Each range In blitzSheet.Rows

        If range.Row > 1 Then
            If blitzSheet.Cells(range.Row, 1).Value = "" Then
                Exit For
            End If
            blitzCount = blitzCount + 1
            id = blitzSheet.Cells(range.Row, 3) + " " + blitzSheet.Cells(range.Row, 2)
            val = blitzSheet.Cells(range.Row, 11).Value
            If val = Empty Then
                val = "0"
            End If

            If collectionContains(redovniStudenti, id) Then
                ' Debug.Print "Found " + id + " in redovniStudenti, val is ", val
                Set studentData = redovniStudenti.Item(id)
                studentData.Remove ("blic")
                studentData.Add val, "blic"
                redovniStudenti.Remove (id)
                redovniStudenti.Add studentData, id
                ' Debug.Print redovniStudenti.Item(id)("vjezbe")
                  
            ElseIf collectionContains(izvanredniStudenti, id) Then
                ' Debug.Print "Found " + id + " in izvanredniStudenti, val is ", val
                Set studentData = izvanredniStudenti.Item(id)
                studentData.Remove ("blic")
                studentData.Add val, "blic"
                izvanredniStudenti.Remove (id)
                izvanredniStudenti.Add studentData, id
                ' Debug.Print izvanredniStudenti.Item(id)("vjezbe")

            End If
        End If

    Next range

    Debug.Print "Got data about", blitzCount, "excercises"

    Dim kolokvij1Count As Integer
    kolokvij1Count = 0

    ' Fetch kolokvij 1
    For Each range In kolokvij1Sheet.Rows

        If range.Row > 1 Then
            If kolokvij1Sheet.Cells(range.Row, 1).Value = "" Then
                Exit For
            End If
            kolokvij1Count = kolokvij1Count + 1
            id = kolokvij1Sheet.Cells(range.Row, 3) + " " + kolokvij1Sheet.Cells(range.Row, 2)
            val = kolokvij1Sheet.Cells(range.Row, 4).Value
            If val = Empty Then
                val = "0"
            End If

            If collectionContains(redovniStudenti, id) Then
                ' Debug.Print "Found " + id + " in redovniStudenti, val is ", val
                Set studentData = redovniStudenti.Item(id)
                studentData.Remove ("kolokvij1")
                studentData.Add val, "kolokvij1"
                redovniStudenti.Remove (id)
                redovniStudenti.Add studentData, id
                ' Debug.Print redovniStudenti.Item(id)("vjezbe")
                  
            ElseIf collectionContains(izvanredniStudenti, id) Then
                ' Debug.Print "Found " + id + " in izvanredniStudenti, val is ", val
                Set studentData = izvanredniStudenti.Item(id)
                studentData.Remove ("kolokvij1")
                studentData.Add val, "kolokvij1"
                izvanredniStudenti.Remove (id)
                izvanredniStudenti.Add studentData, id
                ' Debug.Print izvanredniStudenti.Item(id)("vjezbe")

            End If
        End If

    Next range

    Debug.Print "Got data about", kolokvij1Count, "kolokvij1Count"

    Dim numbers, t
    numbers = Array(10, 20, 30)

    For Index = 0 To UBound(numbers)
        t = numbers(Index)
    Next
    getRowData = t


End Function

Function collectionContains(col As Collection, key As String)
    On Error Resume Next

    Dim flag As Boolean
    Dim data

    flag = False
    Set data = col.Item(key)

    If (data) Then
        flag = True
    End If

    collectionContains = flag
End Function
