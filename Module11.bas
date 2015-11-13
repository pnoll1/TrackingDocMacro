Attribute VB_Name = "Module11"
Sub update()
'update info for jobs that are already in tracking document
'project names in master schedule & tracking must match
'only searches 100 lines in master schedule & tracking document

Dim ms As Workbook, t As Workbook
Dim I As Long, J As Long, K As Long, L As Long
Set ms = Workbooks.Open("E:\Documents\master schedule.xlsx")
Set t = ThisWorkbook

'search Tower Cranes for job names that are in tracking documents amd update the date & crane data

For J = 3 To 100
    For K = 2 To 40
        If ms.Worksheets("Tower Cranes").Cells(J, "G").Value = t.Worksheets("Crane").Cells(K, "B").Value Then
        'HR
        t.Worksheets("Crane").Cells(K, "N").Value = ms.Worksheets("Tower Cranes").Cells(J, "J").Value
        'HUH
        t.Worksheets("Crane").Cells(K, "O").Value = ms.Worksheets("Tower Cranes").Cells(J, "L").Value
        'Crane
        t.Worksheets("Crane").Cells(K, "M").Value = ms.Worksheets("Tower Cranes").Cells(J, "B").Value
        'Base crane
        t.Worksheets("Crane").Cells(K, "G").Value = ms.Worksheets("Tower Cranes").Cells(J, "AB").Value
        'Base date
        t.Worksheets("Crane").Cells(K, "H").Value = ms.Worksheets("Tower Cranes").Cells(J, "AC").Value
        'Erect Crane
        t.Worksheets("Crane").Cells(K, "I").Value = ms.Worksheets("Tower Cranes").Cells(J, "AD").Value
        'Erect date
        t.Worksheets("Crane").Cells(K, "J").Value = ms.Worksheets("Tower Cranes").Cells(J, "Q").Value
        'Disman date
        t.Worksheets("Crane").Cells(K, "L").Value = ms.Worksheets("Tower Cranes").Cells(J, "AF").Value
        'Disman crane
        t.Worksheets("Crane").Cells(K, "K").Value = ms.Worksheets("Tower Cranes").Cells(J, "AE").Value
        'Status
        t.Worksheets("Crane").Cells(K, "C").Value = ms.Worksheets("Tower Cranes").Cells(J, "AM").Value
        'Job number
        t.Worksheets("Crane").Cells(K, "E").Value = ms.Worksheets("Tower Cranes").Cells(J, "T").Value
        End If
    Next
Next
'search Hoists for job names that are in tracking documents amd update the date & crane data
'Doesn't work for duals, empty second line info overwrites first line info

For L = 2 To 90
    For M = 2 To 20
        If ms.Worksheets("Hoists").Cells(L, "F").Value = t.Worksheets("Hoist").Cells(M, "B").Value Then
        'Hoist Model
        t.Worksheets("Hoist").Cells(M, "E").Value = ms.Worksheets("Hoists").Cells(L, "A").Value
        '# of Cars
        t.Worksheets("Hoist").Cells(M, "F").Value = ms.Worksheets("Hoists").Cells(L, "B").Value
        'Initial height
        t.Worksheets("Hoist").Cells(M, "G").Value = ms.Worksheets("Hoists").Cells(L, "I").Value
        'Final height
        t.Worksheets("Hoist").Cells(M, "H").Value = ms.Worksheets("Hoists").Cells(L, "J").Value
        '# Jumps
        t.Worksheets("Hoist").Cells(M, "K").Value = ms.Worksheets("Hoists").Cells(L, "K").Value
        '# Gates
        t.Worksheets("Hoist").Cells(M, "L").Value = ms.Worksheets("Hoists").Cells(L, "L").Value
        '# intercoms
        t.Worksheets("Hoist").Cells(M, "M").Value = ms.Worksheets("Hoists").Cells(L, "M").Value
        'Disman date
        t.Worksheets("Hoist").Cells(M, "J").Value = ms.Worksheets("Hoists").Cells(L, "Z").Value
        'Erect date
        t.Worksheets("Hoist").Cells(M, "I").Value = ms.Worksheets("Hoists").Cells(L, "N").Value
        'Status
        t.Worksheets("Hoist").Cells(M, "C").Value = ms.Worksheets("Hoists").Cells(L, "AE").Value
        End If
    Next
Next
End Sub

