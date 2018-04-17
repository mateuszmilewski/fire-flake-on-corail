Attribute VB_Name = "GlobalFooModule"
Public Function getPlantName(plt) As String
    getPlantName = ""
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(FFOC.G_SH_NM_PLT_LIST)
    
    
    Dim r As Range
    Set r = sh.Range("A2")
    Do
    
        If Trim(r.Value) = Trim(plt) Then
            getPlantName = Trim(Replace(Replace(r.Offset(0, 1).Value, "Maestro", ""), "Corail", ""))
            Exit Do
        End If
    
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    
    
    If getPlantName = "" Then getPlantName = plt
End Function
