' =================================================================================
' MACRO DE DIAGNOSTICARE FINAL - Metoda "Iterare For?ata dupa Activare"
' Versiune cu Filtru pe Numele Foilor
' =================================================================================

Option Explicit

Sub AuditByForcedIteration_Filtered()
    Dim AUDIT_FILE_PATH As String
    AUDIT_FILE_PATH = Environ("USERPROFILE") & "\Desktop\CATIA_Audit_Report_Filtered.txt"

    Dim CATIA As Object
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application")
    If CATIA Is Nothing Then MsgBox "Nu s-a putut conecta la CATIA.", vbCritical: Exit Sub
    On Error GoTo 0
    CATIA.StatusBar = "Pornire audit prin 'Iterare For?ata' (filtrat)..."
    CATIA.DisplayFileAlerts = False

    Dim fso As Object, reportFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set reportFile = fso.CreateTextFile(AUDIT_FILE_PATH, True)
    
    reportFile.WriteLine "======= RAPORT DE AUDIT PRIN 'ITERARE FOR?ATA' (Foi filtrate) ======="
    reportFile.WriteLine "Generat la: " & Now
    reportFile.WriteLine "========================================================================" & vbNewLine

    Dim doc As Object, drawingCount As Integer
    drawingCount = 0
    For Each doc In CATIA.Documents
        If TypeName(doc) = "DrawingDocument" Then
            drawingCount = drawingCount + 1
            reportFile.WriteLine "--- DESEN " & drawingCount & ": " & doc.Name & " ---"
            
            Dim aSheet As Object
            For Each aSheet In doc.Sheets
                ' <<< FILTRUL ESTE AICI: Procesam doar foile care încep cu "Sheet" >>>
                If Left(aSheet.Name, 5) = "Sheet" Then
                    ' Activam foaia pentru a for?a încarcarea datelor
                    On Error Resume Next
                    aSheet.Activate
                    On Error GoTo 0
                    
                    reportFile.WriteLine "  -- Foaie: '" & aSheet.Name & "'"
                    
                    On Error Resume Next
                    Dim viewCount As Long
                    viewCount = aSheet.Views.Count
                    reportFile.WriteLine "    -> Numar de vederi gasite în .Views: " & viewCount
                    If Err.Number <> 0 Then
                        reportFile.WriteLine "    -> Eroare la accesarea colec?iei .Views"
                        Err.Clear
                    End If
                    On Error GoTo 0
                    
                    If viewCount > 0 Then
                        Dim aView As Object
                        For Each aView In aSheet.Views
                            reportFile.WriteLine "      - Procesare Vedere: '" & aView.Name & "' (Tip: " & TypeName(aView) & ")"
                            ListElementsFromViewRobust aView, reportFile
                        Next aView
                    End If
                    reportFile.WriteLine "  -------------------------------------"
                End If
            Next aSheet
            reportFile.WriteLine ""
        End If
    Next doc
    
    reportFile.Close
    CATIA.DisplayFileAlerts = True
    CATIA.StatusBar = "Audit finalizat."
    MsgBox "Raportul final (Iterare For?ata, Filtrata) a fost salvat pe Desktop:" & vbNewLine & "CATIA_Audit_Report_Filtered.txt", vbInformation, "Audit Finalizat"
End Sub

Private Sub ListElementsFromViewRobust(targetView As Object, ByRef outFile As Object)
    If targetView Is Nothing Then Exit Sub
    
    Dim allCollections As Variant
    allCollections = Array("GeometricElements", "Dimensions", "Texts") ' Numele colec?iilor ca string
    
    Dim collectionName As Variant
    For Each collectionName In allCollections
        Dim currentCollection As Object
        On Error Resume Next
        ' Încercam sa accesam colec?ia dupa nume
        Set currentCollection = CallByName(targetView, collectionName, VbGet)
        
        If Err.Number <> 0 Then
            outFile.WriteLine "          (Eroare la accesarea colec?iei ." & collectionName & ")"
            Err.Clear
        ElseIf Not currentCollection Is Nothing Then
            If currentCollection.Count > 0 Then
                outFile.WriteLine "        -> Gasit în ." & collectionName & ":"
                Dim elem As Object
                For Each elem In currentCollection
                    outFile.WriteLine "          - Tip: " & TypeName(elem) & ", Nume: " & elem.Name
                Next elem
            End If
        End If
        On Error GoTo 0
        Set currentCollection = Nothing
    Next collectionName
End Sub
