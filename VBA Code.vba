' =================================================================================
' MACRO DE DIAGNOSTICARE EXHAUSTIV - Metoda "Iterare For?ata dupa Activare"
' Versiune cu Filtru, Detaliere Extinsa a Elementelor si Proprietatilor
' =================================================================================

Option Explicit

Sub AuditByForcedIteration_Comprehensive()
    Dim AUDIT_FILE_PATH As String
    AUDIT_FILE_PATH = Environ("USERPROFILE") & "\Desktop\CATIA_Audit_Report_Comprehensive.txt"

    Dim CATIA As Object
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application")
    If CATIA Is Nothing Then MsgBox "Nu s-a putut conecta la CATIA.", vbCritical: Exit Sub
    On Error GoTo 0
    CATIA.StatusBar = "Pornire audit COMPREHENSIV prin 'Iterare For?ata' (filtrat, detaliat)..."
    CATIA.DisplayFileAlerts = False

    Dim fso As Object, reportFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' True la final pentru a salva in format Unicode (suporta caractere speciale)
    Set reportFile = fso.CreateTextFile(AUDIT_FILE_PATH, True, True)
    
    reportFile.WriteLine "======= RAPORT DE AUDIT COMPREHENSIV (Foi filtrate, Elemente detaliate) ======="
    reportFile.WriteLine "Generat la: " & Now
    reportFile.WriteLine "Acest raport incearca sa extraga un set extins de proprietati pentru fiecare element."
    reportFile.WriteLine "======================================================================================" & vbNewLine

    Dim doc As Object, drawingCount As Integer
    drawingCount = 0
    For Each doc In CATIA.Documents
        If TypeName(doc) = "DrawingDocument" Then
            drawingCount = drawingCount + 1
            reportFile.WriteLine "--- DESEN " & drawingCount & ": " & doc.Name & " ---"
            
            Dim aSheet As Object
            For Each aSheet In doc.Sheets
                ' <<< FILTRUL ESTE AICI: Procesam doar foile care încep cu "Sheet" >>>
                ' <<< Puteti comenta sau sterge aceasta linie pentru a audita TOATE foile >>>
                If Left(aSheet.Name, 5) = "Sheet" Then
                    On Error Resume Next
                    aSheet.Activate
                    If Err.Number <> 0 Then
                        reportFile.WriteLine "  !! Eroare la activarea foii: '" & aSheet.Name & "'"
                        Err.Clear
                    Else
                        reportFile.WriteLine "  -- Foaie: '" & aSheet.Name & "' (Format: " & aSheet.PaperSize & ", Orientare: " & aSheet.Orientation & ", Scala: " & aSheet.Scale & ")"
                    End If
                    On Error GoTo 0
                    
                    Dim viewCount As Long
                    On Error Resume Next
                    viewCount = aSheet.Views.Count
                    If Err.Number <> 0 Then
                        reportFile.WriteLine "    -> Eroare la accesarea colec?iei .Views"
                        Err.Clear
                    Else
                        reportFile.WriteLine "    -> Numar de vederi gasite în .Views: " & viewCount
                    End If
                    On Error GoTo 0
                    
                    If viewCount > 0 Then
                        Dim aView As Object
                        For Each aView In aSheet.Views
                            reportFile.WriteLine "      - Procesare Vedere: '" & aView.Name & "' (Tip: " & TypeName(aView) & ", Scala: " & aView.Scale & ")"
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
    CATIA.StatusBar = "Audit comprehensiv finalizat."
    MsgBox "Raportul final (Comprehensiv) a fost salvat pe Desktop:" & vbNewLine & "CATIA_Audit_Report_Comprehensive.txt", vbInformation, "Audit Finalizat"
End Sub

Private Sub ListElementsFromViewRobust(targetView As Object, ByRef outFile As Object)
    If targetView Is Nothing Then Exit Sub
    
    ' Lista extinsa de colectii de elemente de cautat in fiecare vedere
    Dim allCollections As Variant
    allCollections = Array("GeometricElements", "Dimensions", "Texts", "Tables", "Pictures", "Welds", "Arrows")
    
    Dim collectionName As Variant
    For Each collectionName In allCollections
        Dim currentCollection As Object
        On Error Resume Next
        Set currentCollection = CallByName(targetView, collectionName, VbGet)
        
        If Err.Number <> 0 Then
            outFile.WriteLine "        -> (Colectia ." & collectionName & " nu a putut fi accesata sau nu exista in aceasta vedere)"
            Err.Clear
        ElseIf Not currentCollection Is Nothing Then
            If currentCollection.Count > 0 Then
                outFile.WriteLine "        -> Gasit " & currentCollection.Count & " element(e) in ." & collectionName & ":"
                Dim elem As Object
                For Each elem In currentCollection
                    outFile.WriteLine "          - Tip: " & TypeName(elem) & ", Nume: " & elem.Name
                    outFile.Write GetElementDetails(elem) ' Adaugam detaliile specifice
                    outFile.Write GetGraphicProperties(elem) ' Adaugam proprietatile grafice comune
                Next elem
            End If
        End If
        On Error GoTo 0
        Set currentCollection = Nothing
    Next collectionName
End Sub

'================================================================================
' FUNCTIE EXTINSA pentru extragerea si formatarea detaliilor unui element
'================================================================================
Private Function GetElementDetails(ByVal elem As Object) As String
    Dim details As String: details = ""
    Dim tempArr(1)
    
    On Error Resume Next ' Gestionare robusta a erorilor pentru proprietati inexistente
    
    Select Case TypeName(elem)
        Case "DrawingDimension"
            Dim dimValue As Object, tolType As Long, upTol As Double, lowTol As Double
            Set dimValue = elem.GetValue
            details = details & "            > Valoare Masurata: " & FormatNumber(elem.GetValue.Value, 3) & " " & elem.GetDimExtremity(1).GetUnitSymbol & vbCrLf
            details = details & "            > Tip Cota: " & elem.DimType & vbCrLf
            details = details & "            > Text Prefix: " & dimValue.GetBaultText(1) & ", Sufix: " & dimValue.GetBaultText(2) & vbCrLf
            details = details & "            > Text Superior: " & dimValue.GetUpText(1) & ", Inferior: " & dimValue.GetDownText(1) & vbCrLf
            
            tolType = dimValue.GetToleranceType
            If tolType <> 0 Then ' catToleranceNoTolerance = 0
                dimValue.GetTolerance upTol, lowTol
                details = details & "            > Toleranta: Sup: " & FormatNumber(upTol, 3) & ", Inf: " & FormatNumber(lowTol, 3) & " (Tip: " & tolType & ")" & vbCrLf
            End If

        Case "DrawingText"
            elem.GetAnchoringPosition tempArr
            details = details & "            > Text: """ & elem.Text & """" & vbCrLf
            details = details & "            > Font: " & elem.GetFontName(0, 0) & ", Marime: " & elem.FONTSIZE & ", Unghi: " & FormatNumber(elem.Angle, 2) & " deg" & vbCrLf
            details = details & "            > Coordonate Ancorare (x,y): (" & FormatNumber(tempArr(0), 2) & ", " & FormatNumber(tempArr(1), 2) & ")" & vbCrLf
            details = details & "            > Tip Chenar: " & elem.FrameType & vbCrLf
            If elem.HasLeader Then
                details = details & "            > Are linie de indicatie (Leader)." & vbCrLf
            End If

        Case "DrawingPoint"
            elem.GetCoordinates tempArr
            details = details & "            > Coordonate (x,y): (" & FormatNumber(tempArr(0), 2) & ", " & FormatNumber(tempArr(1), 2) & ")" & vbCrLf

        Case "DrawingLine"
            Dim startCoords(1), endCoords(1)
            elem.GetStartPoint startCoords
            elem.GetEndPoint endCoords
            details = details & "            > Punct Start (x,y): (" & FormatNumber(startCoords(0), 2) & ", " & FormatNumber(startCoords(1), 2) & ")" & vbCrLf
            details = details & "            > Punct Final (x,y): (" & FormatNumber(endCoords(0), 2) & ", " & FormatNumber(endCoords(1), 2) & ")" & vbCrLf

        Case "DrawingCircle"
            elem.GetCenterPoint tempArr
            details = details & "            > Centru (x,y): (" & FormatNumber(tempArr(0), 2) & ", " & FormatNumber(tempArr(1), 2) & ")" & vbCrLf
            details = details & "            > Raza: " & FormatNumber(elem.radius, 3) & vbCrLf

        Case "DrawingArc"
            elem.GetCenterPoint tempArr
            details = details & "            > Centru (x,y): (" & FormatNumber(tempArr(0), 2) & ", " & FormatNumber(tempArr(1), 2) & ")" & vbCrLf
            details = details & "            > Raza: " & FormatNumber(elem.radius, 3) & vbCrLf
            details = details & "            > Unghi Start: " & FormatNumber(elem.StartAngle * 180 / 3.14159, 2) & " deg, Unghi Final: " & FormatNumber(elem.EndAngle * 180 / 3.14159, 2) & " deg" & vbCrLf

        Case "DrawingEllipse"
            elem.GetCenterPoint tempArr
            details = details & "            > Centru (x,y): (" & FormatNumber(tempArr(0), 2) & ", " & FormatNumber(tempArr(1), 2) & ")" & vbCrLf
            details = details & "            > Raza Mare: " & FormatNumber(elem.MajorRadius, 3) & ", Raza Mica: " & FormatNumber(elem.MinorRadius, 3) & vbCrLf

        Case "DrawingTable"
            elem.GetPosition tempArr
            details = details & "            > Randuri: " & elem.NumberOfRows & ", Coloane: " & elem.NumberOfColumns & vbCrLf
            details = details & "            > Pozitie (x,y): (" & FormatNumber(tempArr(0), 2) & ", " & FormatNumber(tempArr(1), 2) & ")" & vbCrLf
            ' ATENTIE: Decomentarea sectiunii de mai jos poate genera rapoarte FOARTE MARI
            ' Dim i As Long, j As Long
            ' For i = 1 To elem.NumberOfRows
            '     For j = 1 To elem.NumberOfColumns
            '         details = details & "            > Celula(" & i & "," & j & "): " & elem.GetCellString(i, j) & vbCrLf
            '     Next j
            ' Next i

        Case "DrawingPicture"
            elem.GetPosition tempArr
            details = details & "            > Cale Fisier: " & elem.GetPicturePath & vbCrLf
            details = details & "            > Pozitie (x,y): (" & FormatNumber(tempArr(0), 2) & ", " & FormatNumber(tempArr(1), 2) & ")" & vbCrLf
            details = details & "            > Scala: " & FormatNumber(elem.Scale, 2) & vbCrLf

        Case "DrawingGDT" ' Geometric Tolerance
            details = details & "            > Caracteristica GDT: " & elem.GetGDTCharacteristic & vbCrLf
            details = details & "            > Valoare Toleranta: " & elem.GetValue & vbCrLf
            Dim i As Long
            For i = 1 To 3 ' Max 3 datums
                If elem.GetDatum(i) <> "" Then
                    details = details & "            > Datum " & i & ": " & elem.GetDatum(i) & vbCrLf
                End If
            Next i

        Case "DrawingRoughness"
            elem.GetPosition tempArr
            details = details & "            > Pozitie (x,y): (" & FormatNumber(tempArr(0), 2) & ", " & FormatNumber(tempArr(1), 2) & ")" & vbCrLf
            details = details & "            > Tip Simbol: " & elem.SymbolType & vbCrLf
            details = details & "            > Valoare Rugozitate 1: " & elem.GetRoughnessValue(1) & ", Valoare 2: " & elem.GetRoughnessValue(2) & vbCrLf
        
        Case "DrawingWelding"
            details = details & "            > Simbol Sudura: Tipul si proprietatile specifice sunt complexe si necesita acces la obiectele asociate." & vbCrLf
        
        Case "DrawingArrow"
            Dim pts As Variant
            elem.GetPoints pts
            details = details & "            > Puncte: (" & FormatNumber(pts(0), 2) & "," & FormatNumber(pts(1), 2) & ") -> (" & FormatNumber(pts(2), 2) & "," & FormatNumber(pts(3), 2) & ")" & vbCrLf
            details = details & "            > Tip Cap: " & elem.HeadSymbol & ", Tip Coada: " & elem.TailSymbol & vbCrLf

    End Select
    
    On Error GoTo 0 ' Resetam gestionarea erorilor
    GetElementDetails = details
End Function

'================================================================================
' NOU: Functie ajutatoare pentru a extrage proprietatile grafice comune
'================================================================================
Private Function GetGraphicProperties(ByVal elem As Object) As String
    Dim graphicProps As String: graphicProps = ""
    Dim r As Long, g As Long, b As Long
    Dim lineType As Long, thickness As Long
    
    On Error Resume Next ' Proprietatile grafice pot sa nu existe pentru toate obiectele
    
    ' Extrage culoare
    elem.GetGraphicAttribute r, g, b
    If Err.Number = 0 Then
        graphicProps = graphicProps & "            > Culoare (RGB): " & r & "," & g & "," & b
    End If
    
    ' Extrage tipul de linie
    elem.GetGraphicAttribute lineType
    If Err.Number = 0 Then
        If graphicProps <> "" Then graphicProps = graphicProps & " | "
        graphicProps = graphicProps & "Tip Linie: " & lineType
    End If
    
    ' Extrage grosimea liniei
    elem.GetGraphicAttribute thickness
    If Err.Number = 0 Then
        If graphicProps <> "" Then graphicProps = graphicProps & " | "
        graphicProps = graphicProps & "Grosime: " & thickness
    End If
    
    If graphicProps <> "" Then
        graphicProps = graphicProps & vbCrLf
    End If
    
    On Error GoTo 0
    GetGraphicProperties = graphicProps
End Function
