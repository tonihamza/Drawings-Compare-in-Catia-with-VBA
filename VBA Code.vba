' =================================================================================
' MACRO DE EXTRAGERE DATE DIN DESENE CATIA (VERSIUNEA FINALA - CU SIMBOL LEADER)
' Metoda: Extragere complet automatizata, cu focus pe extragerea tipului de
'         simbol al leader-ului, in loc de coordonate.
' =================================================================================

Option Explicit

' Declaratie pentru functia Sleep
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


Sub ExtractAllDrawingData_WithSymbols()
    Dim AUDIT_FILE_PATH As String
    AUDIT_FILE_PATH = Environ("USERPROFILE") & "\Desktop\CATIA_Data_Extraction_Report.txt"

    Dim CATIA As Object
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application")
    If CATIA Is Nothing Then MsgBox "Nu s-a putut conecta la CATIA.", vbCritical: Exit Sub
    On Error GoTo 0
    
    CATIA.Visible = True
    CATIA.StatusBar = "Pornire extragere date (mod simboluri)..."
    CATIA.DisplayFileAlerts = False

    Dim fso As Object, reportFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set reportFile = fso.CreateTextFile(AUDIT_FILE_PATH, True, True)
    
    reportFile.WriteLine "======= RAPORT DE EXTRAGERE DATE - DESENE CATIA ======="
    reportFile.WriteLine "Generat la: " & Now & " (Mod de extragere: Simboluri Leader)"
    reportFile.WriteLine "======================================================================================" & vbNewLine

    Dim doc As Object, drawingCount As Integer
    drawingCount = 0
    For Each doc In CATIA.Documents
        If TypeName(doc) = "DrawingDocument" Then
            drawingCount = drawingCount + 1
            reportFile.WriteLine "--- DESEN " & drawingCount & ": " & doc.Name & " ---"
            
            On Error Resume Next
            CATIA.Windows.Item(doc.Name).Activate
            doc.Activate
            CATIA.RefreshDisplay = True
            Err.Clear
            On Error GoTo 0

            Dim aSheet As Object
            For Each aSheet In doc.Sheets
                If Left(UCase(aSheet.Name), 5) = "SHEET" Then
                    On Error Resume Next
                    aSheet.Activate
                    CATIA.RefreshDisplay = True
                    reportFile.WriteLine "  -- Foaie: '" & aSheet.Name & "' (Format: " & aSheet.PaperSize & ", Scala: " & aSheet.Scale & ")"
                    Err.Clear
                    On Error GoTo 0
                    
                    Dim aView As Object
                    For Each aView In aSheet.Views
                        reportFile.WriteLine "      - Vedere: '" & aView.Name & "' (Scala: " & aView.Scale & ")"
                        
                        On Error Resume Next
                        aView.Activate
                        aView.Update
                        CATIA.RefreshDisplay = True
                        Sleep 50 ' Pauza scurta
                        On Error GoTo 0

                        ListElementsFromView aView, reportFile
                    Next aView
                    reportFile.WriteLine "  -------------------------------------"
                End If
            Next aSheet
            reportFile.WriteLine ""
        End If
    Next doc
    
    reportFile.Close
    CATIA.DisplayFileAlerts = True
    CATIA.StatusBar = "Extragere date finalizata."
    MsgBox "Raportul final a fost salvat pe Desktop:" & vbNewLine & "CATIA_Data_Extraction_Report.txt", vbInformation, "Proces Finalizat"
End Sub

Private Sub ListElementsFromView(ByVal targetView As Object, ByRef outFile As Object)
    If targetView Is Nothing Then Exit Sub
    
    Dim allCollections As Variant
    allCollections = Array("GeometricElements", "Dimensions", "Texts", "Tables")
    
    Dim collectionName As Variant, currentCollection As Object
    For Each collectionName In allCollections
        On Error Resume Next
        Set currentCollection = CallByName(targetView, collectionName, VbGet)
        
        If Err.Number = 0 And Not currentCollection Is Nothing And currentCollection.Count > 0 Then
            outFile.WriteLine "        -> Gasit " & currentCollection.Count & " element(e) in ." & collectionName & ":"
            Dim elem As Object
            For Each elem In currentCollection
                Dim elemName As String
                On Error Resume Next
                elemName = elem.Name
                If Err.Number <> 0 Then elemName = "[fara nume]": Err.Clear
                On Error GoTo 0
                
                outFile.WriteLine "          - Tip: " & TypeName(elem) & ", Nume: " & elemName
                outFile.Write GetElementDetails(elem)
            Next elem
        End If
        Err.Clear
        Set currentCollection = Nothing
    Next collectionName
End Sub

Private Function GetElementDetails(ByVal elem As Object) As String
    Dim details As String: details = ""
    Dim tempArr(1)
    
    On Error Resume Next
    
    Select Case TypeName(elem)
        Case "DrawingDimension"
            Dim dimValue As Object, upTol As Double, lowTol As Double
            Set dimValue = elem.GetValue
            If Err.Number = 0 Then
                details = details & "            > Valoare Masurata: " & FormatNumber(dimValue.Value, 3) & vbCrLf
                details = details & "            > Text Inainte/Dupa: """ & dimValue.GetBaultText(1) & """ / """ & dimValue.GetBaultText(2) & """" & vbCrLf
                dimValue.GetTolerance upTol, lowTol
                If Err.Number = 0 Then details = details & "            > Toleranta: Sup: " & FormatNumber(upTol, 3) & ", Inf: " & FormatNumber(lowTol, 3) & vbCrLf
                Err.Clear
            End If
            details = details & "            > Tip Cota: " & elem.DimType & vbCrLf

        Case "DrawingText", "DrawingTextWithLeader"
            Dim xPos As Double, yPos As Double
            Dim numLeaders As Integer
            
            details = details & "            > Text: """ & elem.Text & """" & vbCrLf
            xPos = elem.X: yPos = elem.Y
            If Err.Number = 0 Then details = details & "            > Pozitie Text (x,y): (" & FormatNumber(xPos, 2) & ", " & FormatNumber(yPos, 2) & ")" & vbCrLf
            Err.Clear
            
            numLeaders = 0
            On Error Resume Next
            numLeaders = elem.Leaders.Count
            If Err.Number <> 0 Then numLeaders = 0: Err.Clear
            details = details & "            > Numar de leaderi: " & numLeaders & vbCrLf

            If numLeaders > 0 Then
                Dim i As Integer, aLeader As Object
                For i = 1 To numLeaders
                    Set aLeader = elem.Leaders.Item(i)
                    
                    ' <<< MODIFICARE: Extragem tipul de simbol al leader-ului >>>
                    Dim headSymbolType As Integer
                    Dim headSymbolName As String
                    
                    On Error Resume Next
                    headSymbolType = aLeader.HeadSymbol
                    If Err.Number <> 0 Then
                        headSymbolName = "[Simbol Indisponibil]"
                        Err.Clear
                    Else
                        headSymbolName = GetLeaderSymbolName(headSymbolType)
                    End If
                    On Error GoTo 0
                    
                    details = details & "              > Leader #" & i & " - Simbol Vârf: " & headSymbolName & vbCrLf
                Next i
            End If
            
        Case "Point2D"
            elem.GetCoordinates tempArr: If Err.Number = 0 Then details = details & "            > Coordonate (x,y): (" & FormatNumber(tempArr(0), 2) & ", " & FormatNumber(tempArr(1), 2) & ")" & vbCrLf
            Err.Clear
            
        Case "Line2D"
            Dim startPoint As Object, endPoint As Object
            Set startPoint = elem.startPoint: Set endPoint = elem.endPoint
            startPoint.GetCoordinates tempArr: If Err.Number = 0 Then details = details & "            > Punct Start (x,y): (" & FormatNumber(tempArr(0), 2) & ", " & FormatNumber(tempArr(1), 2) & ")" & vbCrLf
            Err.Clear
            endPoint.GetCoordinates tempArr: If Err.Number = 0 Then details = details & "            > Punct Final (x,y): (" & FormatNumber(tempArr(0), 2) & ", " & FormatNumber(tempArr(1), 2) & ")" & vbCrLf
            Err.Clear

        Case "Circle2D"
            Dim centerPoint As Object
            Set centerPoint = elem.centerPoint
            centerPoint.GetCoordinates tempArr: If Err.Number = 0 Then details = details & "            > Centru (x,y): (" & FormatNumber(tempArr(0), 2) & ", " & FormatNumber(tempArr(1), 2) & ")" & vbCrLf
            Err.Clear
            details = details & "            > Raza: " & FormatNumber(elem.Radius, 3) & vbCrLf

        Case "Axis2D"
            Dim originPoint As Object
            Set originPoint = elem.originPoint
            originPoint.GetCoordinates tempArr: If Err.Number = 0 Then details = details & "            > Origine (x,y): (" & FormatNumber(tempArr(0), 2) & ", " & FormatNumber(tempArr(1), 2) & ")" & vbCrLf
            Err.Clear
            
        Case "DrawingTable"
            details = details & "            > Randuri: " & elem.NumberOfRows & ", Coloane: " & elem.NumberOfColumns & vbCrLf
    End Select
    
    On Error GoTo 0
    GetElementDetails = details
End Function

Private Function GetLeaderSymbolName(ByVal symbolEnum As Integer) As String
    ' Traduce valoarea numerica a simbolului intr-un text descriptiv
    ' Lista a fost extinsa pe baza rezultatelor din raport.
    On Error Resume Next
    Select Case symbolEnum
        Case 1: GetLeaderSymbolName = "Sageata (catArrow)"
        Case 2: GetLeaderSymbolName = "Sageata simpla (catSimpleArrow)"
        Case 3: GetLeaderSymbolName = "Triunghi plin (catTriangle)"
        Case 4: GetLeaderSymbolName = "Cerc gol (catCircle)"
        Case 5: GetLeaderSymbolName = "Cerc plin (catFilledCircle)"
        Case 6: GetLeaderSymbolName = "Cruce (catCross)"
        Case 7: GetLeaderSymbolName = "Fara simbol (catNoSymbol)"
        
        ' <<< MODIFICARE: Adaugam noile simboluri descoperite >>>
        Case 17: GetLeaderSymbolName = "Datum Target / Sageata goala (catDatumTarget / catOpenArrow)"
        Case 20: GetLeaderSymbolName = "Sageata plina (catArrowFilled)" ' Adesea folosit pentru baloane
        Case 25: GetLeaderSymbolName = "Datum Feature (catDatumFeature)"
        
        Case Else: GetLeaderSymbolName = "Necunoscut (" & symbolEnum & ")"
    End Select
End Function
