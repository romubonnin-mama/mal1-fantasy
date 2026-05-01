Option Explicit

' InsertLogos - Insere les logos des clubs dans toutes les feuilles
' Prerequis : lancer d'abord python scripts/download_logos.py
' Utilisation : Alt+F11 -> Insertion -> Module -> coller -> F5

Sub InsertLogos()
    Dim logosPath As String
    logosPath = "E:\Sauvegarde\draft club\mal1-fantasy\logos\"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim success As Boolean
    success = False
    On Error GoTo Cleanup

    ' Dictionnaire joueur -> club ID
    Dim cdict As Object
    Set cdict = CreateObject("Scripting.Dictionary")
    cdict.CompareMode = vbTextCompare

    cdict("RESTES") = "96"
    cdict("CLAUSS") = "84"
    cdict("SIDIBE") = "96"
    cdict("TAGLIAFICO") = "80"
    cdict("M.SARR") = "116"
    cdict("LEPENANT") = "83"
    cdict("NEVES") = "85"
    cdict("BOUFAL") = "111"
    cdict("DOUMBIA") = "106"
    cdict("MORTON") = "80"
    cdict("DIOP") = "84"
    cdict("MORREIRA") = "95"
    cdict("DEL CASTILLO") = "106"
    cdict("GOUIRI") = "81"
    cdict("RISSER") = "116"
    cdict("AGUERD") = "81"
    cdict("PERRAUD") = "79"
    cdict("NIAKHATE") = "80"
    cdict("UDOL") = "116"
    cdict("SENAYA") = "108"
    cdict("GOLOVIN") = "91"
    cdict("PAPE DEMBA") = "96"
    cdict("SZYMANSKI") = "94"
    cdict("RUIZ") = "85"
    cdict("TSITAISHVILI") = "112"
    cdict("DEMBELE") = "85"
    cdict("EMBOLO") = "94"
    cdict("PANICHELLI") = "95"
    cdict("LOPES") = "83"
    cdict("G.DOUE") = "95"
    cdict("LALA") = "106"
    cdict("PAVARD") = "81"
    cdict("PACHO") = "85"
    cdict("DANOIS") = "108"
    cdict("AKLIOUCHE") = "91"
    cdict("MAGNETTI") = "106"
    cdict("SANSON") = "84"
    cdict("BOUADDI") = "79"
    cdict("ENCISO") = "95"
    cdict("GODO") = "95"
    cdict("PAGIS") = "97"
    cdict("SAID") = "116"
    cdict("RULLI") = "81"
    cdict("O.CAMARA") = "77"
    cdict("NEGO") = "111"
    cdict("MATA") = "80"
    cdict("BARD") = "84"
    cdict("MBOW") = "114"
    cdict("ZAKARIA") = "91"
    cdict("CHOTARD") = "106"
    cdict("CASSERES") = "96"
    cdict("BELKEBLA") = "77"
    cdict("PERRIN") = "79"
    cdict("THAUVIN") = "116"
    cdict("ENDRICK") = "80"
    cdict("GBOHO") = "96"
    cdict("OZER") = "79"
    cdict("KOUASSI") = "97"
    cdict("NICOLAISEN") = "96"
    cdict("CHARDONNET") = "106"
    cdict("MERLIN") = "94"
    cdict("HAKIMI") = "85"
    cdict("MUNETSI") = "114"
    cdict("HEIN") = "112"
    cdict("SANGARE") = "116"
    cdict("OWUSU") = "108"
    cdict("ANDRE") = "79"
    cdict("H.DIALLO") = "112"
    cdict("GREENWOOD") = "81"
    cdict("AJORQUE") = "106"
    cdict("GREIF") = "80"
    cdict("KLUIVERT") = "80"
    cdict("NGOY") = "79"
    cdict("GANIOU") = "116"
    cdict("ZABARNYI") = "85"
    cdict("LOCKO") = "106"
    cdict("AL-TAMARI") = "94"
    cdict("THOMASSON") = "116"
    cdict("VITINHA") = "85"
    cdict("KEBBAL") = "114"
    cdict("SULC") = "80"
    cdict("LEPAUL") = "94"
    cdict("SAINT MAXIMAIN") = "116"
    cdict("WAHI") = "84"
    cdict("PENDERS") = "95"
    cdict("N.MENDES") = "85"
    cdict("MAITLAND-NILES") = "80"
    cdict("TEZE") = "91"
    cdict("MCKENZIE") = "96"
    cdict("DOUCOURE") = "95"
    cdict("HARALDSSON") = "79"
    cdict("ZAIRE-EMERY") = "85"
    cdict("CABELLA") = "83"
    cdict("RONGIER") = "94"
    cdict("NWANERY") = "81"
    cdict("DOUE") = "85"
    cdict("EDOUARD") = "116"
    cdict("GIROUD") = "79"
    cdict("SAMBA") = "94"
    cdict("MANDI") = "79"
    cdict("BARCO") = "95"
    cdict("WEAH") = "81"
    cdict("LEFORT") = "77"
    cdict("BRASSIER") = "94"
    cdict("ABERGEL") = "97"
    cdict("HOJBJERG") = "81"
    cdict("M.CAMARA") = "94"
    cdict("LEES-MELOU") = "114"
    cdict("VAN DEN BOOMEN") = "77"
    cdict("F.PARDO") = "79"
    cdict("KVARATSKHELIA") = "85"
    cdict("ABLINE") = "83"
    cdict("DIAW") = "111"
    cdict("AGUILAR") = "116"
    cdict("CRESSWELL") = "96"
    cdict("LLORIS") = "111"
    cdict("ARCUS") = "77"
    cdict("VANDERSON") = "91"
    cdict("NDIAYE") = "111"
    cdict("TOLISSO") = "80"
    cdict("L.CAMARA") = "91"
    cdict("BELKHDIM") = "77"
    cdict("DONNUM") = "96"
    cdict("SOUMARE") = "111"
    cdict("BALOGUN") = "91"
    cdict("SINAYOKO") = "108"

    ' Structure des feuilles
    Dim groupHeaders(2) As Integer
    groupHeaders(0) = 2
    groupHeaders(1) = 42
    groupHeaders(2) = 82

    ' Variables de boucle
    Dim ws As Worksheet
    Dim shNum As Integer
    Dim blockWidth As Integer
    Dim shapeNames() As String
    Dim nShapes As Integer
    Dim k As Integer
    Dim shp As Shape
    Dim done As Object
    Dim colPos As Integer
    Dim logoCol As Integer
    Dim nameCol As Integer
    Dim grp As Integer

    For Each ws In ThisWorkbook.Worksheets
        If IsNumeric(ws.Name) Then
            shNum = CInt(ws.Name)
            If shNum >= 14 Then
                blockWidth = 19
            Else
                blockWidth = 18
            End If

            ' Supprimer les anciens logos (type 13 = msoPicture)
            nShapes = 0
            ReDim shapeNames(0)
            For Each shp In ws.Shapes
                If shp.Type = 13 Then
                    nShapes = nShapes + 1
                    ReDim Preserve shapeNames(nShapes - 1)
                    shapeNames(nShapes - 1) = shp.Name
                End If
            Next shp
            For k = 0 To nShapes - 1
                On Error Resume Next
                ws.Shapes(shapeNames(k)).Delete
                On Error GoTo Cleanup
            Next k

            ' Un logo par club cree en AddPicture, les suivants en Duplicate
            Set done = CreateObject("Scripting.Dictionary")

            For colPos = 0 To 2
                logoCol = 2 + colPos * blockWidth
                nameCol = logoCol + 1
                For grp = 0 To 2
                    Call DoOffsets(ws, groupHeaders(grp), Array(2), logoCol, nameCol, logosPath, cdict, done)
                    Call DoOffsets(ws, groupHeaders(grp), Array(5, 7, 9, 11, 13), logoCol, nameCol, logosPath, cdict, done)
                    Call DoOffsets(ws, groupHeaders(grp), Array(18, 20, 22, 24, 26, 28), logoCol, nameCol, logosPath, cdict, done)
                    Call DoOffsets(ws, groupHeaders(grp), Array(31, 33, 35), logoCol, nameCol, logosPath, cdict, done)
                Next grp
            Next colPos
        End If
    Next ws

    success = True

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    If Err.Number <> 0 Then
        MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical
    ElseIf success Then
        MsgBox "Logos inseres avec succes !", vbInformation
    End If
End Sub


Private Function NormName(s As String) As String
    Dim r As String
    r = UCase(Trim(s))
    r = Replace(r, Chr(214), "O")
    r = Replace(r, Chr(246), "O")
    NormName = r
End Function


Private Sub DoOffsets(ws As Worksheet, headerRow As Integer, offsets As Variant, _
                      logoCol As Integer, nameCol As Integer, logosPath As String, _
                      cdict As Object, done As Object)
    Dim i As Integer
    Dim pRow As Integer
    Dim cellVal As Variant
    Dim pName As String
    Dim clubId As String
    Dim lCell As Range
    Dim logoFile As String
    Dim pic As Shape
    Dim srcShp As Shape
    Dim dup As Shape

    For i = 0 To UBound(offsets)
        pRow = headerRow + offsets(i)
        cellVal = ws.Cells(pRow, nameCol).Value

        If Not IsEmpty(cellVal) And Not IsNull(cellVal) Then
            pName = NormName(CStr(cellVal))
            If pName <> "" Then
                If cdict.Exists(pName) Then
                    clubId = cdict(pName)
                    Set lCell = ws.Cells(pRow, logoCol)

                    If done.Exists(clubId) Then
                        Set srcShp = Nothing
                        On Error Resume Next
                        Set srcShp = ws.Shapes(CStr(done(clubId)))
                        On Error GoTo 0
                        If Not srcShp Is Nothing Then
                            Set dup = srcShp.Duplicate()
                            dup.Left = lCell.Left + 1
                            dup.Top = lCell.Top + 1
                            dup.Width = lCell.Width - 2
                            dup.Height = lCell.Height - 2
                            dup.Placement = 1
                        End If
                    Else
                        logoFile = logosPath & clubId & ".png"
                        If Dir(logoFile) <> "" Then
                            Set pic = Nothing
                            On Error Resume Next
                            Set pic = ws.Shapes.AddPicture(logoFile, False, True, _
                                lCell.Left + 1, lCell.Top + 1, _
                                lCell.Width - 2, lCell.Height - 2)
                            On Error GoTo 0
                            If Not pic Is Nothing Then
                                pic.Placement = 1
                                done.Add clubId, pic.Name
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Sub
