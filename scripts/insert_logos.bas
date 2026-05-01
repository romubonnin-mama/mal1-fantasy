Attribute VB_Name = "ModuleLogos"
Option Explicit

Sub InsertLogos()

    Dim logosPath As String
    logosPath = "E:\Sauvegarde\draft club\mal1-fantasy\logos\"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error Resume Next

    ' --- Dictionnaire joueur -> club ID ---
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

    ' --- Lignes d'en-tete des 3 groupes ---
    Dim groupHdr(2) As Integer
    groupHdr(0) = 2
    groupHdr(1) = 42
    groupHdr(2) = 82

    ' --- 15 offsets de lignes joueurs (G/D/M/A) ---
    Dim posOff(14) As Integer
    posOff(0) = 2
    posOff(1) = 5
    posOff(2) = 7
    posOff(3) = 9
    posOff(4) = 11
    posOff(5) = 13
    posOff(6) = 18
    posOff(7) = 20
    posOff(8) = 22
    posOff(9) = 24
    posOff(10) = 26
    posOff(11) = 28
    posOff(12) = 31
    posOff(13) = 33
    posOff(14) = 35

    ' --- Variables de boucle ---
    Dim ws As Worksheet
    Dim shNum As Integer
    Dim blockWidth As Integer
    Dim shapeNames() As String
    Dim nShapes As Integer
    Dim k As Integer
    Dim shp As Shape
    Dim done As Object
    Dim colPos As Integer
    Dim baseCol As Integer
    Dim logoCol As Integer
    Dim nameCol As Integer
    Dim grp As Integer
    Dim pidx As Integer
    Dim pRow As Integer
    Dim cellVal As Variant
    Dim pName As String
    Dim clubId As String
    Dim lCell As Range
    Dim logoFile As String
    Dim pic As Shape
    Dim marg As Double

    ' --- Parcourir toutes les feuilles numerotees ---
    For Each ws In ThisWorkbook.Worksheets
        If IsNumeric(ws.Name) Then
            shNum = CInt(ws.Name)
            If shNum >= 14 Then
                blockWidth = 19
            Else
                blockWidth = 18
            End If

            ' Supprimer les anciens logos (msoPicture = 13)
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
                ws.Shapes(shapeNames(k)).Delete
            Next k

            ' Cache : clubId -> chemin du fichier logo (evite Dir() repetes)
            Set done = CreateObject("Scripting.Dictionary")

            marg = 3

            For colPos = 0 To 2
                baseCol = 2 + colPos * blockWidth
                nameCol = baseCol + 1
                logoCol = baseCol + 2

                For grp = 0 To 2
                    For pidx = 0 To 14
                        pRow = groupHdr(grp) + posOff(pidx)
                        cellVal = ws.Cells(pRow, nameCol).Value

                        If Not IsEmpty(cellVal) Then
                            If Not IsNull(cellVal) Then
                                pName = UCase(Trim(CStr(cellVal)))
                                pName = Replace(pName, Chr(214), "O")

                                If Len(pName) > 0 Then
                                    If cdict.Exists(pName) Then
                                        clubId = CStr(cdict(pName))

                                        If Not done.Exists(clubId) Then
                                            logoFile = logosPath & clubId & ".png"
                                            If Dir(logoFile) <> "" Then
                                                done.Add clubId, logoFile
                                            Else
                                                done.Add clubId, ""
                                            End If
                                        End If

                                        logoFile = CStr(done(clubId))
                                        If logoFile <> "" Then
                                            Set lCell = ws.Cells(pRow, logoCol).MergeArea
                                            Set pic = Nothing
                                            Set pic = ws.Shapes.AddPicture( _
                                                logoFile, False, True, _
                                                lCell.Left + marg, _
                                                lCell.Top + marg, _
                                                lCell.Width - marg * 2, _
                                                lCell.Height - marg * 2)
                                            If Not pic Is Nothing Then
                                                pic.Placement = 1
                                            End If
                                        End If

                                    End If
                                End If
                            End If
                        End If

                    Next pidx
                Next grp
            Next colPos
        End If
    Next ws

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "Logos inseres avec succes !", vbInformation

End Sub
