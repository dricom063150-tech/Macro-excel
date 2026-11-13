Attribute VB_Name = "Foisonnement"
Option Explicit

' =====================================================================
'  VISUALISATION DU FOISONNEMENT - Matrice Amont -> Aval -> Preuves
' ---------------------------------------------------------------------
'  Lancer "VisualiserFoisonnement" depuis la feuille source.
'  La feuille source n'est PAS modifiée. La macro crée la feuille
'  "Revue foisonnement" : copie mise en forme pour la revue
'        - cellules amont / aval répétées fusionnées (ou grisées)
'        - une bande de couleur par exigence amont, trait épais entre amonts
'        - trait fin entre exigences avales d'un même amont
'        - ROUGE  : exigence amont sans exigence aval
'        - ORANGE : exigence aval sans preuve
'  Relancer la macro régénère la feuille.
' =====================================================================

' ---------------- PARAMÈTRES À ADAPTER ----------------
' Numéros de colonnes : A=1, B=2, C=3, D=4...
' Si un bloc fait plusieurs colonnes (ex. ID + libellé), mettre
' l'identifiant en 1ère colonne du bloc : c'est lui qui sert à regrouper.
Private Const LIGNE_ENTETE As Long = 1
Private Const COL_AMONT_DEB As Long = 1
Private Const COL_AMONT_FIN As Long = 2
Private Const COL_AVAL_DEB As Long = 3
Private Const COL_AVAL_FIN As Long = 4
Private Const COL_PREUVE_DEB As Long = 5
Private Const COL_PREUVE_FIN As Long = 5

' True = fusionne les cellules répétées ; False = grise les répétitions
' (False garde le filtrage Excel utilisable sur la feuille de revue)
Private Const MODE_FUSION As Boolean = True

' True = trie par amont puis aval avant traitement (utile si les lignes
' d'une même exigence ne sont pas contiguës). False = ordre d'origine.
Private Const TRIER As Boolean = False

Private Const NOM_REVUE As String = "Revue foisonnement"
' -------------------------------------------------------


Public Sub VisualiserFoisonnement()
    Dim wb As Workbook
    Dim wsSrc As Worksheet, ws As Worksheet
    Dim lo As ListObject
    Dim derLig As Long, derCol As Long
    Dim i As Long, debut As Long, nbAmont As Long
    Dim bande As Boolean

    Set wsSrc = ActiveSheet
    Set wb = wsSrc.Parent

    If wsSrc.Name = NOM_REVUE Then
        MsgBox "Lancez la macro depuis la feuille source, pas depuis la feuille de revue.", vbExclamation
        Exit Sub
    End If

    derLig = DerniereLigne(wsSrc)
    If derLig <= LIGNE_ENTETE Then
        MsgBox "Aucune donnee trouvee sous la ligne " & LIGNE_ENTETE & ".", vbExclamation
        Exit Sub
    End If

    On Error GoTo Erreur
    Application.ScreenUpdating = False

    SupprimerFeuille wb, NOM_REVUE

    ' --- Copie de travail ---
    wsSrc.Copy After:=wsSrc
    Set ws = ActiveSheet
    ws.Name = NOM_REVUE
    For Each lo In ws.ListObjects          ' un tableau Excel empeche la fusion
        lo.Unlist
    Next lo
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.UnMerge
    ws.Cells.FormatConditions.Delete

    derCol = ws.Cells(LIGNE_ENTETE, ws.Columns.Count).End(xlToLeft).Column
    If derCol < COL_PREUVE_FIN Then derCol = COL_PREUVE_FIN

    If TRIER Then TrierDonnees ws, derLig, derCol

    ' --- Mise en forme generale ---
    StyleEntete ws.Range(ws.Cells(LIGNE_ENTETE, 1), ws.Cells(LIGNE_ENTETE, derCol))
    With ws.Range(ws.Cells(LIGNE_ENTETE + 1, 1), ws.Cells(derLig, derCol))
        .Interior.Pattern = xlNone
        .Font.Color = vbBlack
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(217, 217, 217)
    End With

    ' --- Parcours des groupes amont ---
    bande = False
    i = LIGNE_ENTETE + 1
    Do While i <= derLig
        debut = i
        Do While i < derLig
            If Cle(ws, i + 1, COL_AMONT_DEB) <> Cle(ws, debut, COL_AMONT_DEB) Then Exit Do
            i = i + 1
        Loop
        TraiterAmont ws, debut, i, derCol, bande
        bande = Not bande
        nbAmont = nbAmont + 1
        i = i + 1
    Loop

    ' --- Finitions feuille de revue ---
    ws.Rows(LIGNE_ENTETE + 1 & ":" & derLig).AutoFit
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(LIGNE_ENTETE + 1, 1).Select
    ActiveWindow.FreezePanes = True

    Application.ScreenUpdating = True
    MsgBox nbAmont & " exigences amont traitees." & vbCrLf & _
           "Voir la feuille '" & NOM_REVUE & "'." & vbCrLf & _
           "Rouge : amont sans aval / Orange : aval sans preuve", vbInformation
    Exit Sub

Erreur:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub


' Traite un bloc de lignes d..f partageant la meme exigence amont
Private Sub TraiterAmont(ws As Worksheet, d As Long, f As Long, _
                         derCol As Long, bande As Boolean)
    Dim j As Long, r As Long, c As Long, dAval As Long
    Dim rng As Range

    ' Bande de couleur + traits epais autour du groupe
    Set rng = ws.Range(ws.Cells(d, 1), ws.Cells(f, derCol))
    If bande Then
        rng.Interior.Color = RGB(221, 235, 247)
    Else
        rng.Interior.Color = RGB(255, 255, 255)
    End If
    Bordure rng.Borders(xlEdgeTop), xlMedium, RGB(31, 78, 121)
    Bordure rng.Borders(xlEdgeBottom), xlMedium, RGB(31, 78, 121)

    ' Sous-groupes avales
    j = d
    Do While j <= f
        dAval = j
        Do While j < f
            If Cle(ws, j + 1, COL_AVAL_DEB) <> Cle(ws, dAval, COL_AVAL_DEB) Then Exit Do
            j = j + 1
        Loop

        If Cle(ws, dAval, COL_AVAL_DEB) = "" Then
            ' Amont non decline
            ws.Range(ws.Cells(dAval, COL_AVAL_DEB), ws.Cells(j, COL_AVAL_FIN)).Interior.Color = RGB(255, 199, 206)
        Else
            ' Aval sans preuve
            For r = dAval To j
                If Cle(ws, r, COL_PREUVE_DEB) = "" Then
                    ws.Range(ws.Cells(r, COL_PREUVE_DEB), ws.Cells(r, COL_PREUVE_FIN)).Interior.Color = RGB(255, 235, 156)
                End If
            Next r
        End If

        ' Trait fin entre deux avales d'un meme amont
        If dAval > d Then
            Bordure ws.Range(ws.Cells(dAval, COL_AVAL_DEB), ws.Cells(j, derCol)).Borders(xlEdgeTop), _
                    xlThin, RGB(89, 89, 89)
        End If

        For c = COL_AVAL_DEB To COL_AVAL_FIN
            Fusionner ws, dAval, j, c
        Next c
        j = j + 1
    Loop

    ' Colonnes amont
    ws.Range(ws.Cells(d, COL_AMONT_DEB), ws.Cells(d, COL_AMONT_FIN)).Font.Bold = True
    For c = COL_AMONT_DEB To COL_AMONT_FIN
        Fusionner ws, d, f, c
    Next c

End Sub



' ---------------- Utilitaires ----------------

Private Sub Fusionner(ws As Worksheet, d As Long, f As Long, c As Long)
    If f <= d Then Exit Sub
    If MODE_FUSION Then
        ws.Range(ws.Cells(d + 1, c), ws.Cells(f, c)).ClearContents
        ws.Range(ws.Cells(d, c), ws.Cells(f, c)).Merge
    Else
        ws.Range(ws.Cells(d + 1, c), ws.Cells(f, c)).Font.Color = RGB(191, 191, 191)
    End If
End Sub

Private Sub Bordure(b As Border, poids As XlBorderWeight, couleur As Long)
    b.LineStyle = xlContinuous
    b.Weight = poids
    b.Color = couleur
End Sub

Private Sub StyleEntete(rng As Range)
    With rng
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 78, 121)
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
End Sub

Private Function Cle(ws As Worksheet, r As Long, c As Long) As String
    Dim v As Variant
    v = ws.Cells(r, c).Value
    If IsError(v) Then
        Cle = "#ERREUR"
    Else
        Cle = Trim$(CStr(v))
    End If
End Function

Private Function DerniereLigne(ws As Worksheet) As Long
    Dim m As Long, x As Long
    m = ws.Cells(ws.Rows.Count, COL_AMONT_DEB).End(xlUp).Row
    x = ws.Cells(ws.Rows.Count, COL_AVAL_DEB).End(xlUp).Row: If x > m Then m = x
    x = ws.Cells(ws.Rows.Count, COL_PREUVE_DEB).End(xlUp).Row: If x > m Then m = x
    DerniereLigne = m
End Function

Private Sub TrierDonnees(ws As Worksheet, derLig As Long, derCol As Long)
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(LIGNE_ENTETE + 1, COL_AMONT_DEB), ws.Cells(derLig, COL_AMONT_DEB)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=ws.Range(ws.Cells(LIGNE_ENTETE + 1, COL_AVAL_DEB), ws.Cells(derLig, COL_AVAL_DEB)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange ws.Range(ws.Cells(LIGNE_ENTETE, 1), ws.Cells(derLig, derCol))
        .Header = xlYes
        .Apply
    End With
End Sub

Private Sub SupprimerFeuille(wb As Workbook, nom As String)
    Dim sh As Object
    On Error Resume Next
    Set sh = wb.Sheets(nom)
    On Error GoTo 0
    If Not sh Is Nothing Then
        Application.DisplayAlerts = False
        sh.Delete
        Application.DisplayAlerts = True
    End If
End Sub
