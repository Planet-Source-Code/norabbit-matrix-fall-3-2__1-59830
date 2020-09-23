VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTitre2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE LE DEUXIEME TITRE (GRAND)                                    *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'titres
Public TitreSz1 As String
Public TitreSz2 As String
'
'totaux des lettres à afficher par titre
Dim Tot1 As Integer
Dim Tot2 As Integer
'
'nombre de caractères à afficher
Dim NbrCar As Integer
'
'nombre de caractères dont la durée de vie s'est écoulée
Dim NbrCarFin As Integer
'
'coordonnées du titre #1
Public XSz1 As Single
Public YSz1 As Single
'
'coordonnées du titre #2
Public XSz2 As Single
Public YSz2 As Single
'
'couleur du titre
'Public RSz As Long
'Public VSz As Long
'Public BSz As Long
'
'état du titre
Public Etat As Integer
'
'classe gérant une petite lettre (titre #1)
Dim cLnP() As ClsLettrePetite
'
'classe gérant une lettre normale (titre #2)
Dim cLn() As ClsLettreNormal
'
'classe gérant l'effet #1 et #2
Dim cFX1 As New ClsEffet1
'
'classe se chargeant d'afficher une ligne de caractères spéciaux
Dim cLT1 As New ClsLigneTitre1
'
'classe gérant un effet (#3)
Dim cFX3 As New ClsEffet3
'
'liste des lettres pour le titre #1
Dim LstLT1 As New Collection
'
Dim LettreTmp As Long
Dim i As Integer
Dim iTmp As Integer
'
'initialisation des valeurs
Public Sub InitTitre2()
    '
    'on vérifie s'il y a bien un texte à afficher
    If TxtTitre1.Count = 0 Then
        '
        Etat = 0
        '
        'on quitte ici
        Exit Sub
        '
    End If
    '
    'on récupère les textes à afficher
    TitreSz1 = LCase$(TxtTitre1.Item(1))
    TitreSz2 = TxtTitre2.Item(1)
    '
    dMsg2 = TitreSz1 & " : " & TitreSz2
    '
    'on efface les textes des listes
    TxtTitre1.Remove 1
    TxtTitre2.Remove 1
    '
    'on calcule la position du titre #1 en fonction de la taille des lettres et de l'espacement
    XSz1 = (DimL / 2) - (Len(TitreSz1) * ((LargeurLettreSz / 18) * 7)) - ((LargeurLettreSz / 18) * 14)
    YSz1 = (DimH / 2) - (HauteurLettreSz / 2) - HauteurLettreSz - 10
    '
    'on calcule la position du titre #2 en fonction de la taille des lettres et de l'espacement
    XSz2 = (DimL / 2) - (Len(TitreSz2) * ((LargeurLettreSz / 18) * 9.5)) - ((LargeurLettreSz / 18) * 19)
    YSz2 = (DimH / 2) - (HauteurLettreSz / 2)
    '
    'on initialise certaines variables
    NbrCar = 0
    NbrCarFin = 0
    Etat = 1
    Tot1 = 0
    Tot2 = 0
    '
    'on écrase toutes les classes créées
    Erase cLnP()
    Erase cLn()
    '
    'on ajoute toutes les lettres du titre à la liste
    For i = 1 To Len(TitreSz1)
        '
        If Mid(TitreSz1, i, 1) <> " " Then LstLT1.Add i
        '
    Next
    '
    'on crée le nombre de classes nécessaires pour afficher toutes les lettres du titre #1
    For i = 1 To Len(TitreSz1)
        '
        'on ne crée rien si c'est un espace
        If Mid(TitreSz1, i, 1) <> " " Then
            '
            'on incrémente le nombre de caractères à afficher
            NbrCar = NbrCar + 1
            '
            'on incrémente le nombre de caractères que possède ce titre
            Tot1 = Tot1 + 1
            '
            'on crée une nouvelle classe
            ReDim Preserve cLnP(0 To Tot1)
            Set cLnP(Tot1) = New ClsLettrePetite
            '
            'on assigne un numéro de lettre en fonction de la lettre courante
            Select Case Mid(TitreSz1, i, 1)
                '
                Case "A"
                    '
                    LettreTmp = 1
                    '
                Case "Â"
                    '
                    LettreTmp = 1
                    '
                Case "Ä"
                    '
                    LettreTmp = 1
                    '
                Case "B"
                    '
                    LettreTmp = 2
                    '
                Case "C"
                    '
                    LettreTmp = 3
                    '
                Case "D"
                    '
                    LettreTmp = 4
                    '
                Case "E"
                    '
                    LettreTmp = 5
                    '
                Case "Ê"
                    '
                    LettreTmp = 5
                    '
                Case "Ë"
                    '
                    LettreTmp = 5
                    '
                Case "F"
                    '
                    LettreTmp = 6
                    '
                Case "G"
                    '
                    LettreTmp = 7
                    '
                Case "H"
                    '
                    LettreTmp = 8
                    '
                Case "I"
                    '
                    LettreTmp = 9
                    '
                Case "Î"
                    '
                    LettreTmp = 9
                    '
                Case "Ï"
                    '
                    LettreTmp = 9
                    '
                Case "J"
                    '
                    LettreTmp = 10
                    '
                Case "K"
                    '
                    LettreTmp = 11
                    '
                Case "L"
                    '
                    LettreTmp = 12
                    '
                Case "M"
                    '
                    LettreTmp = 13
                    '
                Case "N"
                    '
                    LettreTmp = 14
                    '
                Case "O"
                    '
                    LettreTmp = 15
                    '
                Case "Ô"
                    '
                    LettreTmp = 15
                    '
                Case "Ö"
                    '
                    LettreTmp = 15
                    '
                Case "P"
                    '
                    LettreTmp = 16
                    '
                Case "Q"
                    '
                    LettreTmp = 17
                    '
                Case "R"
                    '
                    LettreTmp = 18
                    '
                Case "S"
                    '
                    LettreTmp = 19
                    '
                Case "T"
                    '
                    LettreTmp = 20
                    '
                Case "U"
                    '
                    LettreTmp = 21
                    '
                Case "Û"
                    '
                    LettreTmp = 21
                    '
                Case "Ü"
                    '
                    LettreTmp = 21
                    '
                Case "V"
                    '
                    LettreTmp = 22
                    '
                Case "W"
                    '
                    LettreTmp = 23
                    '
                Case "X"
                    '
                    LettreTmp = 24
                    '
                Case "Y"
                    '
                    LettreTmp = 25
                    '
                Case "Z"
                    '
                    LettreTmp = 26
                    '
                Case "a"
                    '
                    LettreTmp = 37
                    '
                Case "â"
                    '
                    LettreTmp = 37
                    '
                Case "ä"
                    '
                    LettreTmp = 37
                    '
                Case "à"
                    '
                    LettreTmp = 37
                    '
                Case "b"
                    '
                    LettreTmp = 38
                    '
                Case "c"
                    '
                    LettreTmp = 74
                    '
                Case "d"
                    '
                    LettreTmp = 39
                    '
                Case "e"
                    '
                    LettreTmp = 40
                    '
                Case "é"
                    '
                    LettreTmp = 40
                    '
                Case "è"
                    '
                    LettreTmp = 40
                    '
                Case "ê"
                    '
                    LettreTmp = 40
                    '
                Case "ë"
                    '
                    LettreTmp = 40
                    '
                Case "f"
                    '
                    LettreTmp = 41
                    '
                Case "g"
                    '
                    LettreTmp = 42
                    '
                Case "h"
                    '
                    LettreTmp = 43
                    '
                Case "i"
                    '
                    LettreTmp = 44
                    '
                Case "î"
                    '
                    LettreTmp = 44
                    '
                Case "ï"
                    '
                    LettreTmp = 44
                    '
                Case "j"
                    '
                    LettreTmp = 45
                    '
                Case "k"
                    '
                    LettreTmp = 46
                    '
                Case "l"
                    '
                    LettreTmp = 47
                    '
                Case "m"
                    '
                    LettreTmp = 48
                    '
                Case "n"
                    '
                    LettreTmp = 49
                    '
                Case "o"
                    '
                    LettreTmp = 50
                    '
                Case "ô"
                    '
                    LettreTmp = 50
                    '
                Case "ö"
                    '
                    LettreTmp = 50
                    '
                Case "p"
                    '
                    LettreTmp = 51
                    '
                Case "q"
                    '
                    LettreTmp = 52
                    '
                Case "r"
                    '
                    LettreTmp = 53
                    '
                Case "s"
                    '
                    LettreTmp = 54
                    '
                Case "t"
                    '
                    LettreTmp = 55
                    '
                Case "u"
                    '
                    LettreTmp = 56
                    '
                Case "û"
                    '
                    LettreTmp = 56
                    '
                Case "ü"
                    '
                    LettreTmp = 56
                    '
                Case "ù"
                    '
                    LettreTmp = 56
                    '
                Case "v"
                    '
                    LettreTmp = 57
                    '
                Case "w"
                    '
                    LettreTmp = 58
                    '
                Case "x"
                    '
                    LettreTmp = 59
                    '
                Case "y"
                    '
                    LettreTmp = 60
                    '
                Case "0" 'ooo
                    '
                    LettreTmp = 61
                    '
                Case "0" 'zéro
                    '
                    LettreTmp = 27
                    '
                Case "1"
                    '
                    LettreTmp = 28
                    '
                Case "2"
                    '
                    LettreTmp = 29
                    '
                Case "3"
                    '
                    LettreTmp = 30
                    '
                Case "4"
                    '
                    LettreTmp = 31
                    '
                Case "5"
                    '
                    LettreTmp = 32
                    '
                Case "6"
                    '
                    LettreTmp = 33
                    '
                Case "7"
                    '
                    LettreTmp = 34
                    '
                Case "8"
                    '
                    LettreTmp = 35
                    '
                Case "9"
                    '
                    LettreTmp = 36
                    '
                Case "-"
                    '
                    LettreTmp = 62
                    '
                Case "+"
                    '
                    LettreTmp = 63
                    '
                Case "."
                    '
                    LettreTmp = 64
                    '
                Case "*"
                    '
                    LettreTmp = 65
                    '
                Case "/"
                    '
                    LettreTmp = 66
                    '
                Case "!"
                    '
                    LettreTmp = 67
                    '
                Case "="
                    '
                    LettreTmp = 68
                    '
                Case "("
                    '
                    LettreTmp = 69
                    '
                Case ")"
                    '
                    LettreTmp = 70
                    '
                Case "&"
                    '
                    LettreTmp = 71
                    '
                Case Chr(34) ' "
                    '
                    LettreTmp = 72
                    '
                Case Chr(39) ' '
                    '
                    LettreTmp = 73
                    '
                Case "@"
                    '
                    LettreTmp = 75
                    '
                Case ","
                    '
                    LettreTmp = 76
                    '
                Case "<"
                    '
                    LettreTmp = 77
                    '
                Case ">"
                    '
                    LettreTmp = 78
                    '
                Case "["
                    '
                    LettreTmp = 79
                    '
                Case "]"
                    '
                    LettreTmp = 80
                    '
                '
            End Select
            '
            iTmp = 1 + Rnd() * (LstLT1.Count - 1)
            '
            With cLnP(Tot1)
                .CarNorm = LettreTmp
                .XTmp = XSz1 + (i * (LargeurLettreSz / 18) * 14)
                .YTmp = YSz1
                .XSz = XSz1 + (LstLT1.Item(iTmp) * (LargeurLettreSz / 18) * 14)
                .YSz = YSz1
                .Etat = 1
                .tTps1 = 250 + (iTmp * 30)
                .ResteSz = .tTps1 - 250
                '.CarBiz = Rnd() * 56
                '.tTps1 = 0
            End With
            '
            LstLT1.Remove iTmp
            '
        End If
        '
    Next
    '
    'on crée le nombre de classes nécessaire pour afficher toutes les lettres du titre #2
    For i = 1 To Len(TitreSz2)
        '
        'on ne crée rien si c'est un espace
        If Mid(TitreSz2, i, 1) <> " " Then
            '
            'on incrément le nombre de caractères à afficher
            NbrCar = NbrCar + 1
            '
            'on incrémente le nombre de caractères que possède ce titre
            Tot2 = Tot2 + 1
            '
            'on crée une nouvelle classe
            ReDim Preserve cLn(0 To Tot2)
            Set cLn(Tot2) = New ClsLettreNormal
            '
            'on calcule la lettre par rapport à sa position dans le tableau des caractères de vb
            'LettreTmp = Asc(Mid(TitreSz2, i, 1)) - 96
            '
            'on fait un calcul afin de spécifier le bon chiffre (qui correspond à la bonne lettre)
            'If LettreTmp > -49 And LettreTmp < 0 Then
                '
                'LettreTmp = LettreTmp + 75
                '
            'End If
            '
            'on assigne un numéro de lettre en fonction de la lettre courante
            Select Case Mid(TitreSz2, i, 1)
                '
                Case "A"
                    '
                    LettreTmp = 1
                    '
                Case "Â"
                    '
                    LettreTmp = 1
                    '
                Case "Ä"
                    '
                    LettreTmp = 1
                    '
                Case "B"
                    '
                    LettreTmp = 2
                    '
                Case "C"
                    '
                    LettreTmp = 3
                    '
                Case "D"
                    '
                    LettreTmp = 4
                    '
                Case "E"
                    '
                    LettreTmp = 5
                    '
                Case "Ê"
                    '
                    LettreTmp = 5
                    '
                Case "Ë"
                    '
                    LettreTmp = 5
                    '
                Case "F"
                    '
                    LettreTmp = 6
                    '
                Case "G"
                    '
                    LettreTmp = 7
                    '
                Case "H"
                    '
                    LettreTmp = 8
                    '
                Case "I"
                    '
                    LettreTmp = 9
                    '
                Case "Î"
                    '
                    LettreTmp = 9
                    '
                Case "Ï"
                    '
                    LettreTmp = 9
                    '
                Case "J"
                    '
                    LettreTmp = 10
                    '
                Case "K"
                    '
                    LettreTmp = 11
                    '
                Case "L"
                    '
                    LettreTmp = 12
                    '
                Case "M"
                    '
                    LettreTmp = 13
                    '
                Case "N"
                    '
                    LettreTmp = 14
                    '
                Case "O"
                    '
                    LettreTmp = 15
                    '
                Case "Ô"
                    '
                    LettreTmp = 15
                    '
                Case "Ö"
                    '
                    LettreTmp = 15
                    '
                Case "P"
                    '
                    LettreTmp = 16
                    '
                Case "Q"
                    '
                    LettreTmp = 17
                    '
                Case "R"
                    '
                    LettreTmp = 18
                    '
                Case "S"
                    '
                    LettreTmp = 19
                    '
                Case "T"
                    '
                    LettreTmp = 20
                    '
                Case "U"
                    '
                    LettreTmp = 21
                    '
                Case "Û"
                    '
                    LettreTmp = 21
                    '
                Case "Ü"
                    '
                    LettreTmp = 21
                    '
                Case "V"
                    '
                    LettreTmp = 22
                    '
                Case "W"
                    '
                    LettreTmp = 23
                    '
                Case "X"
                    '
                    LettreTmp = 24
                    '
                Case "Y"
                    '
                    LettreTmp = 25
                    '
                Case "Z"
                    '
                    LettreTmp = 26
                    '
                Case "a"
                    '
                    LettreTmp = 37
                    '
                Case "â"
                    '
                    LettreTmp = 37
                    '
                Case "ä"
                    '
                    LettreTmp = 37
                    '
                Case "à"
                    '
                    LettreTmp = 37
                    '
                Case "b"
                    '
                    LettreTmp = 38
                    '
                Case "c"
                    '
                    LettreTmp = 74
                    '
                Case "d"
                    '
                    LettreTmp = 39
                    '
                Case "e"
                    '
                    LettreTmp = 40
                    '
                Case "é"
                    '
                    LettreTmp = 40
                    '
                Case "è"
                    '
                    LettreTmp = 40
                    '
                Case "ê"
                    '
                    LettreTmp = 40
                    '
                Case "ë"
                    '
                    LettreTmp = 40
                    '
                Case "f"
                    '
                    LettreTmp = 41
                    '
                Case "g"
                    '
                    LettreTmp = 42
                    '
                Case "h"
                    '
                    LettreTmp = 43
                    '
                Case "i"
                    '
                    LettreTmp = 44
                    '
                Case "î"
                    '
                    LettreTmp = 44
                    '
                Case "ï"
                    '
                    LettreTmp = 44
                    '
                Case "j"
                    '
                    LettreTmp = 45
                    '
                Case "k"
                    '
                    LettreTmp = 46
                    '
                Case "l"
                    '
                    LettreTmp = 47
                    '
                Case "m"
                    '
                    LettreTmp = 48
                    '
                Case "n"
                    '
                    LettreTmp = 49
                    '
                Case "o"
                    '
                    LettreTmp = 50
                    '
                Case "ô"
                    '
                    LettreTmp = 50
                    '
                Case "ö"
                    '
                    LettreTmp = 50
                    '
                Case "p"
                    '
                    LettreTmp = 51
                    '
                Case "q"
                    '
                    LettreTmp = 52
                    '
                Case "r"
                    '
                    LettreTmp = 53
                    '
                Case "s"
                    '
                    LettreTmp = 54
                    '
                Case "t"
                    '
                    LettreTmp = 55
                    '
                Case "u"
                    '
                    LettreTmp = 56
                    '
                Case "û"
                    '
                    LettreTmp = 56
                    '
                Case "ü"
                    '
                    LettreTmp = 56
                    '
                Case "ù"
                    '
                    LettreTmp = 56
                    '
                Case "v"
                    '
                    LettreTmp = 57
                    '
                Case "w"
                    '
                    LettreTmp = 58
                    '
                Case "x"
                    '
                    LettreTmp = 59
                    '
                Case "y"
                    '
                    LettreTmp = 60
                    '
                Case "0" 'ooo
                    '
                    LettreTmp = 61
                    '
                Case "0" 'zéro
                    '
                    LettreTmp = 27
                    '
                Case "1"
                    '
                    LettreTmp = 28
                    '
                Case "2"
                    '
                    LettreTmp = 29
                    '
                Case "3"
                    '
                    LettreTmp = 30
                    '
                Case "4"
                    '
                    LettreTmp = 31
                    '
                Case "5"
                    '
                    LettreTmp = 32
                    '
                Case "6"
                    '
                    LettreTmp = 33
                    '
                Case "7"
                    '
                    LettreTmp = 34
                    '
                Case "8"
                    '
                    LettreTmp = 35
                    '
                Case "9"
                    '
                    LettreTmp = 36
                    '
                Case "-"
                    '
                    LettreTmp = 62
                    '
                Case "+"
                    '
                    LettreTmp = 63
                    '
                Case "."
                    '
                    LettreTmp = 64
                    '
                Case "*"
                    '
                    LettreTmp = 65
                    '
                Case "/"
                    '
                    LettreTmp = 66
                    '
                Case "!"
                    '
                    LettreTmp = 67
                    '
                Case "="
                    '
                    LettreTmp = 68
                    '
                Case "("
                    '
                    LettreTmp = 69
                    '
                Case ")"
                    '
                    LettreTmp = 70
                    '
                Case "&"
                    '
                    LettreTmp = 71
                    '
                Case Chr(34) ' "
                    '
                    LettreTmp = 72
                    '
                Case Chr(39) ' '
                    '
                    LettreTmp = 73
                    '
                Case "@"
                    '
                    LettreTmp = 75
                    '
                Case ","
                    '
                    LettreTmp = 76
                    '
                Case "<"
                    '
                    LettreTmp = 77
                    '
                Case ">"
                    '
                    LettreTmp = 78
                    '
                Case "["
                    '
                    LettreTmp = 79
                    '
                Case "]"
                    '
                    LettreTmp = 80
                    '
                '
            End Select
            '
            cLn(Tot2).CarNorm = LettreTmp
            cLn(Tot2).XSz = XSz2 + (i * (LargeurLettreSz / 18) * 19)
            cLn(Tot2).YSz = YSz2
            cLn(Tot2).Etat = 1
            cLn(Tot2).CarBiz = Rnd() * 56
            'cLn(Tot2).tTps1 = 0
            '
        End If
        '
    Next
    '
    'on configure les effets
    cFX1.Etat = 1
    cFX1.InitFX
    '
    cLT1.YSz = YSz2
    cLT1.PosPremL = XSz2
    cLT1.InitLigne
    '
    cFX3.Etat = 1
    cFX3.YSz = YSz2 - 2
    '
End Sub
'
'procédure qui va afficher le titre
Public Sub AfficherTitre()
    '
    'si l'état du titre = 0, on quitte la procédure (il n'y a rien à afficher)
    If Etat = 0 Then Exit Sub
    '
    'on affiche l'effet #1 (et #2)
    cFX1.AfficherEffet
    '
    'on affiche les lettres du titre #1
    For i = 1 To Tot1
        '
        cLnP(i).AfficherCar NbrCarFin
        '
    Next
    '
    'on affiche les lettres du titre #2
    For i = 1 To Tot2
        '
        cLn(i).AfficherCar NbrCarFin
        '
    Next
    '
    'on affiche la ligne de caractères spéciaux
    cLT1.Afficher
    '
    'on affiche l'effet #3
    cFX3.AfficherEffet
    '
    'on vérifie si tous les caractères ont disparu
    If NbrCarFin = NbrCar Then
        '
        'on initialise la classe (on vérifie si il y a encore un texte à afficher et on fait en fonction...)
        InitTitre2
        '
    End If
    '
End Sub

Private Sub Class_Initialize()
    '
    'InitTitre2
    '
End Sub
