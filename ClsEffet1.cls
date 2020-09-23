VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsEffet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE L'EFFET N° 1                                                 *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'stocke la géométrie d'un objet
Dim VBuffer As Direct3DVertexBuffer8
'
'durée des transitions (ms)
Dim Tps1 As Long
Dim tTps1 As Long
'
'opacité de l'effet
Public OpaciteSz As Long
'
'coordonnées de l'effet
Public XSz As Single
Public YSz As Single
Dim XTmp As Single
Dim YTmp As Single
'
'dimensions
Public HauteurSz As Long
Public LargeurSz As Long
Public HTmp As Long
Public LTmp As Long
'
'couleur de l'effet
'Public RSz As Long
'Public VSz As Long
'Public BSz As Long
'
'état de l'effet
Public Etat As Integer
'
'variables quelconques
Dim i As Integer
Dim i2 As Single
Dim prctSz As Single
'
'compteur (permet d'avoir une mesure de temps)
Dim CompteurSz As Long
'
Dim cFX2(5) As New ClsEffet2
'
'initialisation de la classe
Private Sub Class_Initialize()
    '
    'on initialise les variables de base
    Etat = 1
    HTmp = ((49 / 21) * HauteurLettreSz)
    LTmp = ((512 / 18) * LargeurLettreSz)
    XSz = (DimL / 2) - (LTmp / 2)
    YSz = (DimH / 2) - (HTmp / 2)
    HauteurSz = HTmp
    LargeurSz = LTmp
    '
    'on initialise les valeurs pour les effets #2
    InitFX
    '
End Sub
'
'procédure permettant de changer aléatoirement les valeurs pour les effets #2
Public Sub InitFX()
    '
    i = ((16.5 / 21) * HauteurLettreSz)
    '
    With cFX2(0)
        .Etat = 1
        .DirectionSz = 0
        .XSz = XSz
        .YSz = YSz + i
        .tTps1 = 100 + Rnd() * 200
        .TailleSz = 2 + Rnd() * 2
    End With
    '
    With cFX2(1)
        .Etat = 1
        .DirectionSz = (25 / 21) * HauteurLettreSz
        .XSz = XSz
        .YSz = YSz + i
        .tTps1 = 100 + Rnd() * 200
        .TailleSz = 1 + Rnd() * 1
    End With
    '
    With cFX2(2)
        .Etat = 1
        .DirectionSz = (-20 / 21) * HauteurLettreSz
        .XSz = XSz
        .YSz = YSz + i
        .tTps1 = 100 + Rnd() * 200
        .TailleSz = 1 + Rnd() * 1
    End With
    '
    With cFX2(3)
        .Etat = 1
        .DirectionSz = (17 / 21) * HauteurLettreSz
        .XSz = XSz
        .YSz = YSz + i
        .tTps1 = 100 + Rnd() * 200
        .TailleSz = 1 + Rnd() * 1
    End With
    '
    With cFX2(4)
        .Etat = 1
        .DirectionSz = (-16 / 21) * HauteurLettreSz
        .XSz = XSz
        .YSz = YSz + i
        .tTps1 = 100 + Rnd() * 200
        .TailleSz = 1 + Rnd() * 1
    End With
    '
    With cFX2(5)
        .Etat = 1
        .DirectionSz = (15 / 21) * HauteurLettreSz
        .XSz = XSz
        .YSz = YSz + i
        .tTps1 = 100 + Rnd() * 200
        .TailleSz = 1 + Rnd() * 1
    End With
    '
    Tps1 = 0
    tTps1 = 300
    '
End Sub
'
'procédure permettant d'afficher une image
Private Sub AfficherTypeFX(xFX As Single, yFX As Single, X As Single, Y As Single, HSz As Long, LSz As Long, Optional CouleurSz As Long = vbWhite)
    '
    Dim cTmp As Long
    '
    'on demande au vertex shader d'utiliser le format de vertex approprié
    D3DDevice.SetVertexShader FVF_TLVERTEX
    '
    'on détermine la couleur (on ajoute la valeur alpha)
    CouleurSz = Hex2Long(Hex(OpaciteSz) & Hex(CouleurSz))
    '
    'on génère les points
    vertMatrix(0) = CreateTLVertex(X, Y, 0, 1, CouleurSz, 0, (1 / 512) * xFX, (1 / 512) * yFX)
    vertMatrix(1) = CreateTLVertex(X + LargeurSz, Y, 0, 1, CouleurSz, 0, ((1 / 512) * xFX) + ((1 / 512) * LSz), (1 / 512) * yFX)
    vertMatrix(2) = CreateTLVertex(X, Y + HauteurSz, 0, 1, CouleurSz, 0, (1 / 512) * xFX, ((1 / 512) * yFX) + ((1 / 512) * HSz))
    vertMatrix(3) = CreateTLVertex(X + LargeurSz, Y + HauteurSz, 0, 1, CouleurSz, 0, ((1 / 512) * xFX) + ((1 / 512) * LSz), ((1 / 512) * yFX) + ((1 / 512) * HSz))
    '
    'on dessine l'image
    D3DDevice.SetTexture 0, MatrixTex_Effets
    '
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertMatrix(0), Len(vertMatrix(0))
    '
    'on incrémente de 1 le nombre d'objets affichés
    NbrObjetsAff = NbrObjetsAff + 1
    '
End Sub
'
'procédure qui va afficher l'effet
Public Sub AfficherEffet()
    '
    'si l'état de l'effet = 0, on quitte la procédure (il n'y a rien à afficher)
    If Etat = 0 Then Exit Sub
    '
    'on effectue les opérations en fontion de l'état de l'effet
    Select Case Etat
        '
        Case 1 'on fait apparaitre l'effet
            '
            'on augmente l'opacité et la taille de l'effet en fonction de la vitesse générale
            If Tps1 < tTps1 Then
                '
                Tps1 = Tps1 + (1 * VitesseGenSz)
                If Tps1 > tTps1 Then Tps1 = tTps1
                '
                prctSz = (Tps1 / tTps1) * 100
                prctSz = prctSz * (2 - (1 / 100 * prctSz))
                '
                OpaciteSz = 2.55 * prctSz
                '
                HauteurSz = (HTmp / 2) + ((HTmp / 200) * prctSz)
                LargeurSz = (LTmp / 2) + ((LTmp / 200) * prctSz)
                '
                YTmp = YSz + (HTmp / 4) - ((HTmp / 400) * prctSz)
                XTmp = XSz + (LTmp / 4) - ((LTmp / 400) * prctSz)
                '
            Else
                '
                OpaciteSz = 255
                Etat = 2
                Tps1 = 0
                tTps1 = tTps1 / 2
                '
            End If
            '
        Case 2 'on maintient l'effet
            '
            If Tps1 < tTps1 Then
                '
                Tps1 = Tps1 + (1 * VitesseGenSz)
                '
            Else
                '
                Tps1 = 0
                tTps1 = tTps1 * 4
                Etat = 3
                '
            End If
            '
        Case 3 'on réduit l'opacité de l'effet
            '
            'on réduit l'opacité en fonction de la vitesse générale
            If Tps1 < tTps1 Then
                '
                Tps1 = Tps1 + (1 * VitesseGenSz)
                If Tps1 > tTps1 Then Tps1 = tTps1
                '
                OpaciteSz = 255 - (2.55 * ((Tps1 / tTps1) * 100))
                '
            Else
                '
                OpaciteSz = 0
                Etat = 0
                Tps1 = 0
                tTps1 = tTps1 / 2
                '
                'on remet le compteur à zéro
                'CompteurSz = 0
                '
            End If
            '
        '
    End Select
    '
    'on lance les effets #2 en fonction du temps écoulé
    'if compteursz > 200 and <
    '
    'on affiche les effets #1
    AfficherTypeFX 0, 0, XTmp, YTmp, 49, 512
    '
    cFX2(0).AfficherEffet
    cFX2(1).AfficherEffet
    cFX2(2).AfficherEffet
    cFX2(3).AfficherEffet
    cFX2(4).AfficherEffet
    cFX2(5).AfficherEffet
    '
End Sub

