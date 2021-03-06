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
'* CETTE CLASSE GERE L'EFFET N� 1                                                 *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'stocke la g�om�trie d'un objet
Dim VBuffer As Direct3DVertexBuffer8
'
'dur�e des transitions (ms)
Dim Tps1 As Long
Dim tTps1 As Long
'
'opacit� de l'effet
Public OpaciteSz As Long
'
'coordonn�es de l'effet
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
'�tat de l'effet
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
'proc�dure permettant de changer al�atoirement les valeurs pour les effets #2
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
'proc�dure permettant d'afficher une image
Private Sub AfficherTypeFX(xFX As Single, yFX As Single, X As Single, Y As Single, HSz As Long, LSz As Long, Optional CouleurSz As Long = vbWhite)
    '
    Dim cTmp As Long
    '
    'on demande au vertex shader d'utiliser le format de vertex appropri�
    D3DDevice.SetVertexShader FVF_TLVERTEX
    '
    'on d�termine la couleur (on ajoute la valeur alpha)
    CouleurSz = Hex2Long(Hex(OpaciteSz) & Hex(CouleurSz))
    '
    'on g�n�re les points
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
    'on incr�mente de 1 le nombre d'objets affich�s
    NbrObjetsAff = NbrObjetsAff + 1
    '
End Sub
'
'proc�dure qui va afficher l'effet
Public Sub AfficherEffet()
    '
    'si l'�tat de l'effet = 0, on quitte la proc�dure (il n'y a rien � afficher)
    If Etat = 0 Then Exit Sub
    '
    'on effectue les op�rations en fontion de l'�tat de l'effet
    Select Case Etat
        '
        Case 1 'on fait apparaitre l'effet
            '
            'on augmente l'opacit� et la taille de l'effet en fonction de la vitesse g�n�rale
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
        Case 3 'on r�duit l'opacit� de l'effet
            '
            'on r�duit l'opacit� en fonction de la vitesse g�n�rale
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
                'on remet le compteur � z�ro
                'CompteurSz = 0
                '
            End If
            '
        '
    End Select
    '
    'on lance les effets #2 en fonction du temps �coul�
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

