VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLettreNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE UN CARACTERE NORMAL POUR L'AFFICHAGE DES TITRES              *
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
Public tTps1 As Long
Dim Tps2 As Long
Public tTps2 As Long
'
'faut-il afficher le caractère bizarre ou pas
Dim ChngSz As Integer
'
'reste lors du changement de caractère
Dim ResteSz As Integer
'
'opacité du caractère normal
Public OpaciteSz As Long
'
'opacité du caractère "bizarre"
Public OpaciteSz2 As Long
'
'coordonnées du caractère
Public XSz As Single
Public YSz As Single
'
'couleur du caractère
'Public RSz As Long
'Public VSz As Long
'Public BSz As Long
'
'état du caractère
Public Etat As Integer
'
'caractères
Public CarNorm As Long
Public CarBiz As Long
'
'variables quelconques
Dim prctSz As Single
'
'initialisation de la classe
Private Sub Class_Initialize()
    '
    tTps1 = 250
    tTps2 = 20 + Rnd() * 500
    'tTps2 = 60
    ChngSz = Rnd() * 1
    '
End Sub
'
'procédure permettant d'afficher une lettre
Private Sub AfficherLettre(NuméroLettre As Long, X As Single, Y As Single, Optional CouleurSz As Long = vbWhite)
    '
    Dim CharX As Integer
    Dim CharY As Integer
    Dim cTmp As Long
    '
    'on demande au vertex shader d'utiliser le format de vertex approprié
    D3DDevice.SetVertexShader FVF_TLVERTEX
    '
    'on calcule les coordonnées de la lettre
    cTmp = (NuméroLettre - 1) \ 10
    CharY = 58 * cTmp
    cTmp = (NuméroLettre - 1) Mod 10
    CharX = 50 * cTmp
    '
    'on détermine la couleur (on ajoute la valeur alpha)
    CouleurSz = Hex2Long(Hex(OpaciteSz) & Hex(CouleurSz))
    '
    'on génère les points
    vertMatrix(0) = CreateTLVertex(X, Y, 0, 1, CouleurSz, 0, (1 / 512) * CharX, (1 / 512) * CharY)
    vertMatrix(1) = CreateTLVertex(X + LargeurLettreSz, Y, 0, 1, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), (1 / 512) * CharY)
    vertMatrix(2) = CreateTLVertex(X, Y + HauteurLettreSz, 0, 1, CouleurSz, 0, (1 / 512) * CharX, ((1 / 512) * CharY) + ((1 / 512) * 58))
    vertMatrix(3) = CreateTLVertex(X + LargeurLettreSz, Y + HauteurLettreSz, 0, 1, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), ((1 / 512) * CharY) + ((1 / 512) * 58))
    '
    'on dessine l'image
    D3DDevice.SetTexture 0, MatrixTex_Normal
    '
    'Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
    'Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE)
    'Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_CURRENT)
    '
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertMatrix(0), Len(vertMatrix(0))
    '
    'on incrémente de 1 le nombre d'objets affichés
    NbrObjetsAff = NbrObjetsAff + 1
    '
End Sub
'
'procédure permettant d'afficher une image
Private Sub AfficherCarSz2(NumSz As Long, X As Single, Y As Single, Optional CouleurSz As Long = vbWhite)
    '
    Dim cTmp As Long
    '
    'on demande au vertex shader d'utiliser le format de vertex approprié
    D3DDevice.SetVertexShader FVF_TLVERTEX
    '
    'on calcule les coordonnées du caractère
    cTmp = (NumSz - 1) \ 10
    CharY = 58 * cTmp
    cTmp = (NumSz - 1) Mod 10
    CharX = 50 * cTmp
    '
    'on détermine la couleur (on ajoute la valeur alpha)
    CouleurSz = Hex2Long(Hex(OpaciteSz2) & Hex(CouleurSz))
    '
    'on génère les points
    vertMatrix(0) = CreateTLVertex(X, Y, 0, 1, CouleurSz, 0, (1 / 512) * CharX, (1 / 512) * CharY)
    vertMatrix(1) = CreateTLVertex(X + LargeurLettreSz, Y, 0, 1, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), (1 / 512) * CharY)
    vertMatrix(2) = CreateTLVertex(X, Y + HauteurLettreSz, 0, 1, CouleurSz, 0, (1 / 512) * CharX, ((1 / 512) * CharY) + ((1 / 512) * 58))
    vertMatrix(3) = CreateTLVertex(X + LargeurLettreSz, Y + HauteurLettreSz, 0, 1, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), ((1 / 512) * CharY) + ((1 / 512) * 58))
    '
    'on dessine l'image
    D3DDevice.SetTexture 0, MatrixTex_Blanc3
    '
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertMatrix(0), Len(vertMatrix(0))
    '
    'on incrémente de 1 le nombre d'objets affichés
    NbrObjetsAff = NbrObjetsAff + 1
    '
End Sub
'
'procédure qui va afficher la lettre
Public Sub AfficherCar(NbrTmp As Integer)
    '
    'si l'état de l'effet = 0, on quitte la procédure (il n'y a rien à afficher)
    If Etat = 0 Then Exit Sub
    '
    'dMsg = OpaciteSz & " - " & OpaciteSz2
    '
    'on effectue les opérations en fontion de l'état de l'effet
    Select Case Etat
        '
        Case 1 'on attend un peu
            '
            'on incrémente la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie si le temps est écoulé
            If Tps1 >= tTps1 Then
                '
                Etat = 2
                Tps1 = 0
                tTps1 = 400
                '
            End If
            '
        Case 2 'on augmente l'opacité du caractère bizarre (ou du normal) et on le change
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 >= tTps1 Then
                '
                'on ajuste le pourcentage
                prctSz = 100
                '
                Etat = 3
                Tps1 = 0
                tTps1 = 30 + (Rnd() * 1000)
                ResteSz = tTps1
                '
            Else
                '
                'on calcule le pourcentage
                prctSz = (Tps1 / tTps1) * 100
                '
            End If
            '
            'on vérifie si on doit afficher le caractère bizarre
            If ChngSz = 1 Then
                '
                OpaciteSz2 = 2.55 * prctSz
                '
                'on affiche le caractère bizarre
                AfficherCarSz2 CarBiz, XSz, YSz
                '
                'on change la lettre bizarre en fonction du temps spécifié (tTps2)
                Tps2 = Tps2 + (1 * VitesseGenSz)
                '
                If Tps2 >= tTps2 Then
                    '
                    CarBiz = Rnd() * 56
                    Tps2 = 0
                    'tTps2 = (Rnd() * 250)
                    '
                End If
                '
            Else
                '
                OpaciteSz = 2.55 * prctSz
                '
                'on affiche le caractère normal
                AfficherLettre CarNorm, XSz, YSz
                '
            End If
            '
        Case 3 'on attend un moment et on change le caractère suivant la vitesse spécifiée
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie si le temps est écoulé
            If Tps1 >= tTps1 Then
                '
                Etat = 4
                Tps1 = 0
                tTps1 = 600
                '
            End If
            '
            'on vérifie si on doit afficher le caractère bizarre
            If ChngSz = 1 Then
                '
                'on affiche le caractère bizarre
                AfficherCarSz2 CarBiz, XSz, YSz
                '
                'on change la lettre bizarre en fonction du temps spécifié (tTps2)
                Tps2 = Tps2 + (1 * VitesseGenSz)
                '
                If Tps2 >= tTps2 Then
                    '
                    CarBiz = Rnd() * 56
                    Tps2 = 0
                    'tTps2 = Rnd() * 250
                    '
                End If
                '
            Else
                '
                'on affiche le caractère normal
                AfficherLettre CarNorm, XSz, YSz
                '
            End If
            '
        Case 4 'on augmente l'opacité du caractère blanc et on diminue celle du caractère bizarre
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then
                '
                'on ajuste le pourcentage
                prctSz = 100
                '
                Etat = 5
                Tps1 = 0
                tTps1 = 3000 - ResteSz
                '
            Else
                '
                'on calcule le pourcentage
                prctSz = (Tps1 / tTps1) * 100
                '
            End If
            '
            'on vérifie si on doit afficher le caractère bizarre
            If ChngSz = 1 Then
                '
                OpaciteSz = 2.55 * prctSz
                OpaciteSz2 = 255 - (2.55 * prctSz)
                '
                'on affiche les caractères
                AfficherLettre CarNorm, XSz, YSz
                AfficherCarSz2 CarBiz, XSz, YSz
                '
            Else
                '
                'on affiche le caractère normal
                AfficherLettre CarNorm, XSz, YSz
                '
            End If
            '
        Case 5 'on attend
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on affiche le caractère blanc
            AfficherLettre CarNorm, XSz, YSz
            '
            'on vérifie si le temps est écoulé
            If Tps1 >= tTps1 Then
                '
                Etat = 6
                Tps1 = 0
                tTps1 = 500
                '
            End If
            '
        Case 6 'on diminue l'opacité du caractère blanc
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then
                '
                'on ajuste le pourcentage
                prctSz = 100
                '
                Etat = 0
                Tps1 = 0
                tTps1 = 250
                OpaciteSz = 0
                OpaciteSz2 = 0
                NbrTmp = NbrTmp + 1
                '
            Else
                '
                'on calcule le pourcentage
                prctSz = (Tps1 / tTps1) * 100
                '
            End If
            '
            OpaciteSz = 255 - (2.55 * prctSz)
            '
            'on affiche les caractères
            AfficherLettre CarNorm, XSz, YSz
        '
    End Select
    '
End Sub
