VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLettre1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE UN CARACTERE DE TYPE "KATANA" POUR L'AFFICHAGE DES TITRES    *
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
'opacité du caractère vert
Public OpaciteSz As Long
'
'opacité du caractère blanc
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
'caractère
Public CarSz As Long
'
'variables quelconques
Dim i As Integer
Dim i2 As Single
Dim prctSz As Single
'
'
'initialisation de la classe
Private Sub Class_Initialize()
    '
    'on initialise les variables de base
    Etat = 1
    tTps1 = 30
    '
End Sub
'
'procédure permettant d'afficher une image
Private Sub AfficherCarSz1(NumSz As Long, X As Single, Y As Single, Optional CouleurSz As Long = vbWhite)
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
    CouleurSz = Hex2Long(Hex(OpaciteSz) & Hex(CouleurSz))
    '
    'on génère les points
    vertMatrix(0) = CreateTLVertex(X, Y, 0, 1, CouleurSz, 0, (1 / 512) * CharX, (1 / 512) * CharY)
    vertMatrix(1) = CreateTLVertex(X + LargeurLettreSz, Y, 0, 1, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), (1 / 512) * CharY)
    vertMatrix(2) = CreateTLVertex(X, Y + HauteurLettreSz, 0, 1, CouleurSz, 0, (1 / 512) * CharX, ((1 / 512) * CharY) + ((1 / 512) * 58))
    vertMatrix(3) = CreateTLVertex(X + LargeurLettreSz, Y + HauteurLettreSz, 0, 1, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), ((1 / 512) * CharY) + ((1 / 512) * 58))
    '
    'on dessine l'image
    D3DDevice.SetTexture 0, MatrixTex_Blanc2
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
    'on calcule les coordonnées du caractère
    cTmp = (NumSz - 1) \ 10
    CharY = 58 * cTmp
    cTmp = (NumSz - 1) Mod 10
    CharX = 50 * cTmp
    '
    'on détermine la couleur (on ajoute la valeur alpha)
    CouleurSz = Hex2Long(Hex(OpaciteSz2) & Hex(CouleurSz))
    '
    'on demande au vertex shader d'utiliser le format de vertex approprié
    D3DDevice.SetVertexShader FVF_TLVERTEX
    '
    'on génère les points
    vertMatrix(0) = CreateTLVertex(X, Y, 0, 1, CouleurSz, 0, (1 / 512) * CharX, (1 / 512) * CharY)
    vertMatrix(1) = CreateTLVertex(X + LargeurLettreSz, Y, 0, 1, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), (1 / 512) * CharY)
    vertMatrix(2) = CreateTLVertex(X, Y + HauteurLettreSz, 0, 1, CouleurSz, 0, (1 / 512) * CharX, ((1 / 512) * CharY) + ((1 / 512) * 58))
    vertMatrix(3) = CreateTLVertex(X + LargeurLettreSz, Y + HauteurLettreSz, 0, 1, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), ((1 / 512) * CharY) + ((1 / 512) * 58))
    '
    'on dessine l'image
    D3DDevice.SetTexture 0, MatrixTex_Blanc
    '
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertMatrix(0), Len(vertMatrix(0))
    '
    'on incrémente de 1 le nombre d'objets affichés
    NbrObjetsAff = NbrObjetsAff + 1
    '
End Sub
'
'procédure qui va afficher l'effet
Public Sub AfficherCar()
    '
    'dMsg = OpaciteSz & " - " & OpaciteSz2
    '
    'si l'état de l'effet = 0, on quitte la procédure (il n'y a rien à afficher)
    If Etat = 0 Then Exit Sub
    '
    'on effectue les opérations en fontion de l'état de l'effet
    Select Case Etat
        '
        Case 1 'on augmente l'opacité du caractère vert jusqu'à 50%
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then Tps1 = tTps1
            '
            'on calcule le pourcentage
            prctSz = (Tps1 / tTps1) * 50
            '
            OpaciteSz = 2.55 * prctSz
            '
            'on affiche le caractère vert
            AfficherCarSz1 CarSz, XSz, YSz
            '
            'si on a atteint les 50%, on passe à l'étape suivante
            If prctSz = 50 Then
                '
                Etat = 2
                Tps1 = 0
                'tTps1 = 30
                '
            End If
            '
        Case 2 'on continue d'augmenter l'opacite du caractère vert et on augment celle du blanc
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then Tps1 = tTps1
            '
            'on calcule le pourcentage
            prctSz = (Tps1 / tTps1) * 50
            '
            OpaciteSz = 2.55 * prctSz + 127.5
            OpaciteSz2 = 2.55 * prctSz
            '
            'on affiche les caractères
            AfficherCarSz1 CarSz, XSz, YSz
            AfficherCarSz2 CarSz, XSz, YSz
            '
            'si on a atteint les 50%, on passe à l'étape suivante
            If prctSz = 50 Then
                '
                Etat = 3
                Tps1 = 0
                'tTps1 = 50
                '
            End If
            '
        Case 3 'on continue d'augmenter celle du blanc
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then Tps1 = tTps1
            '
            'on calcule le pourcentage
            prctSz = (Tps1 / tTps1) * 50
            '
            OpaciteSz2 = 2.55 * prctSz + 127.5
            '
            'on affiche les caractères
            AfficherCarSz1 CarSz, XSz, YSz
            AfficherCarSz2 CarSz, XSz, YSz
            '
            'si on a atteint les 50%, on passe à l'étape suivante
            If prctSz = 50 Then
                '
                Etat = 4
                Tps1 = 0
                tTps1 = 10
                '
            End If
            '
        Case 4 'on maintient l'opacité des caractères
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on affiche le caractère blanc
            AfficherCarSz2 CarSz, XSz, YSz
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 >= tTps1 Then
                '
                Etat = 5
                Tps1 = 0
                tTps1 = 50
                '
            End If
            '
        Case 5 'on diminue l'opacité du caractère blanc jusqu'à 50%
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then Tps1 = tTps1
            '
            'on calcule le pourcentage
            prctSz = (Tps1 / tTps1) * 50
            '
            OpaciteSz2 = 255 - (2.55 * prctSz)
            '
            'on affiche les caractères
            AfficherCarSz1 CarSz, XSz, YSz
            AfficherCarSz2 CarSz, XSz, YSz
            '
            'si on a atteint les 50%, on passe à l'étape suivante
            If prctSz = 50 Then
                '
                Etat = 6
                Tps1 = 0
                'tTps1 = 40
                '
            End If
            '
        Case 6 'on continue à diminuer l'opacité du caractère blanc et on commence à diminuer celle du caractère vert
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then Tps1 = tTps1
            '
            'on calcule le pourcentage
            prctSz = (Tps1 / tTps1) * 50
            '
            OpaciteSz2 = 127.5 - (2.55 * prctSz)
            OpaciteSz = 255 - (2.55 * prctSz)
            '
            'on affiche les caractères
            AfficherCarSz1 CarSz, XSz, YSz
            AfficherCarSz2 CarSz, XSz, YSz
            '
            'si on a atteint les 50%, on passe à l'étape suivante
            If prctSz = 50 Then
                '
                Etat = 7
                Tps1 = 0
                'tTps1 = 30
                '
            End If
            '
        Case 7 'on réduit l'opacité du caractère vert
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then Tps1 = tTps1
            '
            'on calcule le pourcentage
            prctSz = (Tps1 / tTps1) * 50
            '
            OpaciteSz = 127.5 - (2.55 * prctSz)
            '
            'on affiche le caractère vert
            AfficherCarSz1 CarSz, XSz, YSz
            '
            'si on a atteint les 50%, on revient à la première étape
            If prctSz = 50 Then
                '
                Etat = 8
                OpaciteSz = 0
                OpaciteSz2 = 0
                Tps1 = 0
                tTps1 = 500 + Rnd() * 1000
                '
            End If
            '
        Case 8 'on attend
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 >= tTps1 Then
                '
                Etat = 9
                Tps1 = 0
                tTps1 = 150 + Rnd() * 200
                '
            End If
            '
        Case 9 'on incrémente l'opacité du caractère vert (40%)
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then Tps1 = tTps1
            '
            'on calcule le pourcentage
            prctSz = (Tps1 / tTps1) * 40
            '
            OpaciteSz = 2.55 * prctSz
            '
            'on affiche le caractère vert
            AfficherCarSz1 CarSz, XSz, YSz
            '
            'si on a atteint les 40%, on passe à l'étape suivante
            If prctSz = 40 Then
                '
                Etat = 10
                Tps1 = 0
                'tTps1 = 30
                '
            End If
            '
        Case 10 'on réduit l'opacité des caractères verts à partir de 30
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then Tps1 = tTps1
            '
            'on calcule le pourcentage
            prctSz = (Tps1 / tTps1) * 30
            '
            OpaciteSz = 76.5 - (2.55 * prctSz)
            '
            'on affiche le caractère vert
            AfficherCarSz1 CarSz, XSz, YSz
            '
            'si on a atteint les 30%, on revient à la première étape
            If prctSz = 30 Then
                '
                Etat = 0
                OpaciteSz = 0
                Tps1 = 0
                tTps1 = 30
                '
            End If
            '
        '
    End Select
    '
End Sub



