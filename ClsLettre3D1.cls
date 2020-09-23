VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLettre3D1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE UNE LETTRE 3D                                                *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'opacités des caractères
Public OpaciteSz As Long
Public OpaciteSz2 As Long
'
'décalage du temps des transitions
Public DecalSz As Long
'
'numéro du caractère
Public NumLettreSz As Long
'
'coordonnées du caractère
Public XSz As Single
Public YSz As Single
Public ZSz As Single
'
'couleur du caractère
'Dim RSz As Long
'Dim VSz As Long
'Dim BSz As Long
'
'transitions
Public tTps1 As Long
Public Tps1 As Long
Public tTps2 As Long
Public Tps2 As Long
'
'le caractère peut il changer dans le temps
Public ChangementSz As Long
'
'variations de couleur
Public VarCouleur As Long
'
'état du caractère
Public Etat As Long
'
'variable d'incrémentation
Dim i As Integer
'
'
'procédure permettan d'affiche une image
Private Sub AfficherCar(NuméroFonte As Long, Optional CouleurSz As Long = vbWhite)
    '
    Dim CharX As Integer
    Dim CharY As Integer
    Dim cTmp As Long
    '
    'on calcule les coordonnées du caractère
    cTmp = (NuméroFonte - 1) \ 10
    CharY = 58 * cTmp
    cTmp = (NuméroFonte - 1) Mod 10
    CharX = 50 * cTmp
    '
    'on détermine la couleur (on ajoute la valeur alpha)
    CouleurSz = Hex2Long(Hex(OpaciteSz) & Hex(CouleurSz))
    '
    'on génère les points
    vertMatrix3D(0) = CreateLitVertex(XSz, YSz, ZSz, CouleurSz, 0, (1 / 512) * CharX, (1 / 512) * CharY)
    vertMatrix3D(1) = CreateLitVertex(XSz + LargeurLettreSz, YSz, ZSz, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), (1 / 512) * CharY)
    vertMatrix3D(2) = CreateLitVertex(XSz, YSz + HauteurLettreSz, ZSz, CouleurSz, 0, (1 / 512) * CharX, ((1 / 512) * CharY) + ((1 / 512) * 58))
    vertMatrix3D(3) = CreateLitVertex(XSz + LargeurLettreSz, YSz + HauteurLettreSz, ZSz, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), ((1 / 512) * CharY) + ((1 / 512) * 58))
    '
    'on vérifie qu'un des points sera visible
    For i = 0 To 3
        '
        If TestVis(MakeVector(vertMatrix3D(i).X, vertMatrix3D(i).Y, vertMatrix3D(i).Z)) = True Then
            '
            'on va à l'affichage
            GoTo AFF_PROC
            '
        End If
        '
    Next
    '
    'si on arrive ici, c'est qu'aucun des points n'est visible, donc on quitte
    Exit Sub
    '
AFF_PROC:
    '
    'on dessine l'image
    D3DDevice.SetTexture 0, MatrixTex_Blanc
    '
    'on spécifie le format de point qu'on utilise
    D3DDevice.SetVertexShader Lit_FVF
    '
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertMatrix3D(0), Len(vertMatrix3D(0))
    '
    'on incrémente de 1 le nombre d'objets affichés
    NbrObjetsAff = NbrObjetsAff + 1
    '
End Sub
'
'procédure permettan d'affiche une image
Private Sub AfficherCar2(NuméroFonte As Long, Optional CouleurSz As Long = vbWhite)
    '
    Dim CharX As Integer
    Dim CharY As Integer
    Dim cTmp As Long
    '
    'on calcule les coordonnées du caractère
    cTmp = (NuméroFonte - 1) \ 10
    CharY = 58 * cTmp
    cTmp = (NuméroFonte - 1) Mod 10
    CharX = 50 * cTmp
    '
    'on détermine la couleur (on ajoute la valeur alpha)
    CouleurSz = Hex2Long(Hex(OpaciteSz2) & Hex(CouleurSz))
    '
    'on génère les points
    vertMatrix3D(0) = CreateLitVertex(XSz, YSz, ZSz, CouleurSz, 0, (1 / 512) * CharX, (1 / 512) * CharY)
    vertMatrix3D(1) = CreateLitVertex(XSz + LargeurLettreSz, YSz, ZSz, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), (1 / 512) * CharY)
    vertMatrix3D(2) = CreateLitVertex(XSz, YSz + HauteurLettreSz, ZSz, CouleurSz, 0, (1 / 512) * CharX, ((1 / 512) * CharY) + ((1 / 512) * 58))
    vertMatrix3D(3) = CreateLitVertex(XSz + LargeurLettreSz, YSz + HauteurLettreSz, ZSz, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), ((1 / 512) * CharY) + ((1 / 512) * 58))
    '
    'on vérifie qu'un des points sera visible
    For i = 0 To 3
        '
        If TestVis(MakeVector(vertMatrix3D(i).X, vertMatrix3D(i).Y, vertMatrix3D(i).Z)) = True Then
            '
            'on va à l'affichage
            GoTo AFF_PROC
            '
        End If
        '
    Next
    '
    'si on arrive ici, c'est qu'aucun des points n'est visible, donc on quitte
    Exit Sub
    '
AFF_PROC:
    '
    'on dessine l'image
    D3DDevice.SetTexture 0, MatrixTex_Vert
    '
    'on spécifie le format de point qu'on utilise
    D3DDevice.SetVertexShader Lit_FVF
    '
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertMatrix3D(0), Len(vertMatrix3D(0))
    '
    'on incrémente de 1 le nombre d'objets affichés
    NbrObjetsAff = NbrObjetsAff + 1
    '
End Sub
'
'procédure qui va calculer la position et l'opacité du caractère dans le temps
Public Sub Calculer(TotSz As Long, opDist As Long) 'TotSz permet de compter le nombre de caractères qui ont fini leur vie, opDist est l'opacité que le caractère a en fonction de sa distance par rapport à la caméra (en %)
    '
    'dMsg3 = dMsg3 & Etat & ";"
    '
    'si l'état de la lettre = 0, on quitte la procédure (il n'y a rien à afficher)
    If Etat = 0 Then Exit Sub
    '
    'on effectue les opérations en fontion de l'état de la lettre
    Select Case Etat
        '
        Case 1 'on augmente l'opacité du caractère blanc
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 >= tTps1 Then
                '
                Etat = 2
                Tps1 = 0
                tTps1 = DecalSz / 2 '150 + Rnd() * 50
                '
                'on ajuste le pourcentage en fonction de la distance
                prctSz = 100 '(100/100)*opdist
                '
            Else
                '
                'on calcule le pourcentage en fonction de la distance
                prctSz = (Tps1 / tTps1) * 100 '(Tps1 / tTps1) * 100
                '
            End If
            '
            OpaciteSz = ((2.55 * prctSz) / 100) * opDist
            '
        Case 2 'on augmente l'opacité du caractère vert tout en diminuant celle du blanc
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 >= tTps1 Then
                '
                Etat = 3
                Tps1 = 0
                tTps1 = DecalSz '+ Rnd() * 150
                '
                'on ajuste le pourcentage en fonction de la distance
                prctSz = 100 '(100/100)*opdist
                '
            Else
                '
                'on calcule le pourcentage en fonction de la distance
                prctSz = (Tps1 / tTps1) * 100 '(Tps1 / tTps1) * 100
                '
            End If
            '
            OpaciteSz = ((255 - (2.55 * prctSz)) / 100) * opDist
            OpaciteSz2 = ((2.55 * prctSz) / 100) * opDist
            '
        Case 3 'on attend
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 >= tTps1 Then
                '
                Etat = 4
                Tps1 = 0
                tTps1 = 150 + Rnd() * 200
                '
            End If
            '
            'on calcule le pourcentage en fonction de la distance
            OpaciteSz2 = 2.55 * opDist
            '
        Case 4 'diminue l'opacité du caractère vert
            '
            'on incrément la variable Tps1 en fonction de la vitesse spécifiée
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 >= tTps1 Then
                '
                Etat = 0
                Tps1 = 0
                tTps1 = 0
                '
                'on ajoute notre caractère au nombre de caractère qui ont fini leur vie
                TotSz = TotSz + 1
                '
                'on ajuste le pourcentage en fonction de la distance
                prctSz = (100 / 100) * opDist
                '
            Else
                '
                'on calcule le pourcentage en fonction de la distance
                prctSz = (Tps1 / tTps1) * 100 '(Tps1 / tTps1) * 100
                '
            End If
            '
            OpaciteSz2 = ((255 - (2.55 * prctSz)) / 100) * opDist
            '
        '
    End Select
    '
    'on incrément la variable Tps2 en fonction de la vitesse spécifiée
    Tps2 = Tps2 + (1 * VitesseGenSz)
    '
    'on vérifie qu'on ne va pas trop loint
    If Tps2 >= tTps2 Then
        '
        NumLettreSz = Rnd() * 56
        Tps2 = 0
        tTps2 = 100 + Rnd() * 1000
        '
    End If
    '
    '
End Sub
'
'procédure qui va afficher la lettre selon son état
Public Sub AfficherLettre()
    '
    'si l'état de la lettre = 0, on quitte la procédure (il n'y a rien à afficher)
    If Etat = 0 Then Exit Sub
    '
    'on effectue les opérations en fontion de l'état de la lettre
    Select Case Etat
        '
        Case 1
            '
            'on affiche le caractère blanc
            AfficherCar NumLettreSz
            '
        Case 2
            '
            'on affiche le caractère blanc
            AfficherCar NumLettreSz
            '
            'on affiche le caractère vert
            AfficherCar2 NumLettreSz
            '
        Case 3
            '
            'on affiche le caractère vert
            AfficherCar2 NumLettreSz
            '
        Case 4
            '
            'on affiche le caractère vert
            AfficherCar2 NumLettreSz
            '
        '
    End Select
    '
End Sub
'
Private Sub Class_Initialize()
    '
    Etat = 1
    OpaciteSz = 255
    NumLettreSz = Rnd() * 56
    tTps1 = 0 'Rnd() * 30
    DecalSz = 50 + Rnd() * 300
    tTps2 = 100 + Rnd() * 1000
    '
End Sub
'
'fonction vérifiant si l'objet que l'on veut afficher sera visible (s'il ne l'est pas, on n'affiche rien)
Private Function TestVis(Point3D As D3DVECTOR) As Boolean
    '
    'on déclare les variables nécessaires à la vérification
    Dim vRet As D3DVECTOR
    Dim VP As D3DVIEWPORT8
    '
    'on récupère les données
    D3DDevice.GetViewport VP
    D3DXVec3Project vRet, Point3D, VP, matProj, matView, matWorld
    '
    'on vérifie les données
    If vRet.X < VP.X Then
        TestVis = False
        Exit Function
        
    ElseIf vRet.X > VP.Width Then
        TestVis = False
        Exit Function
        
    ElseIf vRet.Y < VP.Y Then
        TestVis = False
        Exit Function
        
    ElseIf vRet.Y > VP.Height Then
        TestVis = False
        Exit Function
        
    ElseIf vRet.Z > VP.MaxZ Then
        TestVis = False
        Exit Function
    
    ElseIf vRet.Z < VP.MinZ Then
        TestVis = False
        Exit Function
        
    Else
        TestVis = True
        Exit Function
        
    End If
    
End Function

