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
'opacit�s des caract�res
Public OpaciteSz As Long
Public OpaciteSz2 As Long
'
'd�calage du temps des transitions
Public DecalSz As Long
'
'num�ro du caract�re
Public NumLettreSz As Long
'
'coordonn�es du caract�re
Public XSz As Single
Public YSz As Single
Public ZSz As Single
'
'couleur du caract�re
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
'le caract�re peut il changer dans le temps
Public ChangementSz As Long
'
'variations de couleur
Public VarCouleur As Long
'
'�tat du caract�re
Public Etat As Long
'
'variable d'incr�mentation
Dim i As Integer
'
'
'proc�dure permettan d'affiche une image
Private Sub AfficherCar(Num�roFonte As Long, Optional CouleurSz As Long = vbWhite)
    '
    Dim CharX As Integer
    Dim CharY As Integer
    Dim cTmp As Long
    '
    'on calcule les coordonn�es du caract�re
    cTmp = (Num�roFonte - 1) \ 10
    CharY = 58 * cTmp
    cTmp = (Num�roFonte - 1) Mod 10
    CharX = 50 * cTmp
    '
    'on d�termine la couleur (on ajoute la valeur alpha)
    CouleurSz = Hex2Long(Hex(OpaciteSz) & Hex(CouleurSz))
    '
    'on g�n�re les points
    vertMatrix3D(0) = CreateLitVertex(XSz, YSz, ZSz, CouleurSz, 0, (1 / 512) * CharX, (1 / 512) * CharY)
    vertMatrix3D(1) = CreateLitVertex(XSz + LargeurLettreSz, YSz, ZSz, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), (1 / 512) * CharY)
    vertMatrix3D(2) = CreateLitVertex(XSz, YSz + HauteurLettreSz, ZSz, CouleurSz, 0, (1 / 512) * CharX, ((1 / 512) * CharY) + ((1 / 512) * 58))
    vertMatrix3D(3) = CreateLitVertex(XSz + LargeurLettreSz, YSz + HauteurLettreSz, ZSz, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), ((1 / 512) * CharY) + ((1 / 512) * 58))
    '
    'on v�rifie qu'un des points sera visible
    For i = 0 To 3
        '
        If TestVis(MakeVector(vertMatrix3D(i).X, vertMatrix3D(i).Y, vertMatrix3D(i).Z)) = True Then
            '
            'on va � l'affichage
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
    'on sp�cifie le format de point qu'on utilise
    D3DDevice.SetVertexShader Lit_FVF
    '
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertMatrix3D(0), Len(vertMatrix3D(0))
    '
    'on incr�mente de 1 le nombre d'objets affich�s
    NbrObjetsAff = NbrObjetsAff + 1
    '
End Sub
'
'proc�dure permettan d'affiche une image
Private Sub AfficherCar2(Num�roFonte As Long, Optional CouleurSz As Long = vbWhite)
    '
    Dim CharX As Integer
    Dim CharY As Integer
    Dim cTmp As Long
    '
    'on calcule les coordonn�es du caract�re
    cTmp = (Num�roFonte - 1) \ 10
    CharY = 58 * cTmp
    cTmp = (Num�roFonte - 1) Mod 10
    CharX = 50 * cTmp
    '
    'on d�termine la couleur (on ajoute la valeur alpha)
    CouleurSz = Hex2Long(Hex(OpaciteSz2) & Hex(CouleurSz))
    '
    'on g�n�re les points
    vertMatrix3D(0) = CreateLitVertex(XSz, YSz, ZSz, CouleurSz, 0, (1 / 512) * CharX, (1 / 512) * CharY)
    vertMatrix3D(1) = CreateLitVertex(XSz + LargeurLettreSz, YSz, ZSz, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), (1 / 512) * CharY)
    vertMatrix3D(2) = CreateLitVertex(XSz, YSz + HauteurLettreSz, ZSz, CouleurSz, 0, (1 / 512) * CharX, ((1 / 512) * CharY) + ((1 / 512) * 58))
    vertMatrix3D(3) = CreateLitVertex(XSz + LargeurLettreSz, YSz + HauteurLettreSz, ZSz, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), ((1 / 512) * CharY) + ((1 / 512) * 58))
    '
    'on v�rifie qu'un des points sera visible
    For i = 0 To 3
        '
        If TestVis(MakeVector(vertMatrix3D(i).X, vertMatrix3D(i).Y, vertMatrix3D(i).Z)) = True Then
            '
            'on va � l'affichage
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
    'on sp�cifie le format de point qu'on utilise
    D3DDevice.SetVertexShader Lit_FVF
    '
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertMatrix3D(0), Len(vertMatrix3D(0))
    '
    'on incr�mente de 1 le nombre d'objets affich�s
    NbrObjetsAff = NbrObjetsAff + 1
    '
End Sub
'
'proc�dure qui va calculer la position et l'opacit� du caract�re dans le temps
Public Sub Calculer(TotSz As Long, opDist As Long) 'TotSz permet de compter le nombre de caract�res qui ont fini leur vie, opDist est l'opacit� que le caract�re a en fonction de sa distance par rapport � la cam�ra (en %)
    '
    'dMsg3 = dMsg3 & Etat & ";"
    '
    'si l'�tat de la lettre = 0, on quitte la proc�dure (il n'y a rien � afficher)
    If Etat = 0 Then Exit Sub
    '
    'on effectue les op�rations en fontion de l'�tat de la lettre
    Select Case Etat
        '
        Case 1 'on augmente l'opacit� du caract�re blanc
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
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
        Case 2 'on augmente l'opacit� du caract�re vert tout en diminuant celle du blanc
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
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
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
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
        Case 4 'diminue l'opacit� du caract�re vert
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
            If Tps1 >= tTps1 Then
                '
                Etat = 0
                Tps1 = 0
                tTps1 = 0
                '
                'on ajoute notre caract�re au nombre de caract�re qui ont fini leur vie
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
    'on incr�ment la variable Tps2 en fonction de la vitesse sp�cifi�e
    Tps2 = Tps2 + (1 * VitesseGenSz)
    '
    'on v�rifie qu'on ne va pas trop loint
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
'proc�dure qui va afficher la lettre selon son �tat
Public Sub AfficherLettre()
    '
    'si l'�tat de la lettre = 0, on quitte la proc�dure (il n'y a rien � afficher)
    If Etat = 0 Then Exit Sub
    '
    'on effectue les op�rations en fontion de l'�tat de la lettre
    Select Case Etat
        '
        Case 1
            '
            'on affiche le caract�re blanc
            AfficherCar NumLettreSz
            '
        Case 2
            '
            'on affiche le caract�re blanc
            AfficherCar NumLettreSz
            '
            'on affiche le caract�re vert
            AfficherCar2 NumLettreSz
            '
        Case 3
            '
            'on affiche le caract�re vert
            AfficherCar2 NumLettreSz
            '
        Case 4
            '
            'on affiche le caract�re vert
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
'fonction v�rifiant si l'objet que l'on veut afficher sera visible (s'il ne l'est pas, on n'affiche rien)
Private Function TestVis(Point3D As D3DVECTOR) As Boolean
    '
    'on d�clare les variables n�cessaires � la v�rification
    Dim vRet As D3DVECTOR
    Dim VP As D3DVIEWPORT8
    '
    'on r�cup�re les donn�es
    D3DDevice.GetViewport VP
    D3DXVec3Project vRet, Point3D, VP, matProj, matView, matWorld
    '
    'on v�rifie les donn�es
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

