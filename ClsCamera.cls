VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCamera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE LA CAMERA                                                    *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'état de la classe
Public Etat As Long
'
'transitions
Public Tps1 As Long
Public tTps1 As Long
'
'angle de rotation
Dim rot As Single
'
'pourcentages
Public prctSz As Single
Public prctSz2 As Single
'
'variables quelconques
Public aTmp As Single
Public bTmp As Single
'
'
'initialisation de la classe
Private Sub Class_Initialize()
    '
    Etat = 1
    tTps1 = 5500
    CamDistance = 400
    rot = 270
    '
End Sub
'
'ON EFFECTUE LES CHANGEMENTS
Public Sub Affichage()
    '
    'on effectue les changements en fonction de l'état et du temps alloué
    Select Case Etat
        '
        Case 1 'on avance
            '
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 >= tTps1 Then
                '
                tTps1 = 2000 + Rnd() * 20000
                Tps1 = 0
                Etat = 1 + Rnd() * 4
                '
                '%
                'prctSz = 100
                '
            Else
                '
                'calcul du %
                'prctSz = (Tps1 / tTps1) * 100
                '
            End If
            '
            CamTZ = CamTZ + (0.2 * VitesseGenSz)
            '
        Case 2 'on effectue une rotation dont le centre est le "point de mire" de la caméra" (CamTZ) de sens +
            '
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 >= (tTps1 / 2) Then
                '
                tTps1 = 1500 + Rnd() * 4000
                Tps1 = 0
                Etat = 1
                '
                '%
                prctSz = 100
                prctSz2 = 100
                '
            Else
                '
                'calcul du %
                'aTmp = Tps1 - (tTps1 / 2)
                'bTmp = Sin((aTmp / tTps1) * (3.14 * 2) + (3.14 / 2))
                'prctSz = ((1 + bTmp) / 2) * 100
                prctSz = ((1 + (Sin(((Tps1 - (tTps1 / 2)) / tTps1) * (3.14 * 2) + (3.14 / 2)))) / 2) * 100
                'prctSz2 = ((1 + (Sin(((Tps1 - (tTps1 / 2)) / (tTps1 / 2)) * (3.14 * 2) + (3.14 / 2)))) / 2) * 100
                '
            End If
            '
            rot = 270 + (3.6 * prctSz)
            '
            'on continue à faire avancer la caméra tout en accélérant la vitesse petit à petit, puis en décélérant la vitesse petit à petit
            'CamTZ = CamTZ + ((0.2 + (0.002 * (100 - prctSz2))) * VitesseGenSz)
            '
            'on continue à faire avancer la caméra à la même vitesse
            CamTZ = CamTZ + (0.2 * VitesseGenSz)
            '
            If rot >= 360 Then
                '
                rot = rot - 360
                '
            End If
            '
        Case 3 'on effectue une rotation dont le centre est le "point de mire" de la caméra" (CamTZ) de sens -
            '
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 >= (tTps1 / 2) Then
                '
                tTps1 = 1500 + Rnd() * 4000
                Tps1 = 0
                Etat = 1
                '
                '%
                prctSz = 100
                prctSz2 = 100
                '
            Else
                '
                'calcul du %
                'aTmp = Tps1 - (tTps1 / 2)
                'bTmp = Sin((aTmp / tTps1) * (3.14 * 2) + (3.14 / 2))
                'prctSz = ((1 + bTmp) / 2) * 100
                prctSz = ((1 + (Sin(((Tps1 - (tTps1 / 2)) / tTps1) * (3.14 * 2) + (3.14 / 2)))) / 2) * 100
                prctSz2 = ((1 + (Sin(((Tps1 - (tTps1 / 2)) / (tTps1 / 2)) * (3.14 * 2) + (3.14 / 2)))) / 2) * 100
                '
            End If
            '
            rot = 270 - (3.6 * prctSz)
            '
            'on continue à faire avancer la caméra tout en accélérant la vitesse petit à petit, puis en décélérant la vitesse petit à petit
            'CamTZ = CamTZ + ((0.2 + (0.002 * (100 - prctSz2))) * VitesseGenSz)
            '
            'on continue à faire avancer la caméra à la même vitesse
            CamTZ = CamTZ + (0.2 * VitesseGenSz)
            '
            If rot <= 0 Then
                '
                rot = rot + 360
                '
            End If
            '
        Case 4 'on effectue une rotation dont le centre est le "point de mire" de la caméra" (CamTZ) de sens +, et on fait monter la caméra
            '
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 >= (tTps1 / 2) Then
                '
                tTps1 = 1500 + Rnd() * 4000
                Tps1 = 0
                Etat = 1
                '
                '%
                prctSz = 100
                prctSz2 = 100
                '
            Else
                '
                'calcul du %
                'aTmp = Tps1 - (tTps1 / 2)
                'bTmp = Sin((aTmp / tTps1) * (3.14 * 2) + (3.14 / 2))
                prctSz = ((1 + (Sin(((Tps1 - (tTps1 / 2)) / tTps1) * (3.14 * 2) + (3.14 / 2)))) / 2) * 100
                prctSz2 = ((1 + (Sin(((Tps1 - (tTps1 / 2)) / (tTps1 / 2)) * (3.14 * 2) + (3.14 / 2)))) / 2) * 100
                '
            End If
            '
            rot = 270 + (3.6 * prctSz)
            '
            'on continue à faire avancer la caméra tout en accélérant la vitesse petit à petit, puis en décélérant la vitesse petit à petit
            CamTZ = CamTZ + ((0.2 + (0.004 * (100 - prctSz2))) * VitesseGenSz)
            '
            CamVY = -(2 * (100 - prctSz2))
            '
            If rot >= 360 Then
                '
                rot = rot - 360
                '
            End If
            '
        Case 5 'on effectue une rotation dont le centre est le "point de mire" de la caméra" (CamTZ) de sens -, et on fait descendre la caméra
            '
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 >= (tTps1 / 2) Then
                '
                tTps1 = 1500 + Rnd() * 4000
                Tps1 = 0
                Etat = 1
                '
                '%
                prctSz = 100
                prctSz2 = 100
                '
            Else
                '
                'calcul du %
                'aTmp = Tps1 - (tTps1 / 2)
                'bTmp = Sin((aTmp / tTps1) * (3.14 * 2) + (3.14 / 2))
                prctSz = ((1 + (Sin(((Tps1 - (tTps1 / 2)) / tTps1) * (3.14 * 2) + (3.14 / 2)))) / 2) * 100
                prctSz2 = ((1 + (Sin(((Tps1 - (tTps1 / 2)) / (tTps1 / 2)) * (3.14 * 2) + (3.14 / 2)))) / 2) * 100
                '
            End If
            '
            rot = 270 - (3.6 * prctSz)
            '
            'on continue à faire avancer la caméra tout en accélérant la vitesse petit à petit, puis en décélérant la vitesse petit à petit
            CamTZ = CamTZ + ((0.2 + (0.004 * (100 - prctSz2))) * VitesseGenSz)
            '
            CamVY = -(2 * (100 - prctSz2))
            '
            If rot >= 360 Then
                '
                rot = rot - 360
                '
            End If
            '
        '
    End Select
    '
    CamVX = Cos((rot / 360) * 2 * 3.141) * CamDistance + CamTX
    CamVZ = Sin((rot / 360) * 2 * 3.141) * CamDistance + CamTZ
    '
End Sub
