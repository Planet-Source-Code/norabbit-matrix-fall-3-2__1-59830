Attribute VB_Name = "ModPrincipal"
'*************************************************************************
'*                                                                       *
'* MODULE DX8                                                            *
'*                                                                       *
'* DECLARATION DES VARIABLES PRINCIPALES + FONCTIONS                     *
'*                                                                       *
'* traduit et complété par Thomas John (thomas.john@swing.be)            *
'*                                                                       *
'* source : http://216.5.163.53/DirectX4VB (DirectX 4 VB, Jack Hoxley)   *
'*                                                                       *
'*************************************************************************
'
'l'objet principal
Public Dx As DirectX8
'
'cet objet contrôle tout ce qui est 3D
Public D3D As Direct3D8
'
'cet objet représente le "hardware" (la carte graphique) utilisé pour le rendu
Public D3DDevice As Direct3DDevice8
'
'une "librairie d'aide"
'D3DX8 est une classe d'aide qui contient une multitude de fonctions destinées à faciliter la programmation en DX8
Public D3DX As D3DX8
'
'variable servant à détecter si le programme tourne ou pas
Public bRunning As Boolean
'
'description des différents types de vertex
Public Const FVF_TLVERTEX = (D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR)
Public Const Lit_FVF = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1)
'
'cette structure représente un vertex 2D (identique à la structure "D3DTLVERTEX" pour Dx7)
Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tU As Single
    tV As Single
End Type
'
'structure d'un point (3D)
Public Type LITVERTEX
    X As Single
    Y As Single
    Z As Single
    Color As Long
    Specular As Long
    tU As Single
    tV As Single
End Type
'
'structure vertex pour chaque "sprite" 3D
Public vertMatrix3D(3) As LITVERTEX
'
'structure vertex pour chaque "sprite" 2D
Public vertMatrix(3) As TLVERTEX
'
'fonte
Public MainFont As D3DXFont
Public MainFontDesc As IFont
Public TextRect As RECT
Public fnt As New StdFont
'
'Pi
Public Const pi As Single = 3.14159265358979
'
Public matWorld As D3DMATRIX '//How the vertices are positioned
'où la caméra se trouve et vers où pointe-t-elle
Public matView As D3DMATRIX
'comment la caméra projecte le monde 3D sur un écran (surface) 2D
Public matProj As D3DMATRIX
'
'coordonnées de la caméra
Public CamVX As Single
Public CamVY As Single
Public CamVZ As Single
Public CamTX As Single
Public CamTY As Single
Public CamTZ As Single
Public CamDistance As Single
'
'calcul du fps (images par seconde)
Public FPS_NbrFps As Long
Public FPS_NbrImg As Long
Public lFpsTmp As Long
'
'dimensions de l'affichage
Public DimH As Long
Public DimL As Long
'
'texture de fonte matrix
Public MatrixTex_Blanc As Direct3DTexture8
Public MatrixTex_Blanc2 As Direct3DTexture8
Public MatrixTex_Blanc3 As Direct3DTexture8
Public MatrixTex_Vert As Direct3DTexture8
Public MatrixTex_Trainee As Direct3DTexture8
Public MatrixTex_Normal As Direct3DTexture8
Public MatrixTex_Effets As Direct3DTexture8
'
'pause
Public PauseSz As Boolean
'
'affichage du nombre d'images par seconde
Public AffFps As Boolean
'
'permet de ralentir la vitesse d'affichage à un certain nombre d'images par seconde
Public FpsMod As Long
'
'liste contenant toutes les coordonnées X utilisées par les lignes
Public ListeCooX As New Collection
'
'dimensions des lettres
Public HauteurLettreSz As Long
Public LargeurLettreSz As Long
'
'mode d'affichage
Public ModeAffSzX As Long
Public ModeAffSzY As Long
Public ModeBit As Long
'
'accélération matérielle ou software
Public AccMatSoft As String
'
'carte choisie
Public CarteChoixSz As Long
'
'cycle de rendu en millisecondes (temps d'attente entre chaque rendu --> diminue le nombre d'image par seconde et donc la vitesse)
Public CycleRenduSz As Long
'
'limite de lignes à afficher
Public LimiteLignesAffSz As Long
'
'nombre de lignes chargées
Public LimiteLignesChargeSz As Long
'
'classe se chargeant d'afficher les titres #2 (test)
Public cTitre2 As New ClsTitre2
'
'vitesse générale (multiplicateur)
Public VitesseGenSz As Long
'
'permet d'afficher des messages (debug seulement)
Public dMsg As String
Public dMsg2 As String
Public dMsg3 As String
'
'classe gérant une lettre 3D
'Public CL3D As New ClsLettre3D1
'
'nombre d'objets affichés
Public NbrObjetsAff As Long
'
'nombre de lignes affichées
Public NbrLignesAff As Long
'
'classe qui gère la caméra
Public Camera As New ClsCamera
'
'liste contenant toutes les coordonnées Z utilisées par les lignes
'cette liste va nous permettre d'organiser l'affichage des lignes en fonction de leur coordonnée Z
Public ListeIndexClasses As New ClsCoordonnees
'
'un classe gérant une ligne 3D
Public CLgn3D() As ClsLigne3D1
'
'liste des commandes
Public CmdSz As New Collection
'
'liste des valeurs des commandes
Public ValCmdSz As New Collection
'
'liste des pauses des commandes
Public PauseCmdSz As New Collection
'
'pointeur indiquant la commande courante à exécuter
Public CourCmdSz As Integer
'
'liste des textes à afficher (titre #1)
Public TxtTitre1 As New Collection
'
'liste des textes à afficher (titre #2)
Public TxtTitre2 As New Collection
'
'classe gérant les évènements
Public cEvnmt As New ClsEvenements
'
'pause avant l'affichage de la première ligne
Public PauseAffPremLgn As Long
'
'variables d'incrémentation
Dim i As Integer
Dim i2 As Integer
Dim iTmp As Integer
'
'classe se chargeant d'afficher les caractères spéciaux (katanas et autres...)
'Public cL1 As New ClsLettre1
'
'
'
'*******************************************************************
'* Initialise
'*******************************************************************
'
Public Function Initialise(FrmObjet As Form, DimLargeur As Long, DimHauteur As Long) As Boolean
    '
    On Error Resume Next
    '
    'décrit notre mode d'affichage
    Dim DispMode As D3DDISPLAYMODE
    '
    'décrit notre mode de vue
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    '
    'pour les filtreurs de texture
    Dim Caps As D3DCAPS8 '//For Texture Filters
    '
    'on crée notre objet principal
    Set Dx = New DirectX8
    '
    'on crée l'interface Direct3D via notre objet principal
    Set D3D = Dx.Direct3DCreate()
    '
    'on crée notre librairie d'aide
    Set D3DX = New D3DX8
    '
    'DispMode.Format = D3DFMT_X8R8G8B8
    'DispMode.Format = D3DFMT_A8R8G8B8
    DispMode.Format = D3DFMT_R5G6B5 'si ce mode ne fonctionne pas, utilisez celui juste au-dessus
    DispMode.Width = DimLargeur
    DispMode.Height = DimHauteur
    '
    DimL = DimLargeur
    DimH = DimHauteur
    '
    D3DWindow.BackBufferCount = 1 '1 BackBuffer
    D3DWindow.BackBufferWidth = DispMode.Width
    D3DWindow.BackBufferHeight = DispMode.Height
    D3DWindow.hDeviceWindow = FrmObjet.hWnd
    D3DWindow.EnableAutoDepthStencil = 1
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    D3DWindow.BackBufferFormat = DispMode.Format
    '
    If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D32) = D3D_OK Then
        '
        'on peut utiliser un Z-Buffer de 32-bit
        D3DWindow.AutoDepthStencilFormat = D3DFMT_D32
        '
    Else
        '
        If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D24) = D3D_OK Then
            '
            'on peut utiliser un Z-Buffer de 24-bit
            D3DWindow.AutoDepthStencilFormat = D3DFMT_D24
            '
        Else
            '
            If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D16) = D3D_OK Then
                '
                'on peut utiliser un Z-Buffer de 16-bit
                D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
                '
            End If
            '
        End If
        '
    End If
    '
    'on montre notre feuille pour être sûr
    FrmObjet.Show
    '
    'on la met au-dessus de toutes
    SetWindowPos FrmObjet.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    '
    'cette ligne crée un "device" qui utilise la carte graphique ("hardware") pour effectuer les calculs si possible,
    'ou le processeur ("software") et utilise comme objet de réception notre feuille principale
    'on lance le mode hardware ou software selon les options chargées
    Select Case AccMatSoft
        '
        Case "REF"
            '
            Set D3DDevice = D3D.CreateDevice(CarteChoixSz, D3DDEVTYPE_REF, FrmObjet.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
            '
        Case "HAL"
            '
            Set D3DDevice = D3D.CreateDevice(CarteChoixSz, D3DDEVTYPE_HAL, FrmObjet.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)
            '
        '
    End Select
    '
    'nos points (vertices) n'ont pas besoin de lumière, donc on désactive cette option
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    '
    'ceci dit à dx d'afficher les triangles même s'ils ne sont pas en face de la caméra
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    '
    'déclarations utiles pour le rendu de textures transparantes
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    '
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
    'Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTA_TFACTOR)
    'Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_TFACTOR)
    '
    'filtrage de texture : donne un meilleur résultat lors d'un redimensionnement d'une texture
    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    '
    'on active notre Z-Buffer
    D3DDevice.SetRenderState D3DRS_ZENABLE, 1
    '
    'la matrice "World"
    D3DXMatrixIdentity matWorld
    D3DDevice.SetTransform D3DTS_WORLD, matWorld
    '
    'la matrice "view" (en gros, la caméra)
    D3DXMatrixLookAtLH matView, MakeVector(CamVX, CamVY, CamVZ), MakeVector(CamTX, CamTY, CamTZ), MakeVector(0, -1, 0)
    D3DDevice.SetTransform D3DTS_VIEW, matView
    '
    'la matrice de projection (angle, distance de rendu)
    D3DXMatrixPerspectiveFovLH matProj, pi / 2, 1, 0.1, 100000
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj
    '
    '
    'initialisation du rendu du texte
    fnt.Name = "Arial"
    fnt.Size = 10
    fnt.Bold = True
    Set MainFontDesc = fnt
    Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)
    TextRect.Top = 1
    TextRect.Left = 1
    TextRect.bottom = DimH
    TextRect.Right = DimL
    '
    '**************************************
    '** chargement des textures          **
    '**************************************
    '
    Set MatrixTex_Blanc = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\matrixfall\fontes_blanches.png", 512, 512, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    Set MatrixTex_Blanc2 = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\matrixfall\fontes_blanches2.png", 512, 512, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    Set MatrixTex_Blanc3 = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\matrixfall\fontes_blanches3.png", 512, 512, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    Set MatrixTex_Vert = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\matrixfall\fontes_vertes.png", 512, 512, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    Set MatrixTex_Normal = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\matrixfall\fontes_normales.png", 512, 512, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    Set MatrixTex_Effets = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\matrixfall\effets.png", 512, 512, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    '
    '**************************************
    '** fin du chargement des textures   **
    '**************************************
    '
    'on initialise le nombre de lignes chargées (aucune donc -1 car les tableaux commencent toujours par 0)
    LimiteLignesChargeSz = -1
    '
    'on initialise le tableau des lignes 3d
    ReDim CLgn3D(0 To 0)
    '
    'on ajoute quelques commandes (affichage de textes)
    'CmdSz.Add "txt"
    'ValCmdSz.Add "matrix fall 3"
    'PauseCmdSz.Add "10000"
    '
    CmdSz.Add "txt"
    ValCmdSz.Add "realise par_THOMAS JOHN"
    PauseCmdSz.Add "10000"
    '
    CmdSz.Add "txt"
    ValCmdSz.Add "email_THOMAS.JOHN@OPEN-DESIGN.BE"
    PauseCmdSz.Add "1"
    '
    CmdSz.Add "txt"
    ValCmdSz.Add "website_WWW.OPEN-DESIGN.BE"
    PauseCmdSz.Add "1"
    '
    CmdSz.Add "noop"
    ValCmdSz.Add ""
    PauseCmdSz.Add "20000"
    '
    'on spécifie le temps de la première pause de la classe gérant les évènements
    cEvnmt.tTps2 = CLng(PauseCmdSz.Item(CourCmdSz))
    '
    'on cache le curseur
    ShowCursor False
    '
    'si on arrive jusqu'ici, c'est que tout s'est bien passé
    Initialise = True
    '
    'on gère les erreurs survenues durant l'initialisation ici
    If Err Then
        '
        EcrireLog "### " & Time & " ###" & vbCrLf & "erreur lors de l'initialisation :" & vbCrLf & D3DX.GetErrorString(Err.Number) & vbCrLf
        '
    End If
    '
    Dim RotateAngle As Single
    Dim matTemp As D3DMATRIX 'contient des données temporaires
    '
    bRunning = True
    '
    '-1 pour éviter la division par zéro
    lFpsTmp = GetTickCount - 1
    '
    Do While bRunning
        '
        'on vérifie si il y a une pause
        If PauseSz = False Then
            '
            'nombre d'images par secondes
            FPS_NbrFps = FPS_NbrImg / ((GetTickCount - lFpsTmp) / 1000)
            '
            'on vérifie que la variable qui contient le nombre d'images rendues ne soit pas trop grande
            If FPS_NbrImg > 1000000 Then
                '
                FPS_NbrImg = 0
                '
                lFpsTmp = GetTickCount - 1
                '
            End If
            '
            'on rend la scène si le nombre d'images par seconde a été atteint
            If FPS_NbrFps <= CycleRenduSz Then
                '
                'on incrément le nombre d'images rendues
                FPS_NbrImg = FPS_NbrImg + 1
                '
                'on met à jour la position de la caméra
                Camera.Affichage
                '
                'on modifie la matrice "view"
                D3DXMatrixLookAtLH matView, MakeVector(CamVX, CamVY, CamVZ), MakeVector(CamTX, CamTY, CamTZ), MakeVector(0, -1, 0)
                D3DDevice.SetTransform D3DTS_VIEW, matView
                '
                '*****************************************************************************************
                'on "rend" (dessine) la scène
                '*****************************************************************************************
                '
                'on efface la surface et on lui donne la couleur noir
                D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
                '
                dMsg3 = ""
                '
                'on commence le rendu
                D3DDevice.BeginScene
                    '
                    'tous les appels de rendu doivent être fait entre "BeginScene" et "EndScene"
                    '
                    'évènements
                    cEvnmt.Calcul
                    '
                    'on calcule les lignes en 3D de test
                    i = 0
                    Do While i <= LimiteLignesChargeSz
                        '
                        CLgn3D(i).Calculer
                        '
                        i = i + 1
                        '
                    Loop
                    '
                    'on ajoute tous les index de chaque classe dans la liste vide
                    ListeIndexClasses.ToutEffacer
                    '
                    i = 0
                    Do While i <= LimiteLignesChargeSz 'LimiteLignesAffSz
                        '
                        ListeIndexClasses.Ajout i
                        '
                        i = i + 1
                        '
                    Loop
                    '
                    'on affiche les lignes en fonction de la position de la caméra
                    'il faut savoir que si les lignes ne sont pas affichées dans le bon ordre
                    'la transparence des caractères n'est pas gérée.
                    'il faut, par exemple, que la ligne se trouvant le plus près de la caméra
                    'soit affichée après celle qui se trouve le plus loint de la caméra
                    'vous pouvez voir la différence en inversant la condition ci-dessous --> CamVZ - CamTZ < 0
                    If CamVZ - CamTZ > 0 Then
                        '
                        Do While ListeIndexClasses.TotalElements > 0
                            '
                            iTmp = 1
                            '
                            For i = 1 To ListeIndexClasses.TotalElements
                                '
                                If CLgn3D(ListeIndexClasses.ElementIndex(i)).ZSz < CLgn3D(ListeIndexClasses.ElementIndex(iTmp)).ZSz Then iTmp = i
                                '
                            Next
                            '
                            'on affiche cet élément
                            CLgn3D(ListeIndexClasses.ElementIndex(iTmp)).Afficher
                            '
                            'on enlève cet élément de la liste
                            ListeIndexClasses.Enlever iTmp
                            '
                        Loop
                        '
                    Else
                        '
                        Do While ListeIndexClasses.TotalElements > 0
                            '
                            iTmp = 1
                            '
                            For i = 1 To ListeIndexClasses.TotalElements
                                '
                                If CLgn3D(ListeIndexClasses.ElementIndex(i)).ZSz > CLgn3D(ListeIndexClasses.ElementIndex(iTmp)).ZSz Then iTmp = i
                                '
                            Next
                            '
                            'on affiche cet élément
                            CLgn3D(ListeIndexClasses.ElementIndex(iTmp)).Afficher
                            '
                            'on enlève cet élément de la liste
                            ListeIndexClasses.Enlever iTmp
                            '
                        Loop
                        '
                    End If
                    '
                    'on affiche le titre de test
                    cTitre2.AfficherTitre
                    '
                    'on affiche le caractère de test
                    'cL1.AfficherCar
                    '
                    dMsg = "Z1 caméra : " & CamVZ & vbCrLf
                    dMsg = dMsg & "Z2 caméra : " & CamTZ & vbCrLf
                    dMsg = dMsg & "pourcentage 1 : " & Camera.prctSz & vbCrLf
                    dMsg = dMsg & "pourcentage 2 : " & Camera.prctSz2 & vbCrLf
                    dMsg = dMsg & "a et b : " & Camera.aTmp & " - " & Camera.bTmp & vbCrLf
                    dMsg = dMsg & "lignes chargées: " & LimiteLignesChargeSz + 1 & vbCrLf
                    dMsg = dMsg & "max lignes : " & LimiteLignesAffSz & vbCrLf
                    '
                    'i = 0
                    'Do While i <= LimiteLignesAffSz
                        '
                        'dMsg = dMsg & i & " : " & CLgn3D(i).ZSz & " ; " & CLgn3D(i).NbrCarFin & "/" & CLgn3D(i).NbrCarSz & " ; " & CLgn3D(i).tTps1 & vbCrLf
                        '
                        'i = i + 1
                        '
                    'Loop
                    '
                    'rendu du texte
                    If AffFps = True Then D3DX.DrawText MainFont, &HFFB30000, FPS_NbrFps & " fps (" & Time & ")" & vbCrLf & dMsg3 & vbCrLf & dMsg & vbCrLf & "Nombre d'objets affichés : " & NbrObjetsAff & vbCrLf & "Nombre de lignes affichées : " & NbrLignesAff & vbCrLf & dMsg2, TextRect, DT_LEFT        '
                    '
                D3DDevice.EndScene
                '
                'on met à jour l'image à l'écran
                D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
                '
                'on remet à zéro le nombre d'objets et de lignes affichés
                NbrObjetsAff = 0
                NbrLignesAff = 0
                '
                '*****************************************************************************************
                'fin du rendu
                '*****************************************************************************************
                '
            End If
            '
        End If
        '
        'on laisse vb respirer
        DoEvents
        '
    Loop
    '
    'on gère les erreurs survenues lors du rendu ici
    If Err Then
        '
        EcrireLog "### " & Time & " ###" & vbCrLf & "erreur lors du rendu :" & vbCrLf & D3DX.GetErrorString(Err.Number) & vbCrLf
        '
    End If
    '
    'on affiche le curseur
    ShowCursor True
    '
    'la boucle s'est terminée signifiant que le programme va se fermer
    'il faut avant tout décharger les objets qu'on a chargé précédemment
    '
    On Error Resume Next 'pour être sûr
    '
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set Dx = Nothing
    Set D3DX = Nothing
    '
    'on gère les erreurs survenues lors du déchargement des objets ici
    If Err Then
        '
        EcrireLog "### " & Time & " ###" & vbCrLf & "erreur lors du déchargement des objets :" & vbCrLf & D3DX.GetErrorString(Err.Number) & vbCrLf
        '
    End If
    '
End Function
'
'fonction permettant de créer un vecteur en une ligne
Public Function MakeVector(X As Single, Y As Single, Z As Single) As D3DVECTOR
    '
    MakeVector.X = X
    MakeVector.Y = Y
    MakeVector.Z = Z
    '
End Function
'
'fontion permettant de remplir un objet en une seule ligne
Public Function CreateTLVertex(X As Single, Y As Single, Z As Single, rhw As Single, Color As Long, Specular As Long, tU As Single, tV As Single) As TLVERTEX
    '
    '//NB: whilst you can pass floating point values for the coordinates (single)
    '       there is little point - Direct3D will just approximate the coordinate by rounding
    '       which may well produce unwanted results....
    '
    With CreateTLVertex
        '
        .X = X
        .Y = Y
        .Z = Z
        .rhw = rhw
        .Color = Color
        .Specular = Specular
        .tU = tU
        .tV = tV
        '
    End With
    '
End Function
'
'fontion permettant de remplir un objet en une seule ligne
Public Function CreateLitVertex(X As Single, Y As Single, Z As Single, Color As Long, Specular As Long, tU As Single, tV As Single) As LITVERTEX
    '
    With CreateLitVertex
        '
        .X = X
        .Y = Y
        .Z = Z
        .Color = Color
        .Specular = Specular
        .tU = tU
        .tV = tV
        '
    End With
    '
End Function
'
'convertit une donnée hex en long
Public Function Hex2Long(hHex) As Long
    '
    Hex2Long = "&H" & hHex
    '
End Function
'
'lecture des options
Public Sub ChargerOptions()
    '
    On Error Resume Next
    '
    Dim FichSz As Integer
    Dim sTmp As String
    Dim sTmp2() As String
    Dim sTmp3() As String
    '
    FichSz = FreeFile
    '
    Open App.Path & "\matrixfall\matrixfall.ini" For Binary As #FichSz
    '
    sTmp = Space(LOF(FichSz))
    '
    Get #FichSz, , sTmp
    '
    Close #FichSz
    '
    'on récupère chaque information en séparant celles-ci par "VbCrLf"
    sTmp2() = Split(sTmp, vbCrLf)
    '
    'si le tableau ne contient aucune donnée, on va directement à la fin de la procédure
    If UBound(sTmp2) = -1 Then GoTo FIN_PROC
    '
    'je réutilise "FichSz" comme variable d'incrémentation
    For FichSz = 0 To UBound(sTmp2)
        '
        If sTmp2(FichSz) = vbNullString Then GoTo FIN_PROC
        '
        'on récupère les infos en fonction de leur nom
        Select Case Left$(sTmp2(FichSz), 4)
            '
            Case "vtgn" 'vitesse générale
                '
                VitesseGenSz = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                '
            Case "accm" 'type d'accélération (hardware ou software)
                '
                AccMatSoft = Right(sTmp2(FichSz), 3)
                '
            Case "mode" 'mode d'affichage
                '
                sTmp3() = Split(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7), "x")
                '
                'on vérifie que toutes les infos sont là
                If UBound(sTmp3) = 2 Then
                    '
                    ModeAffSzX = CLng(sTmp3(0))
                    ModeAffSzY = CLng(sTmp3(1))
                    ModeBit = CLng(sTmp3(2))
                    '
                Else
                    '
                    'sinon, on met le mode par défaut
                    ModeAffSzX = 800
                    ModeAffSzY = 600
                    ModeBit = 16
                    '
                End If
                '
            Case "haut" 'hauteur des lettres
                '
                HauteurLettreSz = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                '
            Case "larg" 'largeur des lettres
                '
                LargeurLettreSz = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                '
            Case "cycl" 'nombre d'images par seconde à rendre
                '
                CycleRenduSz = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                '
            Case "cart" 'carte vidéo
                '
                CarteChoixSz = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                '
            Case "liml" 'nombre de lignes à afficher
                '
                LimiteLignesAffSz = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                '
            Case "plng" 'temps d'attente avant l'affichage de la première ligne
                '
                cEvnmt.tTps1 = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                PauseAffPremLgn = cEvnmt.tTps1
                '
            Case "cmds" 'commande
                '
                sTmp3() = Split(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7), ";")
                '
                'on vérifie que toutes les infos sont là sinon on ne fait rien
                If UBound(sTmp3) = 2 Then
                    '
                    CmdSz.Add sTmp3(0)
                    ValCmdSz.Add sTmp3(1)
                    PauseCmdSz.Add sTmp3(2)
                    '
                End If
                '
            '
        End Select
        '
    Next
    '
FIN_PROC:
    '
    'on vérifie que les informations importantes sont présentes
    If HauteurLettreSz <= 0 Then
        '
        HauteurLettreSz = 21
        '
    End If
    '
    If LargeurLettreSz <= 0 Then
        '
        LargeurLettreSz = (HauteurLettreSz / 21) * 18
        '
    End If
    '
    If ModeAffSzX < 800 Or ModeAffSzY < 600 Or ModeBit = 0 Then
        '
        ModeAffSzX = 800
        ModeAffSzY = 600
        ModeBit = 16
        '
    End If
    '
    If LimiteLignesAffSz <= 0 Then
        '
        LimiteLignesAffSz = 10
        '
    End If
    '
    'on vérifie que la "vitesse générale" (multiplicateur) ne soit pas nul
    If VitesseGenSz = 0 Then VitesseGenSz = 17
    '
End Sub
'
'fonction inscrivant dans le fichier log_matrix_fall.txt les erreurs et autres
Public Sub EcrireLog(TexteSz As String)
    '
    Dim FichSz As Integer
    '
    FichSz = FreeFile
    '
    Open App.Path & "\log_matrix_fall.txt" For Binary As #FichSz
    Seek #FichSz, LOF(FichSz)
    Put #FichSz, , TexteSz & vbCrLf
    Close #FichSz
    '
End Sub
