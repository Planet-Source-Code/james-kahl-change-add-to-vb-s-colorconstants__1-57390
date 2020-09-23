VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblColor 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum ColorConstants
    'standard vb color constants
    vbBlack = &H0&
    vbRed = &HFF&
    vbGreen = &HEE00&
    vbYellow = &HEEEE&
    vbBlue = &HFF0000
    vbMagenta = &HFF00FF
    vbCyan = &HFFFF00
    vbWhite = &HFFFFFF
    'standard vb system colors
    vbScrollBars = &H80000000
    vbDesktop = &H80000001
    vbActiveTitleBar = &H80000002
    vbInactiveTitleBar = &H80000003
    vbMenuBar = &H80000004
    vbWindowBackground = &H80000005
    vbWindowFrame = &H80000006
    vbMenuText = &H80000007
    vbWindowText = &H80000008
    vbTitleBarText = &H80000009
    vbActiveTitleBarText = vbTitleBarText
    vbActiveBorder = &H8000000A
    vbInactiveBorder = &H8000000B
    vbApplicationWorkspace = &H8000000C
    vbHighlight = &H8000000D
    vbHighlightText = &H8000000E
    vbButtonFace = &H8000000F
    vb3DFace = vbButtonFace
    vb3DShadow = &H80000010
    vbButtonShadow = vb3DShadow
    vbGrayText = &H80000011
    vbButtonText = &H80000012
    vbInactiveCaptionText = &H80000013
    vbInactiveTitleBarText = vbInactiveCaptionText
    vb3DHighlight = &H80000014
    vb3DDKShadow = &H80000015
    vb3DLight = &H80000016
    vbInfoText = &H80000017
    vbInfoBackground = &H80000018
    'custom color constants
    AliceBlue = &HFFF8F0
    AntiqueWhite = &HD7EBFA
    Aquamarine = &HD4FF7F
    Azure = &HFFFFF0
    Beige = &HDCF5F5
    Bisque = &HC4E4FF
    BlanchedAlmond = &HCDEBFF
    BlueViolet = &HE22B8A
    Brown = &H2A2AA5
    BurlyWood = &H87B8DE
    CadetBlue = &HA09E5F
    Chartreuse = &HFF7F&
    Chocolate = &H1E69D2
    Coral = &H507FFF
    CornflowerBlue = &HED9564
    Cornsilk = &HDCF8FF
    Crimson = &H3C14DC
    DarkBlue = &H8B0000
    DarkCyan = &H8B8B00
    DarkGoldenrod = &HB86B8
    DarkGray = &HA9A9A9
    DarkGreen = &H6400
    DarkKhaki = &H6BB7BD
    DarkMagenta = &H8B008B
    DarkOliveGreen = &H2F6B55
    DarkOrange = &H8CFF&
    DarkOrchid = &HCC3299
    DarkRed = &H8B&
    DarkSalmon = &H7A96E9
    DarkSeaGreen = &H8FBC8F
    DarkSlateBlue = &H8B3D48
    DarkSlateGray = &H4F4F2F
    DarkTurquoise = &HD1CE00
    DarkViolet = &HD30094
    DeepPink = &H9314FF
    DeepSkyBlue = &HFFBF00
    DimGray = &H696969
    DodgerBlue = &HFF901E
    FireBrick = &H2222B2
    FloralWhite = &HF0FAFF
    ForestGreen = &H228B22
    Gainsboro = &HDCDCDC
    GhostWhite = &HFFF8F8
    Gold = &HD7FF&
    Goldenrod = &H20A5DA
    Gray = &H808080
    Green = &H8000&
    GreenYellow = &H2FFFAD
    Honeydew = &HF0FFF0
    HotPink = &HB469FF
    IndianRed = &H5C5CCD
    Indigo = &H82004B
    Ivory = &HF0FFFF
    Khaki = &H8CE6F0
    Lavender = &HFAE6E6
    LavenderBlush = &HF5F0FF
    LawnGreen = &HFC7C&
    LemonChiffon = &HCDFAFF
    LightBlue = &HE6D8AD
    LightCoral = &H8080F0
    LightCyan = &HFFFFE0
    LightGoldenrodYellow = &HD2FAFA
    LightGreen = &H90EE90
    LightGray = &HD3D3D3
    LightPink = &HC1B6FF
    LightSalmon = &H7AA0FF
    LightSeaGreen = &HAAB220
    LightSkyBlue = &HFACE87
    LightSlateGray = &H998877
    LightSteelBlue = &HDEC4B0
    LightYellow = &HE0FFFF
    Lime = &HFF00&
    LimeGreen = &H32CD32
    Linen = &HE6F0FA
    Maroon = &H80&
    MediumAquamarine = &HAACD66
    MediumBlue = &HCD0000
    MediumOrchid = &HD355BA
    MediumPurple = &HDB7093
    MediumSeaGreen = &H71B33C
    MediumSlateBlue = &HEE687B
    MediumSpringGreen = &H9AFA00
    MediumTurquoise = &HCCD148
    MediumVioletRed = &H8515C7
    MidnightBlue = &H701919
    MintCream = &HFAFFF5
    MistyRose = &HE1E4FF
    Moccasin = &HB5E4FF
    NavajoWhite = &HADDEFF
    Navy = &H800000
    OldLace = &HE6F5FD
    Olive = &H8080&
    OliveDrab = &H238E6B
    Orange = &HA5FF&
    OrangeRed = &H45FF&
    Orchid = &HD670DA
    PaleGoldenrod = &HAAE8EE
    PaleGreen = &H98FB98
    PaleTurquoise = &HEEEEAF
    PaleVioletRed = &H9370DB
    PapayaWhip = &HD5EFFF
    PeachPuff = &HB9DAFF
    Peru = &H3F85CD
    Pink = &HCBC0FF
    Plum = &HDDA0DD
    PowderBlue = &HE6E0B0
    Purple = &H800080
    RosyBrown = &H8F8FBC
    RoyalBlue = &HE16941
    SaddleBrown = &H13458B
    Salmon = &H7280FA
    SandyBrown = &H60A4F4
    SeaGreen = &H578B2E
    Seashell = &HEEF5FF
    Sienna = &H2D52A0
    Silver = &HC0C0C0
    SkyBlue = &HEBCE87
    SlateBlue = &HCD5A6A
    SlateGray = &H908070
    Snow = &HFAFAFF
    SpringGreen = &H7FFF00
    SteelBlue = &HB48246
    TanColor = &H8CB4D2
    Teal = &H808000
    Thistle = &HD8BFD8
    Tomato = &H4763FF
    Turquoise = &HD0E040
    Violet = &HEE82EE
    Wheat = &HB3DEF5
    WhiteSmoke = &HF5F5F5
    YellowGreen = &H32CD9A
    OnanGreen = &H609000
    #If False Then
        OnanGreen
        AliceBlue
        AntiqueWhite
        Aquamarine
        Azure
        Beige
        Bisque
        BlanchedAlmond
        BlueViolet
        Brown
        BurlyWood
        CadetBlue
        Chartreuse
        Chocolate
        Coral
        CornflowerBlue
        Cornsilk
        Crimson
        DarkBlue
        DarkCyan
        DarkGoldenrod
        DarkGray
        DarkGreen
        DarkKhaki
        DarkMagenta
        DarkOliveGreen
        DarkOrange
        DarkOrchid
        DarkRed
        DarkSalmon
        DarkSeaGreen
        DarkSlateBlue
        DarkSlateGray
        DarkTurquoise
        DarkViolet
        DeepPink
        DeepSkyBlue
        DimGray
        DodgerBlue
        FireBrick
        FloralWhite
        ForestGreen
        Gainsboro
        GhostWhite
        Gold
        Goldenrod
        Gray
        Green
        GreenYellow
        Honeydew
        HotPink
        IndianRed
        Indigo
        Ivory
        Khaki
        Lavender
        LavenderBlush
        LawnGreen
        LemonChiffon
        LightBlue
        LightCoral
        LightCyan
        LightGoldenrodYellow
        LightGreen
        LightGray
        LightPink
        LightSalmon
        LightSeaGreen
        LightSkyBlue
        LightSlateGray
        LightSteelBlue
        LightYellow
        Lime
        LimeGreen
        Linen
        Maroon
        MediumAquamarine
        MediumBlue
        MediumOrchid
        MediumPurple
        MediumSeaGreen
        MediumSlateBlue
        MediumSpringGreen
        MediumTurquoise
        MediumVioletRed
        MidnightBlue
        MintCream
        MistyRose
        Moccasin
        NavajoWhite
        Navy
        OldLace
        Olive
        OliveDrab
        Orange
        OrangeRed
        Orchid
        PaleGoldenrod
        PaleGreen
        PaleTurquoise
        PaleVioletRed
        PapayaWhip
        PeachPuff
        Peru
        Pink
        Plum
        PowderBlue
        Purple
        RosyBrown
        RoyalBlue
        SaddleBrown
        Salmon
        SandyBrown
        SeaGreen
        Seashell
        Sienna
        Silver
        SkyBlue
        SlateBlue
        SlateGray
        Snow
        SpringGreen
        SteelBlue
        TanColor
        Teal
        Thistle
        Tomato
        Turquoise
        Violet
        Wheat
        WhiteSmoke
        YellowGreen
    #End If
End Enum

Private Sub ShowColors()
    'show the system colors
    lblColor(1).BackColor = vbScrollBars
    lblColor(2).BackColor = vbDesktop
    lblColor(3).BackColor = vbActiveTitleBar
    lblColor(4).BackColor = vbInactiveTitleBar
    lblColor(5).BackColor = vbMenuBar
    lblColor(6).BackColor = vbWindowBackground
    lblColor(7).BackColor = vbWindowFrame
    lblColor(7).ForeColor = vbButtonFace
    lblColor(8).BackColor = vbMenuText
    lblColor(8).ForeColor = vbButtonFace
    lblColor(9).BackColor = vbWindowText
    lblColor(9).ForeColor = vbButtonFace
    lblColor(10).BackColor = vbTitleBarText
    lblColor(11).BackColor = vbActiveBorder
    lblColor(12).BackColor = vbInactiveBorder
    lblColor(13).BackColor = vbApplicationWorkspace
    lblColor(14).BackColor = vbHighlight
    lblColor(15).BackColor = vbHighlightText
    lblColor(16).BackColor = vbButtonFace
    lblColor(17).BackColor = vb3DFace
    lblColor(18).BackColor = vb3DShadow
    lblColor(19).BackColor = vbButtonShadow
    lblColor(20).BackColor = vbInactiveCaptionText
    lblColor(21).BackColor = vbInactiveTitleBarText
    lblColor(22).BackColor = vb3DHighlight
    lblColor(23).BackColor = vb3DDKShadow
    lblColor(24).BackColor = vb3DLight
    lblColor(25).BackColor = vbInfoText
    lblColor(25).ForeColor = vbButtonFace
    lblColor(26).BackColor = vbInfoBackground
    'show vb colors
    lblColor(27).BackColor = vbBlack
    lblColor(27).ForeColor = vbWhite
    lblColor(28).BackColor = vbWhite
    lblColor(29).BackColor = vbRed
    lblColor(30).BackColor = vbGreen
    lblColor(31).BackColor = vbBlue
    lblColor(32).BackColor = vbMagenta
    lblColor(33).BackColor = vbYellow
    lblColor(34).BackColor = vbCyan
    'show custom colors
    lblColor(35).BackColor = AliceBlue
    lblColor(36).BackColor = AntiqueWhite
    lblColor(37).BackColor = Aquamarine
    lblColor(38).BackColor = Azure
    lblColor(39).BackColor = Beige
    lblColor(40).BackColor = BlanchedAlmond
    lblColor(41).BackColor = BlueViolet
    lblColor(42).BackColor = Brown
    lblColor(43).BackColor = BurlyWood
    lblColor(44).BackColor = CadetBlue
    lblColor(45).BackColor = Chartreuse
    lblColor(46).BackColor = Chocolate
    lblColor(47).BackColor = Coral
    lblColor(48).BackColor = CornflowerBlue
    lblColor(49).BackColor = Cornsilk
    lblColor(50).BackColor = Crimson
    lblColor(51).BackColor = DarkBlue
    lblColor(52).BackColor = DarkCyan
    lblColor(53).BackColor = DarkGoldenrod
    lblColor(54).BackColor = DarkGray
    lblColor(55).BackColor = DarkGreen
    lblColor(56).BackColor = DarkKhaki
    lblColor(57).BackColor = DarkMagenta
    lblColor(58).BackColor = DarkOliveGreen
    lblColor(59).BackColor = DarkOrange
    lblColor(60).BackColor = DarkOrchid
    lblColor(61).BackColor = DarkRed
    lblColor(62).BackColor = DarkSalmon
    lblColor(63).BackColor = DarkSeaGreen
    lblColor(64).BackColor = DarkSlateBlue
    lblColor(65).BackColor = DarkSlateGray
    lblColor(66).BackColor = DarkTurquoise
    lblColor(67).BackColor = DarkViolet
    lblColor(68).BackColor = DeepPink
    lblColor(69).BackColor = DeepSkyBlue
    lblColor(70).BackColor = DimGray
    lblColor(71).BackColor = DodgerBlue
    lblColor(72).BackColor = FireBrick
    lblColor(73).BackColor = FloralWhite
    lblColor(74).BackColor = ForestGreen
    lblColor(75).BackColor = Gainsboro
    lblColor(76).BackColor = GhostWhite
    lblColor(77).BackColor = Gold
    lblColor(78).BackColor = Goldenrod
    lblColor(79).BackColor = Gray
    lblColor(80).BackColor = Green
    lblColor(81).BackColor = GreenYellow
    lblColor(82).BackColor = Honeydew
    lblColor(83).BackColor = HotPink
    lblColor(84).BackColor = IndianRed
    lblColor(85).BackColor = Indigo
    lblColor(86).BackColor = Ivory
    lblColor(87).BackColor = Khaki
    lblColor(88).BackColor = Lavender
    lblColor(89).BackColor = LavenderBlush
    lblColor(90).BackColor = LawnGreen
    lblColor(91).BackColor = LemonChiffon
    lblColor(92).BackColor = LightBlue
    lblColor(93).BackColor = LightCoral
    lblColor(94).BackColor = LightCyan
    lblColor(95).BackColor = LightGoldenrodYellow
    lblColor(96).BackColor = LightGreen
    lblColor(97).BackColor = LightGray
    lblColor(98).BackColor = LightPink
    lblColor(99).BackColor = LightSalmon
    lblColor(100).BackColor = LightSeaGreen
    lblColor(101).BackColor = LightSkyBlue
    lblColor(102).BackColor = LightSlateGray
    lblColor(103).BackColor = LightSteelBlue
    lblColor(104).BackColor = LightYellow
    lblColor(105).BackColor = Lime
    lblColor(106).BackColor = LimeGreen
    lblColor(107).BackColor = Linen
    lblColor(108).BackColor = Maroon
    lblColor(109).BackColor = MediumAquamarine
    lblColor(110).BackColor = MediumBlue
    lblColor(111).BackColor = MediumOrchid
    lblColor(112).BackColor = MediumPurple
    lblColor(113).BackColor = MediumSeaGreen
    lblColor(114).BackColor = MediumSlateBlue
    lblColor(115).BackColor = MediumSpringGreen
    lblColor(116).BackColor = MediumTurquoise
    lblColor(117).BackColor = MediumVioletRed
    lblColor(118).BackColor = MidnightBlue
    lblColor(119).BackColor = MintCream
    lblColor(120).BackColor = MistyRose
    lblColor(121).BackColor = Moccasin
    lblColor(122).BackColor = NavajoWhite
    lblColor(123).BackColor = Navy
    lblColor(124).BackColor = OldLace
    lblColor(125).BackColor = Olive
    lblColor(126).BackColor = OliveDrab
    lblColor(127).BackColor = Orange
    lblColor(128).BackColor = OrangeRed
    lblColor(129).BackColor = Orchid
    lblColor(130).BackColor = PaleGoldenrod
    lblColor(131).BackColor = PaleGreen
    lblColor(132).BackColor = PaleTurquoise
    lblColor(133).BackColor = PaleVioletRed
    lblColor(134).BackColor = PapayaWhip
    lblColor(135).BackColor = Peru
    lblColor(136).BackColor = Pink
    lblColor(137).BackColor = Plum
    lblColor(138).BackColor = PowderBlue
    lblColor(139).BackColor = Purple
    lblColor(140).BackColor = RosyBrown
    lblColor(141).BackColor = RoyalBlue
    lblColor(142).BackColor = SaddleBrown
    lblColor(143).BackColor = Salmon
    lblColor(144).BackColor = SandyBrown
    lblColor(145).BackColor = SeaGreen
    lblColor(146).BackColor = Seashell
    lblColor(147).BackColor = Sienna
    lblColor(148).BackColor = Silver
    lblColor(149).BackColor = SkyBlue
    lblColor(150).BackColor = SlateBlue
    lblColor(151).BackColor = SlateGray
    lblColor(152).BackColor = Snow
    lblColor(153).BackColor = SpringGreen
    lblColor(154).BackColor = SteelBlue
    lblColor(155).BackColor = TanColor
    lblColor(156).BackColor = Teal
    lblColor(157).BackColor = Thistle
    lblColor(158).BackColor = Tomato
    lblColor(159).BackColor = Turquoise
    lblColor(160).BackColor = Violet
    lblColor(161).BackColor = Wheat
    lblColor(162).BackColor = WhiteSmoke
    lblColor(163).BackColor = YellowGreen
    lblColor(164).BackColor = OnanGreen
End Sub

Private Sub showCaptions()
    'show the system colors
    lblColor(1).Caption = "vbScrollBars"
    lblColor(2).Caption = "vbDesktop"
    lblColor(3).Caption = "vbActiveTitleBar"
    lblColor(4).Caption = "vbInactiveTitleBar"
    lblColor(5).Caption = "vbMenuBar"
    lblColor(6).Caption = "vbWindowBackground"
    lblColor(7).Caption = "vbWindowFrame"
    lblColor(8).Caption = "vbMenuText"
    lblColor(9).Caption = "vbWindowText"
    lblColor(10).Caption = "vbTitleBarText"
    lblColor(11).Caption = "vbActiveBorder"
    lblColor(12).Caption = "vbInactiveBorder"
    lblColor(13).Caption = "vbApplicationWorkspace"
    lblColor(14).Caption = "vbHighlight"
    lblColor(15).Caption = "vbHighlightText"
    lblColor(16).Caption = "vbButtonFace"
    lblColor(17).Caption = "vb3DFace"
    lblColor(18).Caption = "vb3DShadow"
    lblColor(19).Caption = "vbButtonShadow"
    lblColor(20).Caption = "vbInactiveCaptionText"
    lblColor(21).Caption = "vbInactiveTitleBarText"
    lblColor(22).Caption = "vb3DHighlight"
    lblColor(23).Caption = "vb3DDKShadow"
    lblColor(24).Caption = "vb3DLight"
    lblColor(25).Caption = "vbInfoText"
    lblColor(26).Caption = "vbInfoBackground"
    'show vb colors
    lblColor(27).Caption = "vbBlack"
    lblColor(28).Caption = "vbWhite"
    lblColor(29).Caption = "vbRed"
    lblColor(30).Caption = "vbGreen"
    lblColor(31).Caption = "vbBlue"
    lblColor(32).Caption = "vbMagenta"
    lblColor(33).Caption = "vbYellow"
    lblColor(34).Caption = "vbCyan"
    'show custom colors
    lblColor(35).Caption = "AliceBlue"
    lblColor(36).Caption = "AntiqueWhite"
    lblColor(37).Caption = "Aquamarine"
    lblColor(38).Caption = "Azure"
    lblColor(39).Caption = "Beige"
    lblColor(40).Caption = "BlanchedAlmond"
    lblColor(41).Caption = "BlueViolet"
    lblColor(42).Caption = "Brown"
    lblColor(43).Caption = "BurlyWood"
    lblColor(44).Caption = "CadetBlue"
    lblColor(45).Caption = "Chartreuse"
    lblColor(46).Caption = "Chocolate"
    lblColor(47).Caption = "Coral"
    lblColor(48).Caption = "CornflowerBlue"
    lblColor(49).Caption = "Cornsilk"
    lblColor(50).Caption = "Crimson"
    lblColor(51).Caption = "DarkBlue"
    lblColor(52).Caption = "DarkCyan"
    lblColor(53).Caption = "DarkGoldenrod"
    lblColor(54).Caption = "DarkGray"
    lblColor(55).Caption = "DarkGreen"
    lblColor(56).Caption = "DarkKhaki"
    lblColor(57).Caption = "DarkMagenta"
    lblColor(58).Caption = "DarkOliveGreen"
    lblColor(59).Caption = "DarkOrange"
    lblColor(60).Caption = "DarkOrchid"
    lblColor(61).Caption = "DarkRed"
    lblColor(62).Caption = "DarkSalmon"
    lblColor(63).Caption = "DarkSeaGreen"
    lblColor(64).Caption = "DarkSlateBlue"
    lblColor(65).Caption = "DarkSlateGray"
    lblColor(66).Caption = "DarkTurquoise"
    lblColor(67).Caption = "DarkViolet"
    lblColor(68).Caption = "DeepPink"
    lblColor(69).Caption = "DeepSkyBlue"
    lblColor(70).Caption = "DimGray"
    lblColor(71).Caption = "DodgerBlue"
    lblColor(72).Caption = "FireBrick"
    lblColor(73).Caption = "FloralWhite"
    lblColor(74).Caption = "ForestGreen"
    lblColor(75).Caption = "Gainsboro"
    lblColor(76).Caption = "GhostWhite"
    lblColor(77).Caption = "Gold"
    lblColor(78).Caption = "Goldenrod"
    lblColor(79).Caption = "Gray"
    lblColor(80).Caption = "Green"
    lblColor(81).Caption = "GreenYellow"
    lblColor(82).Caption = "Honeydew"
    lblColor(83).Caption = "HotPink"
    lblColor(84).Caption = "IndianRed"
    lblColor(85).Caption = "Indigo"
    lblColor(86).Caption = "Ivory"
    lblColor(87).Caption = "Khaki"
    lblColor(88).Caption = "Lavender"
    lblColor(89).Caption = "LavenderBlush"
    lblColor(90).Caption = "LawnGreen"
    lblColor(91).Caption = "LemonChiffon"
    lblColor(92).Caption = "LightBlue"
    lblColor(93).Caption = "LightCoral"
    lblColor(94).Caption = "LightCyan"
    lblColor(95).Caption = "LightGoldenrodYellow"
    lblColor(96).Caption = "LightGreen"
    lblColor(97).Caption = "LightGray"
    lblColor(98).Caption = "LightPink"
    lblColor(99).Caption = "LightSalmon"
    lblColor(100).Caption = "LightSeaGreen"
    lblColor(101).Caption = "LightSkyBlue"
    lblColor(102).Caption = "LightSlateGray"
    lblColor(103).Caption = "LightSteelBlue"
    lblColor(104).Caption = "LightYellow"
    lblColor(105).Caption = "Lime"
    lblColor(106).Caption = "LimeGreen"
    lblColor(107).Caption = "Linen"
    lblColor(108).Caption = "Maroon"
    lblColor(109).Caption = "MediumAquamarine"
    lblColor(110).Caption = "MediumBlue"
    lblColor(111).Caption = "MediumOrchid"
    lblColor(112).Caption = "MediumPurple"
    lblColor(113).Caption = "MediumSeaGreen"
    lblColor(114).Caption = "MediumSlateBlue"
    lblColor(115).Caption = "MediumSpringGreen"
    lblColor(116).Caption = "MediumTurquoise"
    lblColor(117).Caption = "MediumVioletRed"
    lblColor(118).Caption = "MidnightBlue"
    lblColor(119).Caption = "MintCream"
    lblColor(120).Caption = "MistyRose"
    lblColor(121).Caption = "Moccasin"
    lblColor(122).Caption = "NavajoWhite"
    lblColor(123).Caption = "Navy"
    lblColor(124).Caption = "OldLace"
    lblColor(125).Caption = "Olive"
    lblColor(126).Caption = "OliveDrab"
    lblColor(127).Caption = "Orange"
    lblColor(128).Caption = "OrangeRed"
    lblColor(129).Caption = "Orchid"
    lblColor(130).Caption = "PaleGoldenrod"
    lblColor(131).Caption = "PaleGreen"
    lblColor(132).Caption = "PaleTurquoise"
    lblColor(133).Caption = "PaleVioletRed"
    lblColor(134).Caption = "PapayaWhip"
    lblColor(135).Caption = "Peru"
    lblColor(136).Caption = "Pink"
    lblColor(137).Caption = "Plum"
    lblColor(138).Caption = "PowderBlue"
    lblColor(139).Caption = "Purple"
    lblColor(140).Caption = "RosyBrown"
    lblColor(141).Caption = "RoyalBlue"
    lblColor(142).Caption = "SaddleBrown"
    lblColor(143).Caption = "Salmon"
    lblColor(144).Caption = "SandyBrown"
    lblColor(145).Caption = "SeaGreen"
    lblColor(146).Caption = "Seashell"
    lblColor(147).Caption = "Sienna"
    lblColor(148).Caption = "Silver"
    lblColor(149).Caption = "SkyBlue"
    lblColor(150).Caption = "SlateBlue"
    lblColor(151).Caption = "SlateGray"
    lblColor(152).Caption = "Snow"
    lblColor(153).Caption = "SpringGreen"
    lblColor(154).Caption = "SteelBlue"
    lblColor(155).Caption = "TanColor"
    lblColor(156).Caption = "Teal"
    lblColor(157).Caption = "Thistle"
    lblColor(158).Caption = "Tomato"
    lblColor(159).Caption = "Turquoise"
    lblColor(160).Caption = "Violet"
    lblColor(161).Caption = "Wheat"
    lblColor(162).Caption = "WhiteSmoke"
    lblColor(163).Caption = "YellowGreen"
    lblColor(164).Caption = "OnanGreen"
End Sub

Private Sub Form_Load()
    Dim lIdx As Long
    Dim ctl(2 To 170) As Control
    
    Const XOFF As Long = 2000
    Const YOFF As Long = 240
    
    On Error Resume Next
    
    For lIdx = 2 To 170
        Set ctl(lIdx) = lblColor(lIdx)
        Load ctl(lIdx)
        With lblColor(1)
            Select Case lIdx
                Case 2 To 34
                    ctl(lIdx).Left = .Left
                    ctl(lIdx).Top = .Top + (YOFF * (lIdx - 1))
                Case 35 To 68
                    ctl(lIdx).Left = .Left + XOFF
                    ctl(lIdx).Top = .Top + (YOFF * (lIdx - 35))
                Case 69 To 102
                    ctl(lIdx).Left = .Left + (XOFF * 2)
                    ctl(lIdx).Top = .Top + (YOFF * (lIdx - 69))
                Case 103 To 136
                    ctl(lIdx).Left = .Left + (XOFF * 3)
                    ctl(lIdx).Top = .Top + (YOFF * (lIdx - 103))
                Case 137 To 170
                    ctl(lIdx).Left = .Left + (XOFF * 4)
                    ctl(lIdx).Top = .Top + (YOFF * (lIdx - 137))
                Case Else
            End Select
        End With
        ctl(lIdx).Visible = True
    Next lIdx
    On Error GoTo 0
    ShowColors
    showCaptions
End Sub

