VERSION 5.00
Begin VB.Form fra_main 
   BackColor       =   &H0074FFEA&
   BorderStyle     =   0  'None
   Caption         =   "Active commands"
   ClientHeight    =   11835
   ClientLeft      =   255
   ClientTop       =   600
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "fra_main.frx":0000
   ScaleHeight     =   11835
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmr_activkey 
      Interval        =   50
      Left            =   6780
      Top             =   780
   End
   Begin VB.Timer tmr_lastreco 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   8175
      Top             =   780
   End
   Begin VB.TextBox txt_lastreco 
      Appearance      =   0  'Flat
      BackColor       =   &H00656136&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CCCCCC&
      Height          =   300
      Left            =   1455
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   8010
      Width           =   2625
   End
   Begin VB.CheckBox chk_noactions 
      Appearance      =   0  'Flat
      BackColor       =   &H004B422B&
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1410
      TabIndex        =   24
      Top             =   7530
      Width           =   240
   End
   Begin VB.Frame fra_center 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      ForeColor       =   &H80000008&
      Height          =   9915
      Left            =   4650
      TabIndex        =   41
      Top             =   1665
      Width           =   6615
      Begin VB.Frame fra_options 
         BackColor       =   &H0098E167&
         BorderStyle     =   0  'None
         Height          =   3030
         Left            =   3720
         TabIndex        =   45
         Top             =   5970
         Visible         =   0   'False
         Width           =   2505
         Begin VB.OptionButton opt_activation 
            Appearance      =   0  'Flat
            BackColor       =   &H0096DE66&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   27
            Top             =   420
            Width           =   225
         End
         Begin VB.OptionButton opt_activation 
            Appearance      =   0  'Flat
            BackColor       =   &H0096DE66&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   28
            Top             =   675
            Width           =   225
         End
         Begin VB.OptionButton opt_activation 
            Appearance      =   0  'Flat
            BackColor       =   &H0096DE66&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   29
            Top             =   930
            Width           =   225
         End
         Begin VB.OptionButton opt_activation 
            Appearance      =   0  'Flat
            BackColor       =   &H0096DE66&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   135
            TabIndex        =   30
            Top             =   1185
            Width           =   225
         End
         Begin VB.CheckBox chk_confidence 
            Appearance      =   0  'Flat
            BackColor       =   &H0096DE66&
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   165
            TabIndex        =   32
            Top             =   2190
            Width           =   240
         End
         Begin PiLfIuS.LiveButton cmd_opt_cancel 
            Height          =   420
            Left            =   2085
            TabIndex        =   35
            Top             =   2610
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   741
            Picture         =   "fra_main.frx":1856B
            PictureOver     =   "fra_main.frx":18975
            BackColor       =   9887334
         End
         Begin PiLfIuS.LiveButton cmd_opt_ok 
            Height          =   420
            Left            =   1635
            TabIndex        =   34
            Top             =   2610
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   741
            Picture         =   "fra_main.frx":18D8A
            PictureOver     =   "fra_main.frx":19067
            BackColor       =   9887334
         End
         Begin PiLfIuS.LiveButton cmd_opt_change 
            Height          =   240
            Left            =   1725
            TabIndex        =   31
            Top             =   1755
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   423
            Picture         =   "fra_main.frx":19472
            PictureOver     =   "fra_main.frx":197BD
            BackColor       =   9887334
         End
         Begin PiLfIuS.LiveButton btn_helpconfidence 
            Height          =   270
            Left            =   2175
            TabIndex        =   33
            Top             =   2190
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   476
            Picture         =   "fra_main.frx":19AB5
            PictureOver     =   "fra_main.frx":19CA6
            BackColor       =   9887334
         End
         Begin VB.Image Image6 
            Height          =   150
            Left            =   150
            Picture         =   "fra_main.frx":19E98
            Top             =   165
            Width           =   2130
         End
         Begin VB.Image img_activation 
            Height          =   150
            Index           =   1
            Left            =   390
            Picture         =   "fra_main.frx":1A1BF
            Top             =   690
            Width           =   1065
         End
         Begin VB.Image img_activation 
            Height          =   150
            Index           =   2
            Left            =   390
            Picture         =   "fra_main.frx":1A338
            Top             =   945
            Width           =   1230
         End
         Begin VB.Image img_activation 
            Height          =   150
            Index           =   3
            Left            =   390
            Picture         =   "fra_main.frx":1A4CC
            Top             =   1200
            Width           =   960
         End
         Begin VB.Image img_activation 
            Height          =   150
            Index           =   0
            Left            =   390
            Picture         =   "fra_main.frx":1A629
            Top             =   435
            Width           =   885
         End
         Begin VB.Image img_opt_activkey 
            Height          =   150
            Left            =   195
            Picture         =   "fra_main.frx":1A770
            Top             =   1545
            Width           =   1230
         End
         Begin VB.Image img_opt_key 
            Height          =   150
            Left            =   1470
            Picture         =   "fra_main.frx":1A904
            Top             =   1545
            Width           =   270
         End
         Begin VB.Label lbl_opt_key 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   240
            TabIndex        =   47
            Top             =   1785
            Width           =   45
         End
         Begin VB.Label lbl_opt_help 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   180
            Left            =   1155
            TabIndex        =   46
            Top             =   2745
            Width           =   405
         End
         Begin VB.Image img_confidence 
            Height          =   120
            Left            =   420
            Picture         =   "fra_main.frx":1A9B3
            Top             =   2235
            Width           =   1755
         End
      End
      Begin VB.Frame fra_addkey 
         BackColor       =   &H0021D1E6&
         BorderStyle     =   0  'None
         Height          =   2160
         Left            =   570
         TabIndex        =   43
         Top             =   90
         Visible         =   0   'False
         Width           =   2760
         Begin VB.ListBox lst_keystrokes 
            Appearance      =   0  'Flat
            BackColor       =   &H0021D1E6&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1290
            Left            =   300
            TabIndex        =   8
            Top             =   540
            Width           =   1980
         End
         Begin PiLfIuS.LiveButton cmd_erasekey 
            Height          =   435
            Left            =   2310
            TabIndex        =   11
            Top             =   1755
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   767
            Picture         =   "fra_main.frx":1ABB1
            PictureOver     =   "fra_main.frx":1AFF8
            BackColor       =   2216422
         End
         Begin PiLfIuS.LiveButton cmd_keydown 
            Height          =   465
            Left            =   2295
            TabIndex        =   10
            Top             =   1155
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   820
            Picture         =   "fra_main.frx":1B633
            PictureOver     =   "fra_main.frx":1BA43
            BackColor       =   2216422
         End
         Begin PiLfIuS.LiveButton cmd_keyup 
            Height          =   420
            Left            =   2295
            TabIndex        =   9
            Top             =   735
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   741
            Picture         =   "fra_main.frx":1C007
            PictureOver     =   "fra_main.frx":1C40A
            BackColor       =   2216422
         End
         Begin VB.Image Image7 
            Height          =   330
            Left            =   300
            Picture         =   "fra_main.frx":1C9B0
            Top             =   135
            Width           =   2010
         End
         Begin VB.Label lbl_key_help 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   180
            Left            =   1770
            TabIndex        =   44
            Top             =   1890
            Width           =   405
         End
      End
      Begin VB.Frame fra_addkeys 
         BackColor       =   &H0000E9F2&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4800
         Left            =   3555
         TabIndex        =   42
         Top             =   240
         Width           =   2715
         Begin VB.TextBox txt_comborepeat 
            Appearance      =   0  'Flat
            BackColor       =   &H0047F8FF&
            Height          =   300
            Left            =   795
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   3585
            Width           =   405
         End
         Begin VB.CheckBox chk_comborepeat 
            Appearance      =   0  'Flat
            BackColor       =   &H0000E9F2&
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   15
            TabIndex        =   19
            Top             =   3630
            Width           =   240
         End
         Begin VB.TextBox txt_combowait 
            Appearance      =   0  'Flat
            BackColor       =   &H0047F8FF&
            Height          =   300
            Left            =   1410
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   3135
            Width           =   840
         End
         Begin VB.CheckBox chk_combowait 
            Appearance      =   0  'Flat
            BackColor       =   &H0000E9F2&
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   15
            TabIndex        =   17
            Top             =   3195
            Width           =   240
         End
         Begin VB.TextBox txt_keycombo 
            Appearance      =   0  'Flat
            BackColor       =   &H0047F8FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   315
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1260
            Width           =   2400
         End
         Begin VB.ComboBox cmb_specialkeys 
            Appearance      =   0  'Flat
            BackColor       =   &H0047F8FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1710
            Width           =   1710
         End
         Begin VB.ComboBox cmb_mousejoy 
            Appearance      =   0  'Flat
            BackColor       =   &H0047F8FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   2160
            Width           =   1650
         End
         Begin VB.CheckBox chk_combohold 
            Appearance      =   0  'Flat
            BackColor       =   &H0000E9F2&
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   15
            TabIndex        =   15
            Top             =   2775
            Width           =   240
         End
         Begin VB.TextBox txt_combohold 
            Appearance      =   0  'Flat
            BackColor       =   &H0047F8FF&
            Height          =   300
            Left            =   870
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   2715
            Width           =   840
         End
         Begin PiLfIuS.LiveButton cmd_addkeycombo 
            Height          =   420
            Left            =   2235
            TabIndex        =   21
            Top             =   4380
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   741
            Picture         =   "fra_main.frx":1CD93
            PictureOver     =   "fra_main.frx":1D1B9
            BackColor       =   59890
         End
         Begin VB.Image Image18 
            Height          =   120
            Left            =   1275
            Picture         =   "fra_main.frx":1D5D8
            Top             =   3690
            Width           =   360
         End
         Begin VB.Image Image17 
            Height          =   135
            Left            =   285
            Picture         =   "fra_main.frx":1D694
            Top             =   3690
            Width           =   435
         End
         Begin VB.Label lbl_addkeys_help 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   180
            Left            =   1800
            TabIndex        =   48
            Top             =   4515
            Width           =   405
         End
         Begin VB.Image Image12 
            Height          =   90
            Left            =   2295
            Picture         =   "fra_main.frx":1D7A4
            Top             =   3270
            Width           =   195
         End
         Begin VB.Image Image11 
            Height          =   150
            Left            =   285
            Picture         =   "fra_main.frx":1D82C
            Top             =   3240
            Width           =   1050
         End
         Begin VB.Image Image10 
            Height          =   150
            Left            =   0
            Picture         =   "fra_main.frx":1D99C
            Top             =   1350
            Width           =   255
         End
         Begin VB.Image Image9 
            Height          =   870
            Left            =   75
            Picture         =   "fra_main.frx":1DA82
            Top             =   300
            Width           =   2670
         End
         Begin VB.Image Image8 
            Height          =   150
            Left            =   630
            Picture         =   "fra_main.frx":1E4EB
            Top             =   0
            Width           =   1905
         End
         Begin VB.Image Image14 
            Height          =   150
            Left            =   15
            Picture         =   "fra_main.frx":1E74E
            Top             =   1800
            Width           =   960
         End
         Begin VB.Image Image15 
            Height          =   150
            Left            =   15
            Picture         =   "fra_main.frx":1E8F0
            Top             =   2265
            Width           =   990
         End
         Begin VB.Image Image16 
            Height          =   120
            Left            =   300
            Picture         =   "fra_main.frx":1EAF7
            Top             =   2820
            Width           =   510
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00828282&
            X1              =   570
            X2              =   2010
            Y1              =   2595
            Y2              =   2595
         End
         Begin VB.Image Image21 
            Height          =   90
            Left            =   1755
            Picture         =   "fra_main.frx":1EBC8
            Top             =   2865
            Width           =   195
         End
      End
      Begin PiLfIuS.LiveButton img_opttext 
         Height          =   195
         Left            =   4890
         TabIndex        =   26
         Top             =   5520
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   344
         Picture         =   "fra_main.frx":1EC50
         PictureOver     =   "fra_main.frx":1F048
         BackColor       =   255
      End
      Begin VB.Image img_optback 
         Height          =   3540
         Left            =   3375
         Picture         =   "fra_main.frx":1F576
         Top             =   5700
         Width           =   3225
      End
      Begin VB.Image img_opt_button 
         Height          =   345
         Left            =   4635
         MouseIcon       =   "fra_main.frx":20EBD
         MousePointer    =   99  'Custom
         Picture         =   "fra_main.frx":2100F
         Top             =   5385
         Width           =   1605
      End
      Begin VB.Image Image20 
         Height          =   360
         Left            =   30
         Picture         =   "fra_main.frx":215C8
         Top             =   495
         Width           =   225
      End
      Begin VB.Image img_listen_on 
         Height          =   315
         Left            =   15
         Picture         =   "fra_main.frx":21798
         ToolTipText     =   "PiLfIuS! is listening"
         Top             =   105
         Width           =   315
      End
      Begin VB.Image img_listen_off 
         Height          =   315
         Left            =   15
         Picture         =   "fra_main.frx":21ABD
         ToolTipText     =   "PiLfIuS! is not listening"
         Top             =   105
         Width           =   315
      End
      Begin VB.Image Image23 
         Height          =   4740
         Left            =   0
         Picture         =   "fra_main.frx":21DD5
         Top             =   -15
         Width           =   510
      End
      Begin VB.Image img_keys 
         Height          =   2475
         Left            =   135
         Picture         =   "fra_main.frx":22817
         Top             =   75
         Width           =   3420
      End
      Begin VB.Image img_addkeys 
         Height          =   5010
         Left            =   3180
         Picture         =   "fra_main.frx":23775
         Top             =   135
         Width           =   3300
      End
      Begin VB.Image Image24 
         Height          =   6075
         Left            =   -495
         Picture         =   "fra_main.frx":24F17
         Top             =   -345
         Width           =   7260
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   5340
      TabIndex        =   40
      Top             =   7125
      Width           =   6075
      Begin VB.Image Image13 
         Height          =   1005
         Left            =   0
         Picture         =   "fra_main.frx":35A72
         Top             =   270
         Width           =   4155
      End
      Begin VB.Image img_about 
         Height          =   375
         Left            =   4230
         MouseIcon       =   "fra_main.frx":37159
         MousePointer    =   99  'Custom
         Picture         =   "fra_main.frx":372AB
         Top             =   600
         Width           =   1710
      End
      Begin VB.Line Line2 
         X1              =   4155
         X2              =   6075
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Image Image22 
         Height          =   1860
         Left            =   -120
         Picture         =   "fra_main.frx":37D5F
         Top             =   -285
         Width           =   6210
      End
   End
   Begin VB.Timer tmr_showopt 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   7245
      Top             =   780
   End
   Begin PiLfIuS.LiveButton cmd_close 
      Height          =   150
      Left            =   7155
      TabIndex        =   36
      Top             =   300
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   265
      Picture         =   "fra_main.frx":3DA0D
      BackColor       =   7929834
   End
   Begin VB.Timer tmr_showkeys 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   7710
      Top             =   780
   End
   Begin PiLfIuS.LiveButton cmd_saveas 
      Height          =   525
      Left            =   2640
      TabIndex        =   23
      Top             =   6555
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   926
      Picture         =   "fra_main.frx":3DFC1
      PictureOver     =   "fra_main.frx":3E919
      BackColor       =   7929834
   End
   Begin PiLfIuS.LiveButton cmd_save 
      Height          =   525
      Left            =   840
      TabIndex        =   22
      Top             =   6555
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   926
      Picture         =   "fra_main.frx":3F519
      PictureOver     =   "fra_main.frx":3FE75
      BackColor       =   7929834
   End
   Begin PiLfIuS.LiveButton cmd_erasecommand 
      Height          =   420
      Left            =   4050
      TabIndex        =   7
      Top             =   5805
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   741
      Picture         =   "fra_main.frx":40A65
      PictureOver     =   "fra_main.frx":40EA6
      BackColor       =   7464660
   End
   Begin PiLfIuS.LiveButton cmd_addcommand 
      Height          =   435
      Left            =   3585
      TabIndex        =   6
      Top             =   5820
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      Picture         =   "fra_main.frx":414B5
      PictureOver     =   "fra_main.frx":41922
      BackColor       =   7464660
   End
   Begin PiLfIuS.LiveButton cmd_modifycommand 
      Height          =   435
      Left            =   2745
      TabIndex        =   5
      Top             =   5820
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   767
      Picture         =   "fra_main.frx":41F35
      PictureOver     =   "fra_main.frx":424AE
      BackColor       =   7464660
   End
   Begin PiLfIuS.LiveButton cmd_editgroup 
      Height          =   420
      Left            =   3975
      TabIndex        =   1
      Top             =   885
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   741
      Picture         =   "fra_main.frx":42BD0
      PictureOver     =   "fra_main.frx":43144
      BackColor       =   3665108
   End
   Begin PiLfIuS.LiveButton cmd_erasegroup 
      Height          =   420
      Left            =   5265
      TabIndex        =   3
      Top             =   870
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   741
      Picture         =   "fra_main.frx":43647
      PictureOver     =   "fra_main.frx":43A8B
      BackColor       =   3665108
   End
   Begin PiLfIuS.LiveButton cmd_addgroup 
      Height          =   420
      Left            =   4815
      TabIndex        =   2
      Top             =   870
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   741
      Picture         =   "fra_main.frx":440D7
      PictureOver     =   "fra_main.frx":44507
      BackColor       =   3665108
   End
   Begin VB.ComboBox cmb_group 
      Appearance      =   0  'Flat
      BackColor       =   &H0037ECD4&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   705
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   900
      Width           =   3240
   End
   Begin VB.ListBox lst_commands 
      Appearance      =   0  'Flat
      BackColor       =   &H0071E6D4&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3810
      Left            =   795
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   3750
   End
   Begin VB.Image Image26 
      Height          =   150
      Left            =   1470
      Picture         =   "fra_main.frx":44B25
      Top             =   7860
      Width           =   1755
   End
   Begin VB.Image Image25 
      Height          =   150
      Left            =   1665
      Picture         =   "fra_main.frx":44D52
      Top             =   7575
      Width           =   2115
   End
   Begin VB.Image Image5 
      Height          =   1155
      Left            =   630
      Picture         =   "fra_main.frx":44FBC
      Top             =   7365
      Width           =   3930
   End
   Begin VB.Line Line3 
      X1              =   8700
      X2              =   285
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Label lbl_version 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ver 0.6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   10110
      TabIndex        =   39
      Top             =   855
      Width           =   540
   End
   Begin VB.Line Line1 
      X1              =   11415
      X2              =   11415
      Y1              =   1470
      Y2              =   7920
   End
   Begin VB.Label lbl_cmd_help 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   2235
      TabIndex        =   38
      Top             =   5940
      Width           =   405
   End
   Begin VB.Label lbl_grp_help 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   6120
      TabIndex        =   37
      Top             =   1005
      Width           =   405
   End
   Begin VB.Image Image4 
      Height          =   4725
      Left            =   255
      Picture         =   "fra_main.frx":45ABF
      Top             =   1665
      Width           =   4890
   End
   Begin VB.Image Image3 
      Height          =   120
      Left            =   3690
      Picture         =   "fra_main.frx":472B2
      Top             =   1455
      Width           =   1365
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   525
      Picture         =   "fra_main.frx":47656
      Top             =   810
      Width           =   5565
   End
   Begin VB.Image Image1 
      Height          =   150
      Left            =   630
      Picture         =   "fra_main.frx":47D9C
      Top             =   615
      Width           =   1185
   End
End
Attribute VB_Name = "fra_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ToAscii Lib "USER32" (ByVal wVirtKey As Long, ByVal wScanCode As Long, lpKeyState As Any, lpChar As Any, ByVal wFlags As Long) As Long
Private Declare Function GetKeyboardState Lib "USER32" (pbKeyState As Byte) As Long
'API calls for key capture
Private Declare Function GetAsyncKeyState Lib "USER32" (ByVal vKey As Long) As Integer

'constants for key capture stuff
Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Const MOD_SHIFT = &H4

Public vCommandFile As String
Public oCommand As cls_command
Private WithEvents oSpeech As cls_speech
Attribute oSpeech.VB_VarHelpID = -1
Private ChangesMade As Boolean

Private Enum e_ShowMode
    sm_show = 0
    sm_hide = 1
End Enum

Private speechActive As Boolean
Private toggleRequested As Boolean

Dim Key_ShowMode As e_ShowMode
Dim Option_ShowMode As e_ShowMode

Public unloading As Boolean

'These constants have to be the same as the ones in cls_command
Const KP_COUNT = 6
Const KP_KEY = 0
Const KP_HOLD = 1
Const KP_DELAY = 2
Const KP_REPEAT = 3
Const KP_SHIFT = 4
Const KP_CTRL = 5
Const KP_ALT = 6

Private Sub btn_helpconfidence_Click()
    mdi_main.helpWindow "What is ""use Confidence Threshold""?", "If checked, ""use Confidence Threshold"" will verify how certain the Speech Recognition Engine is that you said a command." & vbNewLine & _
    vbNewLine & _
    "If there is a reasonable degree of certainty, the command will be identified and all actions associated with it executed; otherwise it will be discarded." & vbNewLine & _
    vbNewLine & _
    "PROS: Using this will reduce false recognitions (commands being recognized when nothing is said)." & vbNewLine & _
    vbNewLine & _
    "CONS: Depending on your pronunciation accuracy and engine training, genuine commands may not be recognized. Disable this option if you feel PiLfIuS! doesn 't respond to some commands you speak."
End Sub

Private Sub chk_combohold_Click()
    If chk_combohold.Value = 1 Then
        txt_combohold.Locked = False
    Else
        txt_combohold.text = ""
        txt_combohold.Locked = True
    End If
End Sub

Private Sub chk_comborepeat_Click()
    If chk_comborepeat.Value = 1 Then
        txt_comborepeat.Locked = False
    Else
        txt_comborepeat.text = ""
        txt_comborepeat.Locked = True
    End If
End Sub

Private Sub chk_combowait_Click()
    If chk_combowait.Value = 1 Then
        txt_combowait.Locked = False
    Else
        txt_combowait.text = ""
        txt_combowait.Locked = True
    End If
End Sub

Private Sub chk_noactions_Click()
    If chk_noactions.Value = 1 Then
        oSpeech.vNoActions = True
    Else
        oSpeech.vNoActions = False
    End If
End Sub

Private Sub cmb_group_Click()
    lst_commands.Clear
    lst_keystrokes.Clear
    HideAddKeyControls
    
    For i = 1 To oCommand.cCommands.Count
        If cmb_group.ListIndex = 0 Or oCommand.cCommands(i)(1) = cmb_group.ItemData(cmb_group.ListIndex) Then
            lst_commands.AddItem oCommand.cCommands(i)(0)
            lst_commands.ItemData(lst_commands.NewIndex) = i
        End If
    Next
End Sub

Private Sub ShowAddKeyControls()
    If Not fra_options.Visible Then
        Key_ShowMode = sm_show
        tmr_showkeys.Enabled = True
        fra_addkey.Visible = True
        fra_addkeys.Visible = True
    End If
End Sub

Private Sub HideAddKeyControls()
    Key_ShowMode = sm_hide
    tmr_showkeys.Enabled = True
    
    ClearAddKeyControls
End Sub

Private Sub ShowOptionControls()
    Option_ShowMode = sm_show
    tmr_showopt.Enabled = True
    fra_options.Visible = True
End Sub

Private Sub HideOptionControls()
    Option_ShowMode = sm_hide
    tmr_showopt.Enabled = True
    
    'ClearOptionControls
End Sub

Private Sub cmb_mousejoy_Click()
    If cmb_mousejoy.ListIndex > 0 Then
        txt_keycombo.text = ""
        cmb_specialkeys.ListIndex = 0
    End If
End Sub

Private Sub cmb_specialkeys_Click()
    If cmb_specialkeys.ListIndex > 0 Then
        txt_keycombo.text = ""
        cmb_mousejoy.ListIndex = 0
    End If
End Sub

Private Sub cmd_addcommand_Click()
    Dim FoundSameCommand As Boolean
    
    cmd_addcommand.RevertPicture
    
    frm_addcommand.GroupNr = cmb_group.ItemData(cmb_group.ListIndex)
    frm_addcommand.Show 1
    If Not frm_addcommand.Cancelled Then
        'First, we make sure there is no other command for this phrase
        FoundSameCommand = False
        For i = 1 To oCommand.cCommands.Count
            If UCase(oCommand.cCommands(i)(0)) = UCase(frm_addcommand.Command) Then FoundSameCommand = True
        Next
        
        If Not FoundSameCommand Then
            'Add command to oCommand
            oCommand.addCommand frm_addcommand.Command, frm_addcommand.cmb_group.ItemData(frm_addcommand.cmb_group.ListIndex)
            
            'Add new command to the speech recognition dictionary
            oSpeech.LoadGrammarItem frm_addcommand.Command, oCommand.cCommands.Count
            oSpeech.CommitGrammar
            
            'Update data display
            If cmb_group.ItemData(cmb_group.ListIndex) = frm_addcommand.cmb_group.ItemData(frm_addcommand.cmb_group.ListIndex) Or cmb_group.ListIndex = 0 Then
                lst_commands.AddItem frm_addcommand.Command
                lst_commands.ItemData(lst_commands.NewIndex) = oCommand.cCommands.Count
                
                lst_commands.ListIndex = lst_commands.NewIndex
            End If
            
            ChangesMade = True
            
            frm_addcommand.Command = ""
        End If
    End If
    Unload frm_addcommand
End Sub

Private Sub cmd_addcommand_GotFocus()
    lbl_cmd_help.Caption = "Add new voice command (INS)"
End Sub

Private Sub cmd_addcommand_LostFocus()
    lbl_cmd_help.Caption = ""
End Sub

Private Sub cmd_addcommand_MouseLeave()
    cmd_addcommand_LostFocus
End Sub

Private Sub cmd_addcommand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_addcommand_GotFocus
End Sub

Private Sub cmd_addgroup_Click()
    Dim FoundSameGroup As Boolean
    
    frm_addgroup.Show 1
    If Not frm_addgroup.Cancelled Then
        'First, make sure there is no other group with this name
        FoundSameGroup = False
        For i = 0 To cmb_group.ListCount - 1
            If UCase(cmb_group.List(i)) = UCase(frm_addgroup.Group) Then FoundSameGroup = True
        Next
        
        If Not FoundSameGroup Then
            'Add group to oCommand
            oCommand.addGroup frm_addgroup.Group
            
            'Update data display
            cmb_group.AddItem frm_addgroup.Group
            cmb_group.ItemData(cmb_group.NewIndex) = cmb_group.NewIndex
            
            cmb_group.ListIndex = cmb_group.NewIndex
            
            
            ChangesMade = True
        Else
            MsgBox "Another group already exists by that name", vbOKOnly + vbExclamation, "NOT ADDED"
        End If
        frm_addgroup.Group = ""
    End If
    Unload frm_addgroup
End Sub

Private Sub cmd_addgroup_GotFocus()
    lbl_grp_help.Caption = "Add new group"
End Sub

Private Sub cmd_addgroup_LostFocus()
    lbl_grp_help.Caption = ""
End Sub

Private Sub cmd_addgroup_MouseLeave()
    cmd_addgroup_LostFocus
End Sub

Private Sub cmd_addgroup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_addgroup_GotFocus
End Sub

Private Sub cmd_addkeycombo_Click()
    If Len(Trim(txt_keycombo.text)) > 0 Or cmb_specialkeys.ListIndex > 0 Or cmb_mousejoy.ListIndex > 0 Then
        'Add key/button to command in oCommand
        Select Case True
            Case (Len(Trim(txt_keycombo.text)) > 0):
                oCommand.addKey txt_keycombo.text, IIf(Len(txt_combohold.text) > 0, txt_combohold.text, 0), IIf(Len(txt_combowait.text) > 0, txt_combowait.text, 0), lst_commands.ItemData(lst_commands.ListIndex), IIf(Len(txt_comborepeat.text) > 0, txt_comborepeat.text, 0)
            Case cmb_specialkeys.ListIndex > 0:
                oCommand.addKey cmb_specialkeys.List(cmb_specialkeys.ListIndex), IIf(Len(txt_combohold.text) > 0, txt_combohold.text, 0), IIf(Len(txt_combowait.text) > 0, txt_combowait.text, 0), lst_commands.ItemData(lst_commands.ListIndex), IIf(Len(txt_comborepeat.text) > 0, txt_comborepeat.text, 0)
            Case cmb_mousejoy.ListIndex > 0:
                oCommand.addKey cmb_mousejoy.List(cmb_mousejoy.ListIndex), IIf(Len(txt_combohold.text) > 0, txt_combohold.text, 0), IIf(Len(txt_combowait.text) > 0, txt_combowait.text, 0), lst_commands.ItemData(lst_commands.ListIndex), IIf(Len(txt_comborepeat.text) > 0, txt_comborepeat.text, 0)
        End Select
        
        'Update data display
        lst_commands_Click
        
        'Clear add controls
        ClearAddKeyControls
        
        txt_keycombo.SetFocus
        
        ChangesMade = True
    End If
End Sub

Private Sub ClearAddKeyControls()
    txt_keycombo.text = ""
    cmb_specialkeys.ListIndex = 0
    cmb_mousejoy.ListIndex = 0
    chk_combohold.Value = 0
    txt_combohold.text = ""
    chk_combowait.Value = 0
    txt_combowait.text = ""
    chk_comborepeat.Value = 0
    txt_comborepeat.text = ""
End Sub

Private Sub cmd_close_click()
    Unload Me
End Sub

Private Sub cmd_editgroup_Click()
    Dim FoundSameGroup As Boolean
    
    If cmb_group.ListIndex > 0 Then
        lbl_grp_help.Caption = ""
        
        frm_addgroup.Group = cmb_group.List(cmb_group.ListIndex)
        frm_addgroup.Show 1
        If Not frm_addgroup.Cancelled Then
            'First, make sure there is no other group with this name
            FoundSameGroup = False
            For i = 0 To cmb_group.ListCount - 1
                If i <> cmb_group.ListIndex Then
                    If UCase(cmb_group.List(i)) = UCase(frm_addgroup.Group) Then FoundSameGroup = True
                End If
            Next
            
            If Not FoundSameGroup Then
                'Modify group's name
                oCommand.cGroups.Remove cmb_group.ItemData(cmb_group.ListIndex)
                If cmb_group.ItemData(cmb_group.ListIndex) > oCommand.cGroups.Count Then
                    oCommand.cGroups.Add frm_addgroup.Group
                Else
                    oCommand.cGroups.Add frm_addgroup.Group, , cmb_group.ItemData(cmb_group.ListIndex)
                End If
                
                'Update data display
                cmb_group.List(cmb_group.ListIndex) = frm_addgroup.Group
                
                frm_addgroup.Group = ""
                
                ChangesMade = True
            Else
                MsgBox "Another group already exists by that name", vbOKOnly + vbExclamation, "NOT MODIFIED"
            End If
        End If
        Unload frm_addgroup
    End If
End Sub

Private Sub cmd_editgroup_GotFocus()
    lbl_grp_help.Caption = "Edit selected group"
End Sub

Private Sub cmd_editgroup_LostFocus()
    lbl_grp_help.Caption = ""
End Sub

Private Sub cmd_editgroup_MouseLeave()
    cmd_editgroup_LostFocus
End Sub

Private Sub cmd_editgroup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_editgroup_GotFocus
End Sub

Private Sub cmd_erasecommand_Click()
    If lst_commands.ListCount > 0 Then
        If lst_commands.ListIndex > -1 Then
            If MsgBox("Delete command '" & lst_commands.List(lst_commands.ListIndex) & "'?", vbOKCancel + vbQuestion, "DELETE COMMAND") = vbOK Then
                'erase command from oCommand
                cmdErased = lst_commands.ItemData(lst_commands.ListIndex)
                oCommand.eraseCommand cmdErased
                
                'remove command from command list
                lst_commands.RemoveItem lst_commands.ListIndex
                For i = 0 To lst_commands.ListCount - 1
                    If lst_commands.ItemData(i) > cmdErased Then lst_commands.ItemData(i) = lst_commands.ItemData(i) - 1
                Next
                
                'remove command from SAPI dictionary
                oSpeech.RemoveAllGrammar
                
                For i = 1 To oCommand.cCommands.Count
                    oSpeech.LoadGrammarItem CStr(oCommand.cCommands(i)(0)), CInt(i)
                Next
                oSpeech.CommitGrammar
                
                lst_keystrokes.Clear
                HideAddKeyControls
                
                ChangesMade = True
            End If
        End If
    End If
End Sub

Private Sub cmd_erasecommand_GotFocus()
    lbl_cmd_help.Caption = "Delete selected command (DEL)"
End Sub

Private Sub cmd_erasecommand_LostFocus()
    lbl_cmd_help.Caption = ""
End Sub

Private Sub cmd_erasecommand_MouseLeave()
    cmd_erasecommand_LostFocus
End Sub

Private Sub cmd_erasecommand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_erasecommand_GotFocus
End Sub

Private Sub cmd_erasegroup_Click()
    If cmb_group.ListIndex > 0 Then
        lbl_grp_help.Caption = ""
        If MsgBox("Are you sure you want to delete group """ & cmb_group.List(cmb_group.ListIndex) & """?" & vbNewLine & "All commands in the group will be assigned to no group.", vbYesNo + vbExclamation, "ERASE GROUP") = vbYes Then
            oCommand.eraseGroup cmb_group.ItemData(cmb_group.ListIndex)
            
            'Rearrange the combo's itemdata to point to the correct position of group in oCommands
            For i = cmb_group.ListIndex + 1 To cmb_group.ListCount - 1
                cmb_group.ItemData(i) = cmb_group.ItemData(i) - 1
            Next
            
            cmb_group.RemoveItem cmb_group.ListIndex
            cmb_group.ListIndex = 0
            
            ChangesMade = True
        End If
    End If
End Sub

Private Sub cmd_erasegroup_GotFocus()
    lbl_grp_help.Caption = "Delete selected group"
End Sub

Private Sub cmd_erasegroup_LostFocus()
    lbl_grp_help.Caption = ""
End Sub

Private Sub cmd_erasegroup_MouseLeave()
    cmd_erasegroup_LostFocus
End Sub

Private Sub cmd_erasegroup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_erasegroup_GotFocus
End Sub

Private Sub cmd_erasekey_Click()
    If lst_keystrokes.ListCount > 0 Then
        If lst_keystrokes.ListIndex > -1 Then
            oCommand.eraseKey lst_commands.ItemData(lst_commands.ListIndex), lst_keystrokes.ItemData(lst_keystrokes.ListIndex)
            
            lst_commands_Click
            
            ChangesMade = True
        End If
    End If
End Sub

Private Sub cmd_erasekey_GotFocus()
    lbl_key_help.Caption = "Delete selected keystroke (DEL)"
End Sub

Private Sub cmd_erasekey_LostFocus()
    lbl_key_help.Caption = ""
End Sub

Private Sub cmd_erasekey_MouseLeave()
    cmd_erasekey_LostFocus
End Sub

Private Sub cmd_erasekey_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_erasekey_GotFocus
End Sub

Private Sub cmd_addkeycombo_GotFocus()
    lbl_addkeys_help.Caption = "Add new keystroke"
End Sub

Private Sub cmd_addkeycombo_LostFocus()
    lbl_addkeys_help.Caption = ""
End Sub

Private Sub cmd_addkeycombo_MouseLeave()
    cmd_addkeycombo_LostFocus
End Sub

Private Sub cmd_addkeycombo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_addkeycombo_GotFocus
End Sub

Private Sub cmd_keydown_Click()
    MoveKeyStroke -1
End Sub

Private Sub MoveKeyStroke(MoveDir As Integer)
    Dim tempOrder As Variant
    
    If lst_keystrokes.ListCount > 0 Then
        If lst_keystrokes.ListIndex > -1 Then
            If (MoveDir > 0 And lst_keystrokes.ListIndex > 0) Or (MoveDir < 0 And lst_keystrokes.ListIndex < lst_keystrokes.ListCount - 1) Then
                'Move keystroke up and store in tempOrder
                For i = 0 To UBound(oCommand.cKeys(lst_commands.ItemData(lst_commands.ListIndex)), 1)
                    If i = 0 Then
                        ReDim tempOrder(0)
                    Else
                        ReDim Preserve tempOrder(UBound(tempOrder, 1) + 1)
                    End If
                    
                    If i = lst_keystrokes.ListIndex - MoveDir Then
                        tempOrder(UBound(tempOrder, 1)) = oCommand.cKeys(lst_commands.ItemData(lst_commands.ListIndex))(i + MoveDir)
                    ElseIf i = lst_keystrokes.ListIndex Then
                        tempOrder(UBound(tempOrder, 1)) = oCommand.cKeys(lst_commands.ItemData(lst_commands.ListIndex))(i - MoveDir)
                    Else
                        tempOrder(UBound(tempOrder, 1)) = oCommand.cKeys(lst_commands.ItemData(lst_commands.ListIndex))(i)
                    End If
                Next
                
                'Reflect tempOrder in the cKeys
                oCommand.cKeys.Remove lst_commands.ItemData(lst_commands.ListIndex)
                
                If lst_commands.ItemData(lst_commands.ListIndex) > oCommand.cKeys.Count Then
                    oCommand.cKeys.Add tempOrder
                Else
                    oCommand.cKeys.Add tempOrder, , lst_commands.ItemData(lst_commands.ListIndex)
                End If
                
                keySelect = lst_keystrokes.ListIndex
                lst_commands_Click
                lst_keystrokes.ListIndex = keySelect - MoveDir
            End If
        End If
    End If
End Sub

Private Sub cmd_keyup_Click()
    MoveKeyStroke 1
End Sub

Private Sub cmd_keyup_GotFocus()
    lbl_key_help.Caption = "Move keystroke up"
End Sub

Private Sub cmd_keyup_LostFocus()
    lbl_key_help.Caption = ""
End Sub

Private Sub cmd_keyup_MouseLeave()
    cmd_keyup_LostFocus
End Sub

Private Sub cmd_keyup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_keyup_GotFocus
End Sub

Private Sub cmd_keydown_GotFocus()
    lbl_key_help.Caption = "Move keystroke down"
End Sub

Private Sub cmd_keydown_LostFocus()
    lbl_key_help.Caption = ""
End Sub

Private Sub cmd_keydown_MouseLeave()
    cmd_keydown_LostFocus
End Sub

Private Sub cmd_keydown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_keydown_GotFocus
End Sub


Private Sub cmd_modifycommand_Click()
    Dim foundCommand As Boolean
    
    If lst_commands.ListCount > 0 Then
        If lst_commands.ListIndex > -1 Then
            frm_addcommand.Command = lst_commands.List(lst_commands.ListIndex)
            frm_addcommand.GroupNr = oCommand.cCommands(lst_commands.ItemData(lst_commands.ListIndex))(1)
            cmd_modifycommand.RevertPicture
            frm_addcommand.Show 1
            If Not frm_addcommand.Cancelled Then
                'Modify command in oCommand
                oCommand.modifyCommand lst_commands.ItemData(lst_commands.ListIndex), frm_addcommand.Command, frm_addcommand.cmb_group.ItemData(frm_addcommand.cmb_group.ListIndex)
                
                'Modify command speech recognition dictionary
                oSpeech.RemoveAllGrammar
                For i = 1 To oCommand.cCommands.Count
                    oSpeech.LoadGrammarItem CStr(oCommand.cCommands(i)(0)), CInt(i)
                Next
                oSpeech.CommitGrammar
                
                'Update data display
                If cmb_group.ItemData(cmb_group.ListIndex) = frm_addcommand.cmb_group.ItemData(frm_addcommand.cmb_group.ListIndex) Or cmb_group.ListIndex = 0 Then
                    lst_commands.List(lst_commands.ListIndex) = frm_addcommand.Command
                Else
                    lst_commands.RemoveItem lst_commands.ListIndex
                End If
                
                frm_addcommand.Command = ""
                frm_addcommand.GroupNr = Empty
                
                ChangesMade = True
            End If
            Unload frm_addcommand
        End If
    End If
End Sub

Private Sub cmd_modifycommand_GotFocus()
    lbl_cmd_help.Caption = "Edit selected command"
End Sub

Private Sub cmd_modifycommand_LostFocus()
    lbl_cmd_help.Caption = ""
End Sub

Private Sub cmd_modifycommand_MouseLeave()
    cmd_modifycommand_LostFocus
End Sub

Private Sub cmd_modifycommand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_modifycommand_GotFocus
End Sub


Private Sub cmd_opt_cancel_Click()
    HideOptionControls
End Sub


Private Sub cmd_opt_change_Click()
    frm_changeactivation.selectedAction = lbl_opt_key.Caption
    frm_changeactivation.Show 1
    If Not frm_changeactivation.Cancelled Then
        lbl_opt_key.Caption = frm_changeactivation.selectedAction
    End If
    Unload frm_changeactivation
End Sub

Private Sub cmd_opt_ok_Click()
    Dim msgerror As String
    
    msgerror = ""
    
    If Len(Trim(lbl_opt_key.Caption)) = 0 And Not opt_activation(0) Then msgerror = msgerror & "- You must select an activation key." & vbNewLine
    
    If Len(msgerror) > 0 Then
        MsgBox msgerror, vbOKOnly + vbExclamation, "CANNOT CHANGE OPTIONS"
    Else
        'change options
        For i = 0 To opt_activation.Count - 1
            If opt_activation(i) Then
                oCommand.vActivationType = i
                oSpeech.vActivationType = i
            End If
        Next
        oCommand.assignActivationKey lbl_opt_key.Caption
        oSpeech.vActivationKey = oCommand.vActivationKey
        
        'Confidence Threshold
        If chk_confidence.Value = 1 Then oCommand.vConfidenceThreshold = True Else oCommand.vConfidenceThreshold = False
        
        HideOptionControls
    End If
End Sub

Private Sub cmd_opt_ok_GotFocus()
    lbl_opt_help.Caption = "Accept changes"
End Sub

Private Sub cmd_opt_ok_LostFocus()
    lbl_opt_help.Caption = ""
End Sub

Private Sub cmd_opt_ok_MouseLeave()
    cmd_opt_ok_LostFocus
End Sub

Private Sub cmd_opt_ok_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_opt_ok_GotFocus
End Sub

Private Sub cmd_opt_cancel_GotFocus()
    lbl_opt_help.Caption = "Cancel changes"
End Sub

Private Sub cmd_opt_cancel_LostFocus()
    lbl_opt_help.Caption = ""
End Sub

Private Sub cmd_opt_cancel_MouseLeave()
    cmd_opt_cancel_LostFocus
End Sub

Private Sub cmd_opt_cancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_opt_cancel_GotFocus
End Sub

Private Sub cmd_save_Click()
    If Len(vCommandFile) > 0 Then
        cmd_save.RevertPicture
        SaveFile vCommandFile
        
        ChangesMade = False
    Else
        cmd_saveas_Click
    End If
End Sub

Private Sub cmd_saveas_Click()
    Dim proceedLoad As Boolean
    Dim framainLoaded As Boolean
    
    cmd_saveas.RevertPicture
    mdi_main.dlg_main.DialogTitle = "Save PiLfIuS! Command List"
    mdi_main.dlg_main.Filter = ".lcl (PiLfIuS! Command List)|*.lcl"
    mdi_main.dlg_main.FileName = vCommandFile
    mdi_main.dlg_main.CancelError = True
    On Error GoTo EH_SAVEAS
    mdi_main.dlg_main.ShowSave
    If Len(mdi_main.dlg_main.FileName) > 0 Then
        vCommandFile = mdi_main.dlg_main.FileName
        SaveFile vCommandFile
        
        ChangesMade = False
    End If
    On Error GoTo 0
EH_SAVEAS:
End Sub



Private Sub Form_Activate()
    'start check loop for hotkey messages
    toggleRequested = False
    tmr_activkey.Enabled = True
End Sub

Public Sub FillSpecialKeys(destObj As Object)
    destObj.AddItem "-- special keys --"
    destObj.AddItem "{TAB}"
    destObj.AddItem "Shift + {TAB}"
    destObj.AddItem "{F10}"
    destObj.AddItem "{PRNT SCRN}"
    destObj.AddItem "Alt + {F4}"
    destObj.AddItem "Ctrl + {F4}"
    destObj.AddItem "Alt + {SPACE}"
    destObj.AddItem "Ctrl + Alt + {DEL}"
    destObj.AddItem "Ctrl + Alt + {KP_DECIMAL}"
    
    destObj.ListIndex = 0
End Sub

Public Sub FillMouseJoy(destObj As Object, addWheel As Boolean, addX678 As Boolean)
    destObj.AddItem "-- buttons --"
    destObj.AddItem "MOUSE LEFT"
    destObj.AddItem "MOUSE MIDDLE"
    destObj.AddItem "MOUSE RIGHT"
    If addWheel Then
        destObj.AddItem "MOUSE WHEELUP"
        destObj.AddItem "MOUSE WHEELDN"
    End If
    destObj.AddItem "MOUSE4"
    destObj.AddItem "MOUSE5"
    If addX678 Then
        destObj.AddItem "MOUSE6"
        destObj.AddItem "MOUSE7"
        destObj.AddItem "MOUSE8"
    End If
    
    destObj.ListIndex = 0
End Sub

Private Sub Form_Load()
    'Fill Special Keys combo
    FillSpecialKeys cmb_specialkeys
    
    'Fill Mouse/Joystick combo
    FillMouseJoy cmb_mousejoy, True, False
    
    'Load file into oCommand
    Set oCommand = New cls_command
    
    If Len(vCommandFile) > 0 Then
        oCommand.LoadCommands vCommandFile
    Else
        oCommand.vConfidenceThreshold = True
    End If
    
    'Load group combo based on oCommand data
    LoadGroups
    
    'Load option data based on oCommand data
    'LoadOptions
    
    'Initialize speech
    Set oSpeech = New cls_speech
    
    If oSpeech.InitRecognition Then
        'Add commands to speech dictionary
        For i = 1 To oCommand.cCommands.Count
            oSpeech.LoadGrammarItem CStr(oCommand.cCommands(i)(0)), CInt(i)
        Next
        
        oSpeech.CommitGrammar
        Set oSpeech.oCommandSet = oCommand
        
        'initialize speechActive state
        Select Case oCommand.vActivationType
            Case 0, 2: RecognitionOn True
            Case 1, 3: RecognitionOn False
        End Select
        showActivationKeyAssignment
        
        'Save activation type and key to oSpeech - activation key will temporarily be "unpushed" before
        'sending keystokes, and then "pushed" again
        oSpeech.vActivationType = oCommand.vActivationType
        oSpeech.vActivationKey = oCommand.vActivationKey
        
        ChangesMade = False
        
        Me.Left = 0
        Me.Top = 0
        Me.Height = 8505
        fra_center.Height = 5730
        lbl_version.Caption = mdi_main.vVersionNumber
        
        fra_addkey.Top = fra_addkey.Top - 2550
        img_keys.Top = img_keys.Top - 2550
        
        fra_addkeys.Left = fra_addkeys.Left + 3400
        img_addkeys.Left = img_addkeys.Left + 3400
    End If
    unloading = False
End Sub

Private Sub LoadGroups()
    cmb_group.Clear
    
    cmb_group.AddItem "All groups"
    cmb_group.ItemData(cmb_group.NewIndex) = 0

    For i = 1 To oCommand.cGroups.Count
        cmb_group.AddItem oCommand.cGroups(i)
        cmb_group.ItemData(cmb_group.NewIndex) = i
    Next
    
    cmb_group.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not unloading Then
        If ChangesMade Then
            Select Case MsgBox("The file has changed since the last save." & vbNewLine & vbNewLine & "¿Do you want to save?", vbYesNoCancel + vbExclamation, "SAVE?")
                Case vbYes: If Len(vCommandFile) > 0 Then SaveFile vCommandFile Else cmd_saveas_Click
                            unloading = True
                Case vbCancel: Cancel = 1
                Case vbNo: unloading = True
            End Select
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmr_activkey.Enabled = False
    
    vCommandFile = ""
    Set oCommand = Nothing
    Set oSpeech = Nothing
    
    frm_main.Show
End Sub



Private Sub img_about_Click()
    frm_about.Show
End Sub

Private Sub img_opt_button_Click()
    If Not fra_options.Visible Then
        If fra_addkey.Visible Then HideAddKeyControls
        showActivationKeyAssignment
        chk_confidence.Value = IIf(oCommand.vConfidenceThreshold, 1, 0)
        fra_options.Visible = True
        ShowOptionControls
    End If
End Sub


Private Sub img_opttext_Click()
    img_opt_button_Click
End Sub

Private Sub lbl_opt_key_Click()
''''''''''''''''''OPT_KEY
End Sub

Private Sub lst_commands_Click()
    Dim keyDesc As String
    
    lst_keystrokes.Clear
    If IsArray(oCommand.cKeys(lst_commands.ItemData(lst_commands.ListIndex))) Then
        For i = 0 To UBound(oCommand.cKeys(lst_commands.ItemData(lst_commands.ListIndex)), 1)
            keyDesc = ""
            If oCommand.cKeys(lst_commands.ItemData(lst_commands.ListIndex))(i)(KP_SHIFT) = 1 Then keyDesc = keyDesc & "Shift + "
            If oCommand.cKeys(lst_commands.ItemData(lst_commands.ListIndex))(i)(KP_CTRL) = 1 Then keyDesc = keyDesc & "Ctrl + "
            If oCommand.cKeys(lst_commands.ItemData(lst_commands.ListIndex))(i)(KP_ALT) = 1 Then keyDesc = keyDesc & "Alt + "
            keyDesc = keyDesc & oCommand.cKeys(lst_commands.ItemData(lst_commands.ListIndex))(i)(KP_KEY)
            If oCommand.cKeys(lst_commands.ItemData(lst_commands.ListIndex))(i)(KP_REPEAT) > 0 Then keyDesc = keyDesc & " x" & oCommand.cKeys(lst_commands.ItemData(lst_commands.ListIndex))(i)(KP_REPEAT)
            If oCommand.cKeys(lst_commands.ItemData(lst_commands.ListIndex))(i)(KP_HOLD) > 0 Then keyDesc = keyDesc & " h:" & oCommand.cKeys(lst_commands.ItemData(lst_commands.ListIndex))(i)(KP_HOLD) & "ms"
            If oCommand.cKeys(lst_commands.ItemData(lst_commands.ListIndex))(i)(KP_DELAY) > 0 Then keyDesc = keyDesc & " w:" & oCommand.cKeys(lst_commands.ItemData(lst_commands.ListIndex))(i)(KP_DELAY) & "ms"
            
            lst_keystrokes.AddItem keyDesc
            lst_keystrokes.ItemData(lst_keystrokes.NewIndex) = i
        Next
    End If
    ShowAddKeyControls
End Sub

Private Sub mnu_file_close_Click()
    Unload Me
End Sub

Private Sub SaveFile(vDestFile As String)
    Dim activationType As String
    
    Open vDestFile For Output As #1
    
    'First of all, save the options
    'ACTIVATION TYPE
    Select Case oCommand.vActivationType
         Case 0: activationType = "AA"
         Case 1: activationType = "PTA"
         Case 2: activationType = "PTD"
         Case 3: activationType = "PTT"
    End Select
    
    Print #1, "ACTIVATION " & activationType
    If oCommand.vActivationType > 0 Then
        Print #1, "KEY " & oCommand.vActivationKey(KP_KEY)
        Print #1, "SHIFT=" & oCommand.vActivationKey(KP_SHIFT)
        Print #1, "CTRL=" & oCommand.vActivationKey(KP_CTRL)
        Print #1, "ALT=" & oCommand.vActivationKey(KP_ALT)
    End If
    
    'CONFIDENCE THRESHOLD
    Print #1, "CONFIDENCETHRESHOLD " & IIf(oCommand.vConfidenceThreshold, "1", "0")
    
    'Commands: First add commands with no group
    For i = 1 To oCommand.cCommands.Count
        If oCommand.cCommands(i)(1) = 0 Then
            Print #1, "COMMAND " & oCommand.cCommands(i)(0)
            If IsArray(oCommand.cKeys(i)) Then
                For j = 0 To UBound(oCommand.cKeys(i), 1)
                    Print #1, "KEY " & oCommand.cKeys(i)(j)(KP_KEY)
                    If oCommand.cKeys(i)(j)(KP_HOLD) > 0 Then Print #1, "HOLD=" & oCommand.cKeys(i)(j)(KP_HOLD)
                    If oCommand.cKeys(i)(j)(KP_DELAY) > 0 Then Print #1, "DELAY=" & oCommand.cKeys(i)(j)(KP_DELAY)
                    If oCommand.cKeys(i)(j)(KP_REPEAT) > 0 Then Print #1, "REPEAT=" & oCommand.cKeys(i)(j)(KP_REPEAT)
                    Print #1, "SHIFT=" & oCommand.cKeys(i)(j)(KP_SHIFT)
                    Print #1, "CTRL=" & oCommand.cKeys(i)(j)(KP_CTRL)
                    Print #1, "ALT=" & oCommand.cKeys(i)(j)(KP_ALT)
                Next
            End If
        End If
    Next
    
    'Now Print grouped commands
    For i = 1 To oCommand.cGroups.Count
        Print #1, "GROUP " & oCommand.cGroups(i)
        
        For k = 1 To oCommand.cCommands.Count
            If oCommand.cCommands(k)(1) = i Then
                Print #1, "COMMAND " & oCommand.cCommands(k)(0)
                If IsArray(oCommand.cKeys(k)) Then
                    For j = 0 To UBound(oCommand.cKeys(k), 1)
                        Print #1, "KEY " & oCommand.cKeys(k)(j)(KP_KEY)
                        If oCommand.cKeys(k)(j)(KP_HOLD) > 0 Then Print #1, "HOLD=" & oCommand.cKeys(k)(j)(KP_HOLD)
                        If oCommand.cKeys(k)(j)(KP_DELAY) > 0 Then Print #1, "DELAY=" & oCommand.cKeys(k)(j)(KP_DELAY)
                        If oCommand.cKeys(k)(j)(KP_REPEAT) > 0 Then Print #1, "REPEAT=" & oCommand.cKeys(k)(j)(KP_REPEAT)
                        Print #1, "SHIFT=" & oCommand.cKeys(k)(j)(KP_SHIFT)
                        Print #1, "CTRL=" & oCommand.cKeys(k)(j)(KP_CTRL)
                        Print #1, "ALT=" & oCommand.cKeys(k)(j)(KP_ALT)
                    Next
                End If
            End If
        Next
    Next
    Close #1
    
    MsgBox "Command list has been successfully saved to:" & vbNewLine & vDestFile, vbOKOnly + vbInformation, "SAVED"
End Sub

Private Sub lst_commands_DblClick()
    cmd_modifycommand_Click
End Sub


Private Sub lst_commands_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 45: cmd_addcommand_Click
        Case 46: cmd_erasecommand_Click
    End Select
End Sub

Private Sub lst_keystrokes_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 46: cmd_erasekey_Click
    End Select
End Sub

Private Sub opt_activation_Click(Index As Integer)
    Set img_opt_activkey.Picture = img_activation(Index).Picture
    img_opt_key.Left = img_opt_activkey.Left + img_opt_activkey.Width + 45
    
    If Index = 0 Then
        lbl_opt_key.Visible = False
        cmd_opt_change.Visible = False
        img_opt_activkey.Visible = False
        img_opt_key.Visible = False
    Else
        lbl_opt_key.Visible = True
        cmd_opt_change.Visible = True
        img_opt_activkey.Visible = True
        img_opt_key.Visible = True
    End If
End Sub

Private Sub oSpeech_Recognized(Command As String)
    txt_lastreco.text = Command
    txt_lastreco.BackColor = &H959166
    tmr_lastreco.Enabled = True
End Sub

Private Sub tmr_activkey_Timer()
    If oCommand.vActivationType > 0 Then
        'NOTE: always remember to pass vActivationType and vActivationKey from oCommand to oSpeech when they change!!!
        'Otherwise, isPressedActivation (and a bunch of other stuff) WON'T work.
        If oSpeech.isPressedActivation Then
            Select Case oCommand.vActivationType
                Case 1:
                        If Not speechActive Then
                            RecognitionOn True
                        End If
                Case 2:
                        If speechActive Then
                            RecognitionOn False
                        End If
                Case 3:
                        toggleRequested = True
            End Select
        Else
            Select Case oCommand.vActivationType
                Case 1:
                        If speechActive Then
                            RecognitionOn False
                        End If
                Case 2:
                        If Not speechActive Then
                            RecognitionOn True
                        End If
                Case 3:
                        If toggleRequested Then
                            If speechActive Then
                                RecognitionOn False
                            Else
                                RecognitionOn True
                            End If
                            toggleRequested = False
                        End If
            End Select
        End If
    Else
        If Not speechActive Then
            RecognitionOn True
        End If
    End If
End Sub

Private Sub tmr_lastreco_Timer()
    txt_lastreco.BackColor = &H656136
    tmr_lastreco.Enabled = False
End Sub

Private Sub tmr_showkeys_Timer()
    Select Case Key_ShowMode
        Case sm_show:
                    If img_keys.Top < 75 Then
                        img_keys.Top = img_keys.Top + 200
                        fra_addkey.Top = fra_addkey.Top + 200
                    End If
                    
                    If img_addkeys.Left > 3180 Then
                        img_addkeys.Left = img_addkeys.Left - 200
                        fra_addkeys.Left = fra_addkeys.Left - 200
                    Else
                        tmr_showkeys.Enabled = False
                    End If
        Case sm_hide:
                    If img_keys.Top > -2550 Then
                        img_keys.Top = img_keys.Top - 200
                        fra_addkey.Top = fra_addkey.Top - 200
                    Else
                        fra_addkey.Visible = False
                    End If
                    
                    If img_addkeys.Left < 6615 Then
                        img_addkeys.Left = img_addkeys.Left + 200
                        fra_addkeys.Left = fra_addkeys.Left + 200
                    Else
                        tmr_showkeys.Enabled = False
                        fra_addkeys.Visible = False
                    End If
    End Select
End Sub

Private Sub tmr_showopt_Timer()
    Select Case Option_ShowMode
        Case sm_show:
                    If img_opt_button.Top > 1600 Then
                        If Not fra_addkey.Visible Then
                            img_opt_button.Top = img_opt_button.Top - 200
                            img_opttext.Top = img_opttext.Top - 200
                            img_optback.Top = img_optback.Top - 200
                            fra_options.Top = fra_options.Top - 200
                        End If
                    Else
                        tmr_showopt.Enabled = False
                    End If
        Case sm_hide:
                    If img_opt_button.Top < 5355 Then
                        img_opt_button.Top = img_opt_button.Top + 200
                        img_opttext.Top = img_opttext.Top + 200
                        img_optback.Top = img_optback.Top + 200
                        fra_options.Top = fra_options.Top + 200
                    Else
                        tmr_showopt.Enabled = False
                        fra_options.Visible = False
                    End If
    End Select
End Sub

Private Sub txt_combohold_Change()
    If Not IsNumeric(txt_combohold.text) Then
        txt_combohold.text = ""
    Else
        If txt_combohold.text > 65535 Then txt_combohold.text = 65535
        txt_combohold.text = Abs(txt_combohold.text)
    End If
End Sub

Private Sub txt_comborepeat_Change()
    If Not IsNumeric(txt_comborepeat.text) Then
        txt_comborepeat.text = ""
    Else
        If txt_comborepeat.text > 50 Then txt_comborepeat.text = 50
        txt_comborepeat.text = Abs(txt_comborepeat.text)
    End If
End Sub

Private Sub txt_combowait_Change()
    If Not IsNumeric(txt_combowait.text) Then
        txt_combowait.text = ""
    Else
        If txt_combowait.text > 65535 Then txt_combowait.text = 65535
        txt_combowait.text = Abs(txt_combowait.text)
    End If
End Sub

Private Sub txt_combowait_GotFocus()
    txt_combowait.SelStart = 0
    txt_combowait.SelLength = Len(txt_combowait.text)
End Sub


Private Sub txt_keycombo_KeyDown(KeyCode As Integer, Shift As Integer)
    comboName = getComboName(KeyCode, Shift)
    
    txt_keycombo.text = comboName
    
    cmb_specialkeys.ListIndex = 0
    cmb_mousejoy.ListIndex = 0
End Sub

Private Function GetAsciiFromKeyCode(KeyCode As Integer) As Integer
    Dim ShiftDown, AltDown, CtrlDown, Txt
    Dim KB_State(0 To 255) As Byte
    
    GetKeyboardState KB_State(0)
    KB_State(&H12) = 0
    KB_State(&H11) = 0
    KB_State(&H10) = 0
    
    Dim ansi As Integer
    RetValue = ToAscii(CLng(KeyCode), 0, KB_State(0), ansi, 0&)
    GetAsciiFromKeyCode = ansi
End Function


Private Sub RecognitionOn(status As Boolean)
    If status Then
        speechActive = True
        oSpeech.ResumeRecognition
        
        img_listen_on.Visible = True
        img_listen_off.Visible = False
    Else
        speechActive = False
        oSpeech.PauseRecognition
        
        img_listen_on.Visible = False
        img_listen_off.Visible = True
    End If
End Sub

Private Sub showActivationKeyAssignment()
    Dim keyDesc As String
    
    opt_activation(oCommand.vActivationType).Value = True
    
    keyDesc = ""
    If oCommand.vActivationType > 0 Then
        If oCommand.vActivationKey(KP_SHIFT) = 1 Then keyDesc = keyDesc & "Shift + "
        If oCommand.vActivationKey(KP_CTRL) = 1 Then keyDesc = keyDesc & "Ctrl + "
        If oCommand.vActivationKey(KP_ALT) = 1 Then keyDesc = keyDesc & "Alt + "
        keyDesc = keyDesc & oCommand.vActivationKey(KP_KEY)
    End If
    lbl_opt_key.Caption = keyDesc
    lbl_opt_key.Visible = True
End Sub

Public Function getComboName(KeyCode As Integer, Shift As Integer) As String
    Dim comboName As String
    
    comboName = ""
    
    Select Case Shift
        Case 1: comboName = comboName & "Shift + "
        Case 2: comboName = comboName & "Ctrl + "
        Case 3: comboName = comboName & "Shift + Ctrl + "
        Case 4: comboName = comboName & "Alt + "
        Case 5: comboName = comboName & "Shift + Alt + "
        Case 6: comboName = comboName & "Ctrl + Alt + "
        Case 7: comboName = comboName & "Shift + Ctrl + Alt + "
    End Select
    
    ansi = GetAsciiFromKeyCode(KeyCode)
    
    If KeyCode < 16 Or KeyCode > 18 Then
        Select Case KeyCode
            Case vbKeyF1: comboName = comboName & "{F1}"
            Case vbKeyF2: comboName = comboName & "{F2}"
            Case vbKeyF3: comboName = comboName & "{F3}"
            Case vbKeyF4: comboName = comboName & "{F4}"
            Case vbKeyF5: comboName = comboName & "{F5}"
            Case vbKeyF6: comboName = comboName & "{F6}"
            Case vbKeyF7: comboName = comboName & "{F7}"
            Case vbKeyF8: comboName = comboName & "{F8}"
            Case vbKeyF9: comboName = comboName & "{F9}"
            Case vbKeyF10: comboName = comboName & "{F10}"
            Case vbKeyF11: comboName = comboName & "{F11}"
            Case vbKeyF12: comboName = comboName & "{F12}"
            Case vbKeyReturn: comboName = comboName & "{ENTER}"
            Case vbKeyBack: comboName = comboName & "{BACKSPACE}"
            Case vbKeyEscape: comboName = comboName & "{ESC}"
            Case vbKeyPrint: comboName = comboName & "{PRNT SCRN}"
            Case vbKeyScrollLock: comboName = comboName & "{SCROLL LOCK}"
            Case vbKeyPause: comboName = comboName & "{PAUSE}"
            'Case vbKeyTab: comboName = comboName & "{TAB}"
            Case vbKeyCapital: comboName = comboName & "{CAPS LOCK}"
            Case vbKeySpace: comboName = comboName & "{SPACE}"
            Case vbKeyInsert: comboName = comboName & "{INSERT}"
            Case vbKeyHome: comboName = comboName & "{HOME}"
            Case vbKeyPageUp: comboName = comboName & "{PGUP}"
            Case vbKeyPageDown: comboName = comboName & "{PGDN}"
            Case vbKeyDelete: comboName = comboName & "{DEL}"
            Case vbKeyEnd: comboName = comboName & "{END}"
            Case vbKeyUp: comboName = comboName & "{ARROWUP}"
            Case vbKeyDown: comboName = comboName & "{ARROWDOWN}"
            Case vbKeyLeft: comboName = comboName & "{ARROWLEFT}"
            Case vbKeyRight: comboName = comboName & "{ARROWRIGHT}"
            Case vbKeyNumlock: comboName = comboName & "{NUM LOCK}"
            Case vbKeyNumpad0: comboName = comboName & "{KP_0}"
            Case vbKeyNumpad1: comboName = comboName & "{KP_1}"
            Case vbKeyNumpad2: comboName = comboName & "{KP_2}"
            Case vbKeyNumpad3: comboName = comboName & "{KP_3}"
            Case vbKeyNumpad4: comboName = comboName & "{KP_4}"
            Case vbKeyNumpad5: comboName = comboName & "{KP_5}"
            Case vbKeyNumpad6: comboName = comboName & "{KP_6}"
            Case vbKeyNumpad7: comboName = comboName & "{KP_7}"
            Case vbKeyNumpad8: comboName = comboName & "{KP_8}"
            Case vbKeyNumpad9: comboName = comboName & "{KP_9}"
            Case vbKeyMultiply: comboName = comboName & "{KP_MULTIPLY}"
            Case vbKeyDivide: comboName = comboName & "{KP_DIVIDE}"
            Case vbKeySubtract: comboName = comboName & "{KP_SUBTRACT}"
            Case vbKeyAdd: comboName = comboName & "{KP_ADD}"
            Case vbKeyDecimal: comboName = comboName & "{KP_DECIMAL}"
            
            Case Else: comboName = comboName & (Chr(ansi))
        End Select
    End If
    
    getComboName = comboName
End Function
