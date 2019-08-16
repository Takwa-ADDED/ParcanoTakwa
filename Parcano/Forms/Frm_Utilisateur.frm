VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form Frm_Utilisateur 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Parcano"
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Utilisateur.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   16545
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   2160
      ScaleHeight     =   2985
      ScaleWidth      =   4065
      TabIndex        =   129
      Top             =   2280
      Width           =   4095
      Begin SToolBox.SGrid SGrid1 
         Height          =   3855
         Left            =   0
         TabIndex        =   130
         Top             =   0
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6800
         RowMode         =   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisableIcons    =   -1  'True
         MaxVisibleRows  =   0
      End
   End
   Begin VB.CheckBox chk_Actif 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   3000
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listes des Accées"
      ForeColor       =   &H000040C0&
      Height          =   7815
      Left            =   0
      TabIndex        =   9
      Top             =   3360
      Width           =   15375
      Begin VB.CheckBox Chk_tous 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sélectionner tous"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   131
         Top             =   480
         Width           =   1815
      End
      Begin VB.Frame Frame22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bon / CMD / Factures"
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   120
         TabIndex        =   95
         Top             =   3960
         Width           =   15135
         Begin VB.Frame Frame26 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   6600
            TabIndex        =   123
            Top             =   240
            Width           =   2055
            Begin VB.CheckBox Chk_consult_pr 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consulter PR"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   127
               Top             =   960
               Width           =   1815
            End
            Begin VB.CheckBox chk_maj_pr 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj PR"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   126
               Top             =   480
               Width           =   1815
            End
            Begin VB.CheckBox chk_ins_pr 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Insérer  PR"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   125
               Top             =   240
               Width           =   1815
            End
            Begin VB.CheckBox chk_supp_pr 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Supprimer PR"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   124
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label28 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "P. Réparation"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   128
               Top             =   0
               Width           =   1140
            End
         End
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   4440
            TabIndex        =   117
            Top             =   240
            Width           =   2055
            Begin VB.CheckBox Chk_supp_bcr 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Supprimer B.Cmd.R"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   121
               Top             =   720
               Width           =   1815
            End
            Begin VB.CheckBox chk_ins_bcr 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Inserer B.Cmd.R"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   120
               Top             =   240
               Width           =   1815
            End
            Begin VB.CheckBox Chk_maj_bcr 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj B.Cmd.R"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   119
               Top             =   480
               Width           =   1815
            End
            Begin VB.CheckBox Chk_consult_bcr 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consulter B.Cmd.R"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   118
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "B. Cmd Réparation"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   122
               Top             =   0
               Width           =   1560
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   8760
            TabIndex        =   111
            Top             =   240
            Width           =   2055
            Begin VB.CheckBox chk_supp_ff 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Supprimer Fature"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   115
               Top             =   720
               Width           =   1815
            End
            Begin VB.CheckBox chk_ins_ff 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Inserer Facture"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   114
               Top             =   240
               Width           =   1815
            End
            Begin VB.CheckBox chk_maj_ff 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj Facture"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   113
               Top             =   480
               Width           =   1815
            End
            Begin VB.CheckBox chk_consult_ff 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consulter Fature"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   112
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label Label5 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Facture Fournisseur"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   116
               Top             =   0
               Width           =   1680
            End
         End
         Begin VB.Frame Frame25 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   10920
            TabIndex        =   108
            Top             =   240
            Width           =   2055
            Begin VB.CheckBox Chk_consult_alrt 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consultation G.Alerte"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   120
               TabIndex        =   109
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label27 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Gestion d'Alerte"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   110
               Top             =   0
               Width           =   1365
            End
         End
         Begin VB.Frame Frame24 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   2280
            TabIndex        =   102
            Top             =   240
            Width           =   2055
            Begin VB.CheckBox Chk_consult_bv 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consulter BV"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   106
               Top             =   960
               Width           =   1815
            End
            Begin VB.CheckBox Chk_Maj_BV 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj. Bon Vidange"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   105
               Top             =   480
               Width           =   1815
            End
            Begin VB.CheckBox chk_ins_bv 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Ins. Bon Vidange"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   104
               Top             =   240
               Width           =   1815
            End
            Begin VB.CheckBox Chk_supp_bv 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Supprimer BV"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   103
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label26 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Bon Vidange"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   107
               Top             =   0
               Width           =   1035
            End
         End
         Begin VB.Frame Frame23 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   120
            TabIndex        =   96
            Top             =   240
            Width           =   2055
            Begin VB.CheckBox chk_Supp_BC 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Supprimer BC"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   100
               Top             =   720
               Width           =   1815
            End
            Begin VB.CheckBox Chk_Ins_bc 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Ins. Bon Carburant"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   99
               Top             =   240
               Width           =   1815
            End
            Begin VB.CheckBox Chk_Maj_bc 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj. Bon Carburant"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   98
               Top             =   480
               Width           =   1815
            End
            Begin VB.CheckBox Chk_consult_bc 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consulter BC"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   97
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label Label25 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Bon Carburant"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   101
               Top             =   0
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gestion de Trafic"
         ForeColor       =   &H000000FF&
         Height          =   2895
         Left            =   8760
         TabIndex        =   63
         Top             =   1080
         Width           =   6495
         Begin VB.Frame Frame6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   4560
            TabIndex        =   92
            Top             =   1560
            Width           =   1815
            Begin VB.CheckBox chk_SC 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consulter Statistiques"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   120
               TabIndex        =   93
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label8 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Statistiques"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   94
               Top             =   0
               Width           =   1020
            End
         End
         Begin VB.Frame Frame11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   4560
            TabIndex        =   86
            Top             =   240
            Width           =   1815
            Begin VB.CheckBox chk_ins_Cng 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Inserer Conge"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   90
               Top             =   240
               Width           =   1455
            End
            Begin VB.CheckBox chk_maj_Cng 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj Conge"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   89
               Top             =   480
               Width           =   1455
            End
            Begin VB.CheckBox chk_supp_Cng 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Supprimer Conge"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   88
               Top             =   720
               Width           =   1575
            End
            Begin VB.CheckBox chk_consult_Cng 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consulter Conge"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   87
               Top             =   960
               Width           =   1575
            End
            Begin VB.Label Label15 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Conge"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   91
               Top             =   0
               Width           =   525
            End
         End
         Begin VB.Frame Frame10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   2280
            TabIndex        =   79
            Top             =   1560
            Width           =   2175
            Begin VB.CheckBox chk_consult_PLING 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consulter PLANNING"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   83
               Top             =   960
               Width           =   1935
            End
            Begin VB.CheckBox chk_supp_PLING 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Supprimer PLANNING"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   82
               Top             =   720
               Width           =   1935
            End
            Begin VB.CheckBox chk_maj_PLING 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj PLANNING"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   81
               Top             =   480
               Width           =   1815
            End
            Begin VB.CheckBox chk_ins_PLING 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Inserer PLANNING"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   80
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label14 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "PLANNING"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   84
               Top             =   0
               Width           =   825
            End
         End
         Begin VB.Frame Frame9 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   2280
            TabIndex        =   73
            Top             =   240
            Width           =   2175
            Begin VB.CheckBox chk_ins_PCH 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Inserer Programme"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   77
               Top             =   240
               Width           =   1815
            End
            Begin VB.CheckBox chk_maj_PCH 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj Programme"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   76
               Top             =   480
               Width           =   1815
            End
            Begin VB.CheckBox chk_supp_PCH 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Supprimer Programme"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   75
               Top             =   720
               Width           =   1935
            End
            Begin VB.CheckBox chk_consult_PCH 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consulter Programme"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   74
               Top             =   960
               Width           =   1935
            End
            Begin VB.Label Label13 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Programme Chauffeurs"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   78
               Top             =   0
               Width           =   1965
            End
         End
         Begin VB.Frame Frame21 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   2595
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   2055
            Begin VB.CheckBox ChK_Maj_Cmpt 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj Compteur"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   85
               Top             =   1920
               Width           =   1575
            End
            Begin VB.CheckBox Chk_Consult_Compteurs 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consulter Compteurs"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   72
               Top             =   1680
               Width           =   1815
            End
            Begin VB.CheckBox chk_Consult_sup 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consul En/H Services"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   71
               Top             =   1200
               Width           =   1815
            End
            Begin VB.CheckBox chk_Maj_Dispo 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj En/Hors Service"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   70
               Top             =   1440
               Width           =   1815
            End
            Begin VB.CheckBox chk_consult_FT 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consulter FT"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   68
               Top             =   960
               Width           =   1815
            End
            Begin VB.CheckBox Chk_Ins_FT 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Inserer FT"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   67
               Top             =   240
               Width           =   1815
            End
            Begin VB.CheckBox Chk_Maj_FT 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj FT"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   66
               Top             =   480
               Width           =   1815
            End
            Begin VB.CheckBox chk_supp_FT 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Supprimer FT"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   65
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   " Traffic"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   69
               Top             =   0
               Width           =   585
            End
         End
      End
      Begin VB.Frame Frame12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fichiers des Bases"
         ForeColor       =   &H000000FF&
         Height          =   2895
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   8655
         Begin VB.Frame Frame19 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   6360
            TabIndex        =   57
            Top             =   1560
            Width           =   2175
            Begin VB.CheckBox chk_consult_tv 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consult Type Vidange"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   61
               Top             =   960
               Width           =   1935
            End
            Begin VB.CheckBox chk_ins_tv 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Inserer Type Vidange"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   60
               Top             =   240
               Width           =   1935
            End
            Begin VB.CheckBox chk_maj_tv 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj Type Vidange"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   59
               Top             =   480
               Width           =   1935
            End
            Begin VB.CheckBox chk_supp_tv 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Supp Type Vidange"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   58
               Top             =   720
               Width           =   1935
            End
            Begin VB.Label Label23 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Type Vidange"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   62
               Top             =   0
               Width           =   1140
            End
         End
         Begin VB.Frame Frame18 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   6360
            TabIndex        =   51
            Top             =   240
            Width           =   2175
            Begin VB.CheckBox Chk_supp_tc 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Supp Type carburant"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   55
               Top             =   720
               Width           =   1935
            End
            Begin VB.CheckBox Chk_maj_tc 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj Type Carburant"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   54
               Top             =   480
               Width           =   1935
            End
            Begin VB.CheckBox chk_ins_tc 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Ins.Type Carburant"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   53
               Top             =   240
               Width           =   1935
            End
            Begin VB.CheckBox Chk_consult_tc 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consul Type carburant"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   52
               Top             =   960
               Width           =   1935
            End
            Begin VB.Label Label21 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Type Carburant"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   56
               Top             =   0
               Width           =   1320
            End
         End
         Begin VB.Frame Frame17 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   4320
            TabIndex        =   45
            Top             =   1560
            Width           =   1935
            Begin VB.CheckBox chk_consult_prod 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consulter Produit"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   49
               Top             =   960
               Width           =   1695
            End
            Begin VB.CheckBox chk_ins_prod 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Inserer Produit"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   48
               Top             =   240
               Width           =   1455
            End
            Begin VB.CheckBox chk_maj_prod 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj Produit"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   47
               Top             =   480
               Width           =   1455
            End
            Begin VB.CheckBox chk_supp_prod 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Supprimer Produit"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   46
               Top             =   720
               Width           =   1695
            End
            Begin VB.Label Label20 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Produit"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   50
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame Frame16 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   4320
            TabIndex        =   39
            Top             =   240
            Width           =   1935
            Begin VB.CheckBox chk_consult_fr 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consul Fournisseur"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   960
               Width           =   1695
            End
            Begin VB.CheckBox chk_ins_fr 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Ins. Fournisseur"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   240
               Width           =   1695
            End
            Begin VB.CheckBox chk_maj_fr 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj Fournisseur"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   41
               Top             =   480
               Width           =   1455
            End
            Begin VB.CheckBox chk_supp_fr 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Supp Fournisseur"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   720
               Width           =   1695
            End
            Begin VB.Label Label19 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Fournisseur"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   44
               Top             =   0
               Width           =   990
            End
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   2280
            TabIndex        =   33
            Top             =   1560
            Width           =   1935
            Begin VB.CheckBox chk_consult_dest 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consul Destination"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   960
               Width           =   1695
            End
            Begin VB.CheckBox chk_ins_dest 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Ins  Destination"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Width           =   1455
            End
            Begin VB.CheckBox chk_maj_dest 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj Destination"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   480
               Width           =   1455
            End
            Begin VB.CheckBox chk_supp_dest 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Supp Destination"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   34
               Top             =   720
               Width           =   1695
            End
            Begin VB.Label Label7 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Destination"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   38
               Top             =   0
               Width           =   975
            End
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   2280
            TabIndex        =   27
            Top             =   240
            Width           =   1935
            Begin VB.CheckBox Chk_supp_vh 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Supprimer Vehicule"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   720
               Width           =   1695
            End
            Begin VB.CheckBox chk_maj_vh 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj vehicule"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   480
               Width           =   1455
            End
            Begin VB.CheckBox chk_ins_VH 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Inserer vehicule"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   240
               Width           =   1455
            End
            Begin VB.CheckBox Chk_consult_vh 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consulter Vehicule"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label18 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Véhicule"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   32
               Top             =   0
               Width           =   705
            End
         End
         Begin VB.Frame Frame14 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   120
            TabIndex        =   21
            Top             =   1560
            Width           =   2055
            Begin VB.CheckBox chk_consult_per 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consulter Personnel"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   25
               Top             =   960
               Width           =   1815
            End
            Begin VB.CheckBox chk_ins_per 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Inserer Personnel"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   240
               Width           =   1815
            End
            Begin VB.CheckBox chk_maj_per 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj Personnel"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   480
               Width           =   1815
            End
            Begin VB.CheckBox chk_supp_per 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Supprimer Personnel"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label17 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Personnel"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   26
               Top             =   0
               Width           =   840
            End
         End
         Begin VB.Frame Frame13 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1280
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   2055
            Begin VB.CheckBox chk_supp_user 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Supprimer Utilisateur"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   720
               Width           =   1815
            End
            Begin VB.CheckBox chk_maj_user 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Maj Utilisateur"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   480
               Width           =   1815
            End
            Begin VB.CheckBox chk_ins_user 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Inserer Utilisateur"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Width           =   1815
            End
            Begin VB.CheckBox chk_consult_user 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Consulter Utilisateur"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label Label16 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Utilisateur"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   240
               TabIndex        =   20
               Top             =   0
               Width           =   885
            End
         End
      End
   End
   Begin VB.TextBox txt_Cmp 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   8160
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2640
      Width           =   4095
   End
   Begin VB.TextBox txt_mp 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2640
      Width           =   4095
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   1800
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Frm_Utilisateur.frx":0ECA
      ButtonType      =   1
   End
   Begin VB.TextBox txt_Nom 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "M"
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox txt_Matricule 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   2160
      TabIndex        =   0
      Tag             =   "M"
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Image Cmd_ReAjouter 
      Height          =   375
      Left            =   13320
      Picture         =   "Frm_Utilisateur.frx":121D
      Stretch         =   -1  'True
      Top             =   1890
      Width           =   1575
   End
   Begin VB.Image CmdSave 
      Height          =   495
      Left            =   12360
      Picture         =   "Frm_Utilisateur.frx":12F3F
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   9360
      Picture         =   "Frm_Utilisateur.frx":251C1
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Image CmdFind 
      Height          =   540
      Left            =   10800
      Picture         =   "Frm_Utilisateur.frx":3703B
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Image CmdAdd 
      Height          =   495
      Left            =   7800
      Picture         =   "Frm_Utilisateur.frx":47C3D
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Lbl_RaAjouter 
      BackStyle       =   0  'Transparent
      Caption         =   "Utilisateur est déja supprime, Voulez-vous ré-ajouter?..."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8040
      TabIndex        =   13
      Top             =   1920
      Width           =   5535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fiche utilisateur"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   360
      Width           =   2655
   End
   Begin VB.Image PicBox_Header 
      Height          =   1575
      Left            =   -120
      Picture         =   "Frm_Utilisateur.frx":59D67
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
   Begin VB.Line Line14 
      X1              =   5040
      X2              =   7800
      Y1              =   9360
      Y2              =   9360
   End
   Begin VB.Line Line13 
      X1              =   5040
      X2              =   5040
      Y1              =   8280
      Y2              =   9360
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actif : O/N"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   360
      TabIndex        =   10
      Top             =   3000
      Width           =   1035
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm m.p :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   6360
      TabIndex        =   8
      Top             =   2640
      Width           =   1275
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mot de passe :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   1440
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nom et prénom :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Matricule"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "Frm_Utilisateur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Chk_tous_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim objControl As Control
If Chk_tous.Value = Unchecked Then

    If MsgBox("Vous allez sélectionner tous les droits d'accès!", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
        Chk_tous.Value = Unchecked
        Exit Sub
    Else
            
        For Each objControl In Frm_Utilisateur.Controls
                If TypeOf objControl Is CheckBox Then
                    objControl.Value = Checked
                End If
        Next
    End If
Else

    If MsgBox("Vous allez désélectionner tous les droits d'accès!", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
        Chk_tous.Value = Checked
        Exit Sub
    Else
        For Each objControl In Frm_Utilisateur.Controls
                If TypeOf objControl Is CheckBox Then
                    objControl.Value = Unchecked
                End If
        Next
    End If
End If
End Sub

Private Sub Form_Load()
    Call InitialGrid
    Picture2.Visible = False
    Me.WindowState = 2
    CmdDelete.Enabled = False
    CmdSave.Enabled = False
    Lbl_RaAjouter.Visible = False
    Cmd_ReAjouter.Visible = False
End Sub
Private Sub Form_Resize()
On Error Resume Next
    Dim WidthForm   As Integer
    WidthForm = Me.Width
    PicBox_Header.Width = WidthForm
    CmdSave.Left = WidthForm - 2000
    CmdFind.Left = WidthForm - 3500
    CmdDelete.Left = WidthForm - 5000
    CmdAdd.Left = WidthForm - 6500
End Sub
Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo erreur
'    If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
'        Cancel = True
'    Else
'        Unload Me
'    End If
'Exit Sub
'erreur:
'   MsgBox Err.Description, 48
End Sub

Private Sub SGrid1_DblClick(ByVal lRow As Long, ByVal lCol As Long)
Dim VCode
    On Error GoTo Err
    
     VCode = SGrid1.CellText(lRow, 1)
     cmdFindMatricule.SetFocus
     AfficheRow (VCode)
     
     Picture2.Visible = False
     Exit Sub
Err:
        MsgBox Err.Description, vbInformation

End Sub

Private Sub SGrid1_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    Dim VCode
    On Error GoTo Err
        VCode = SGrid1.CellText(SGrid1.SelectedRow, 1)
        Select Case KeyCode
            Case vbKeyReturn
                cmdFindMatricule.SetFocus
                AfficheRow (VCode)
                Picture2.Visible = False
            Case vbKeyEscape
                Picture2.Visible = False
        End Select
     Exit Sub
Err:
        MsgBox Err.Description, vbInformation

End Sub

Private Sub SGrid1_LostFocus()
    Picture2.Visible = False
End Sub

Private Sub txt_Matricule_GotFocus()
Call ViderZone(Frm_Utilisateur)
End Sub

'ControlBox***
Private Sub txt_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
If IsNumeric(Trim(txt_Matricule.Text)) Then
    If KeyCode = vbKeyReturn Then
        If Len(Trim(txt_Matricule.Text)) > 0 Then Call AfficheRow(txt_Matricule.Text)
    End If
Else
    If KeyCode = vbKeyRight Then
        If Len(Trim(txt_Matricule.Text)) > 0 Then Call AfficheRow(txt_Matricule.Text)
    End If
 End If
End Sub




Private Sub txt_Nom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt_Nom_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° matricule obligatoire      ", vbInformation, App.ProductName
'        txt_Matricule.SetFocus
    End If
End Sub
Private Sub CmdFind_Click()
 On Error Resume Next
    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cours?", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then Exit Sub
    End If
    Unload FrmFind_Actif
    Unload FrmFind
    Unload FrmFind_Fils
    Unload Frm_FindView
    cmdFindMatricule.SetFocus
    With Frm_FindView
        .StrSource = "Utilisateur"
        .Show
    End With
End Sub
Private Sub cmdFindMatricule_Click()
   On Error Resume Next
    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cours?", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then Exit Sub
    End If
    Unload FrmFind_Actif
    Unload FrmFind
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "Utilisateur_actif"
        .Show
    End With
End Sub
'Nouveau***
Private Sub CmdAdd_Click()
On Error GoTo Err
    If (CHECK_ACCES("Ins_Utilisateur", LInt_UserId) = False) Then
        MsgBox "Insertion n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Ajouter utilisateur", vbInformation, App.ProductName
        Exit Sub
    End If
    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then Exit Sub
    End If
    Call ViderZone(Frm_Utilisateur)
    Call EnbDisb(True)
    Cmd_ReAjouter.Visible = False
    Lbl_RaAjouter.Visible = False
    txt_Matricule.Text = "Auto"
    txt_Nom.SetFocus
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
'Afficher Acces Utilisateur***
Public Sub AfficheRow(ByVal VCode As String)
    Dim LObj_Find As New Utilisateur
    Dim Lrs_User As Recordset
    Dim w
    Dim k
    Dim i
On Error GoTo Err
    If (CHECK_ACCES("Consult_Utilisateur", LInt_UserId) = False) Then
        MsgBox "Consultation n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Consulter liste des utilisateurs", vbInformation, App.ProductName
        Exit Sub
    End If
    Set Lrs_User = LObj_Find.GetRow_UserByCode(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
        Exit Sub
    End If
    Set LObj_Find = Nothing
    Call ViderZone(Frm_Utilisateur)
    SGrid1.ClearRows
    If Lrs_User.RecordCount = 1 Then
    
        'Charge
        txt_Matricule.Text = Lrs_User("Code")
        txt_Nom.Text = Lrs_User("NOMPRN")
        w = Lrs_User.Fields("mp").Value
        k = ""
        For i = 1 To Len(w) Step 2
            k = k & Chr(Mid(w, i, 2))
        Next
        txt_mp.Text = k
        txt_Cmp.Text = k
        Chk_Ins_bc.Value = Lrs_User("Ins_BC")
        Chk_Maj_bc.Value = Lrs_User("Maj_BC")
        chk_Supp_BC.Value = Lrs_User("Supp_BC")
        Chk_consult_bc.Value = Lrs_User("Consult_BC")
        chk_ins_bv.Value = Lrs_User("Ins_BV")
        Chk_Maj_BV.Value = Lrs_User("Maj_BV")
        Chk_supp_bv.Value = Lrs_User("Supp_BV")
        Chk_consult_bv.Value = Lrs_User("Consult_BV")
        Chk_consult_alrt.Value = Lrs_User("Consult_Alerte")
        chk_ins_bcr.Value = Lrs_User("Ins_BCR")
        Chk_maj_bcr.Value = Lrs_User("Maj_BCR")
        Chk_supp_bcr.Value = Lrs_User("Supp_BCR")
        Chk_consult_bcr.Value = Lrs_User("Consult_BCR")
        chk_ins_pr.Value = Lrs_User("InS_PR")
        chk_maj_pr.Value = Lrs_User("Maj_PR")
        chk_supp_pr.Value = Lrs_User("Supp_PR")
        Chk_consult_pr.Value = Lrs_User("Consult_PR")
        chk_ins_ff.Value = Lrs_User("Ins_FF")
        chk_maj_ff.Value = Lrs_User("Maj_FF")
        chk_supp_ff.Value = Lrs_User("Supp_FF")
        chk_consult_ff.Value = Lrs_User("Consult_FF")
        chk_SC.Value = Lrs_User("Consult_SC")
        Chk_Ins_FT.Value = Lrs_User("Ins_FT")
        Chk_Maj_FT.Value = Lrs_User("Maj_FT")
        chk_supp_FT.Value = Lrs_User("Supp_FT")
        chk_consult_FT.Value = Lrs_User("Consult_FT")
        chk_Consult_sup.Value = Lrs_User("Consult_SUp")
        chk_ins_VH.Value = Lrs_User("Ins_Vehicule")
        chk_maj_vh.Value = Lrs_User("Maj_vehicule")
        Chk_supp_vh.Value = Lrs_User("Supp_vehicule")
        Chk_consult_vh.Value = Lrs_User("Consult_vehicule")
        chk_ins_fr.Value = Lrs_User("Ins_Fournisseur")
        chk_maj_fr.Value = Lrs_User("Maj_Fournisseur")
        chk_supp_fr.Value = Lrs_User("Supp_Fournisseur")
        chk_consult_fr.Value = Lrs_User("Conslt_Fournisseur")
        chk_ins_tc.Value = Lrs_User("Ins_TC")
        Chk_maj_tc.Value = Lrs_User("Maj_TC")
        Chk_supp_tc.Value = Lrs_User("Supp_TC")
        Chk_consult_tc.Value = Lrs_User("Consult_TC")
        chk_ins_tv.Value = Lrs_User("Ins_TV")
        chk_maj_tv.Value = Lrs_User("Maj_TV")
        chk_supp_tv.Value = Lrs_User("supp_TV")
        chk_consult_tv.Value = Lrs_User("Consult_TV")
        chk_ins_dest.Value = Lrs_User("Ins_Destination")
        chk_maj_dest.Value = Lrs_User("Maj_Destination")
        chk_supp_dest.Value = Lrs_User("Supp_Destination")
        chk_consult_dest.Value = Lrs_User("Consult_Destination")
        chk_ins_prod.Value = Lrs_User("Ins_Produit")
        chk_maj_prod.Value = Lrs_User("Maj_produit")
        chk_supp_prod.Value = Lrs_User("Supp_Produit")
        chk_consult_prod.Value = Lrs_User("Consult_Produit")
        chk_ins_per.Value = Lrs_User("Ins_Personnel")
        chk_maj_per.Value = Lrs_User("Maj_Personnel")
        chk_supp_per.Value = Lrs_User("Supp_personnel")
        chk_consult_per.Value = Lrs_User("Consult_personnel")
        chk_ins_user.Value = Lrs_User("Ins_Utilisateur")
        chk_maj_user.Value = Lrs_User("Maj_Utilisateur")
        chk_supp_user.Value = Lrs_User("Supp_Utilisateur")
        chk_consult_user.Value = Lrs_User("Consult_Utilisateur")
        chk_ins_PCH.Value = Lrs_User("Ins_PCH")
        chk_maj_PCH.Value = Lrs_User("Maj_PCH")
        chk_supp_PCH.Value = Lrs_User("Supp_PCH")
        chk_consult_PCH.Value = Lrs_User("Consult_PCH")
        chk_ins_PLING.Value = Lrs_User("Ins_PLING")
        chk_maj_PLING.Value = Lrs_User("Maj_PLING")
        chk_supp_PLING.Value = Lrs_User("Supp_PLING")
        chk_consult_PLING.Value = Lrs_User("Consult_PLING")
        chk_ins_Cng.Value = Lrs_User("Ins_Conge")
        chk_maj_Cng.Value = Lrs_User("Maj_Conge")
        chk_supp_Cng.Value = Lrs_User("Supp_Conge")
        chk_consult_Cng.Value = Lrs_User("Consult_Conge")
        chk_Actif.Value = Lrs_User("Actif")
        chk_Maj_Dispo.Value = Lrs_User("Maj_Disp")
        Chk_Consult_Compteurs.Value = Lrs_User("Consult_Compteurs")
        ChK_Maj_Cmpt.Value = Lrs_User("Maj_Compt")
        If Lrs_User("Supp") = "N" Then
            EnbDisb (True)
            Lbl_RaAjouter.Visible = False
            Cmd_ReAjouter.Visible = False
        Else
            EnbDisb (False)
            Lbl_RaAjouter.Visible = True
            Cmd_ReAjouter.Visible = True
        End If
    
ElseIf Not Lrs_User.EOF Then
        While Not Lrs_User.EOF
        SGrid1.Redraw = False
        With SGrid1
            .AddRow
            .CellDetails .Rows, .ColumnIndex("Code"), Lrs_User.Fields("Code"), DT_RIGHT
            .CellDetails .Rows, .ColumnIndex("NomPrn"), Lrs_User.Fields("NomPrn"), DT_LEFT
            .CellDetails .Rows, .ColumnIndex("Actif"), Lrs_User.Fields("actif"), DT_RIGHT
        End With
        Lrs_User.MoveNext
        Wend
    SGrid1.Redraw = True
    Picture2.Visible = True
    If SGrid1.Rows > 0 Then
        SGrid1.SelectedRow = 1
        SGrid1.SetFocus
    End If
        
    
    Else
        MsgBox "Code introuvable", vbInformation, App.ProductName
        EnbDisb (True)
        txt_Matricule.SetFocus
    End If
    Set Lrs_User = Nothing
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
'Enabled / Disabled ControlBox***
Private Sub EnbDisb(ByVal TYP As Boolean)
    txt_Matricule.Enabled = TYP
    txt_Nom.Enabled = TYP
    txt_mp.Enabled = TYP
    txt_Cmp.Enabled = TYP
    Chk_Ins_bc.Enabled = TYP
    Chk_Maj_bc.Enabled = TYP
    chk_Supp_BC.Enabled = TYP
    Chk_consult_bc.Enabled = TYP
    chk_ins_bv.Enabled = TYP
    Chk_Maj_BV.Enabled = TYP
    Chk_supp_bv.Enabled = TYP
    Chk_consult_bv.Enabled = TYP
    Chk_consult_alrt.Enabled = TYP
    chk_ins_bcr.Enabled = TYP
    Chk_maj_bcr.Enabled = TYP
    Chk_supp_bcr.Enabled = TYP
    Chk_consult_bcr.Enabled = TYP
    chk_ins_pr.Enabled = TYP
    chk_maj_pr.Enabled = TYP
    chk_supp_pr.Enabled = TYP
    Chk_consult_pr.Enabled = TYP
    chk_ins_ff.Enabled = TYP
    chk_maj_ff.Enabled = TYP
    chk_supp_ff.Enabled = TYP
    chk_consult_ff.Enabled = TYP
    chk_SC.Enabled = TYP
    Chk_Ins_FT.Enabled = TYP
    Chk_Maj_FT.Enabled = TYP
    chk_supp_FT.Enabled = TYP
    chk_consult_FT.Enabled = TYP
    chk_Consult_sup.Enabled = TYP
    chk_ins_VH.Enabled = TYP
    chk_maj_vh.Enabled = TYP
    Chk_supp_vh.Enabled = TYP
    Chk_consult_vh.Enabled = TYP
    chk_ins_fr.Enabled = TYP
    chk_maj_fr.Enabled = TYP
    chk_supp_fr.Enabled = TYP
    chk_consult_fr.Enabled = TYP
    chk_ins_tc.Enabled = TYP
    Chk_maj_tc.Enabled = TYP
    Chk_supp_tc.Enabled = TYP
    Chk_consult_tc.Enabled = TYP
    chk_ins_tv.Enabled = TYP
    chk_maj_tv.Enabled = TYP
    chk_supp_tv.Enabled = TYP
    chk_consult_tv.Enabled = TYP
    chk_ins_dest.Enabled = TYP
    chk_maj_dest.Enabled = TYP
    chk_supp_dest.Enabled = TYP
    chk_consult_dest.Enabled = TYP
    chk_ins_prod.Enabled = TYP
    chk_maj_prod.Enabled = TYP
    chk_supp_prod.Enabled = TYP
    chk_consult_prod.Enabled = TYP
    chk_ins_per.Enabled = TYP
    chk_maj_per.Enabled = TYP
    chk_supp_per.Enabled = TYP
    chk_consult_per.Enabled = TYP
    chk_ins_user.Enabled = TYP
    chk_maj_user.Enabled = TYP
    chk_supp_user.Enabled = TYP
    chk_consult_user.Enabled = TYP
    chk_ins_PCH.Enabled = TYP
    chk_maj_PCH.Enabled = TYP
    chk_supp_PCH.Enabled = TYP
    chk_consult_PCH.Enabled = TYP
    chk_ins_PLING.Enabled = TYP
    chk_maj_PLING.Enabled = TYP
    chk_supp_PLING.Enabled = TYP
    chk_consult_PLING.Enabled = TYP
    chk_ins_Cng.Enabled = TYP
    chk_maj_Cng.Enabled = TYP
    chk_supp_Cng.Enabled = TYP
    chk_consult_Cng.Enabled = TYP
    chk_Actif.Enabled = TYP
    chk_Maj_Dispo.Enabled = TYP
    Chk_Consult_Compteurs.Enabled = TYP
    ChK_Maj_Cmpt.Enabled = TYP
    CmdDelete.Enabled = TYP
    CmdSave.Enabled = TYP
End Sub
'Suppression***
Private Sub CmdDelete_Click()
    Dim LObj_Find   As New Utilisateur
    Dim VCode       As String
On Error GoTo Err
    If (CHECK_ACCES("Supp_Utilisateur", LInt_UserId) = False) Then
        MsgBox "Suppression n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Supprimer des utilisateurs", vbInformation, App.ProductName
        Exit Sub
    End If
    If txt_Matricule.Text <> "Auto" And txt_Matricule.Text <> "" Then
        If MsgBox("Confirmez vous la suppression", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
            VCode = txt_Matricule.Text
            Call LObj_Find.Delete_Add_USER(ErrNumber, ErrDescription, ErrSourceDetail, VCode, "O", LInt_UserId, CNB)
            If ErrNumber <> 0 Then
                ErrNumber = 0
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                Exit Sub
            End If
            Set LObj_Find = Nothing
            MsgBox "Utilisateur Supprimer avec Succes!...", vbInformation, App.ProductName
            Call EnbDisb(True)
            Call ViderZone(Frm_Utilisateur)
            Cmd_ReAjouter.Visible = False
            Lbl_RaAjouter.Visible = False
        End If
    Else
        MsgBox "Séléctionner un 'Utilisateur' puis Supprimer!...", vbExclamation, App.ProductName
    End If
    txt_Matricule.SetFocus
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
'Ré-ajouter***
Private Sub Cmd_ReAjouter_Click()
    Dim LObj_Find   As New Utilisateur
    Dim VCode       As String
On Error GoTo Err
    If (CHECK_ACCES("Supp_Utilisateur", LInt_UserId) = False) Then
        MsgBox "Ré-ajouter n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour ré-ajouter des utilisateurs", vbInformation, App.ProductName
        Exit Sub
    End If
    If txt_Matricule.Text <> "Auto" And txt_Matricule.Text <> "" Then
        If MsgBox("Confirmez vous la ré-ajouter", vbYesNo + vbDefaultButton2 + vbInformation, App.ProductName) = vbYes Then
            VCode = txt_Matricule.Text
            Call LObj_Find.Delete_Add_USER(ErrNumber, ErrDescription, ErrSourceDetail, VCode, "N", LInt_UserId, CNB)
            If ErrNumber <> 0 Then
                ErrNumber = 0
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                Exit Sub
            End If
            Set LObj_Find = Nothing
            MsgBox "Utilisateur ré-ajouter avec Succes!...", vbInformation, App.ProductName
            Call AfficheRow(VCode)
        End If
    Else
        MsgBox "Séléctionner un 'Utilisateur' puis Supprimer!...", vbExclamation, App.ProductName
    End If
    txt_Matricule.SetFocus
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
'Enregistrement***
Private Sub CmdSave_Click()
    Dim LObj_Find           As New Utilisateur
    Dim LInt_NumCompteur    As Long
    Dim VCode               As String
    Dim rs As New ADODB.Recordset
    Dim w
    Dim A
    Dim i
On Error GoTo Err
    If txt_Nom = "" Or txt_Cmp = "" Or txt_Matricule = "" Or txt_mp = "" Then
        MsgBox "Remplir tous le(s) champ(s) d'identite(s)!...", vbInformation, App.ProductName
        Exit Sub
    End If
    If Left(CheckMandatory(Frm_Utilisateur), 1) = 1 Then Exit Sub
    A = UCase(txt_mp.Text)
    w = ""
    For i = 1 To Len(A)
       w = w & Asc(Mid(A, i, 1))
    Next
    '===================
    'Modifier***
    If txt_Matricule.Text <> "Auto" And txt_Matricule.Text <> "" Then
        '===================
        'USER ACCES***
        If (CHECK_ACCES("MAJ_Utilisateur", LInt_UserId) = False) Then
            MsgBox "Modification n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour modifier des utilisateurs", vbInformation, App.ProductName
            Exit Sub
        End If
        If txt_mp.Text = txt_Cmp.Text Then
            If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation, App.ProductName) = vbYes Then
                VCode = txt_Matricule.Text
                Call UpdateUser(VCode, w)
            End If
        Else
        MsgBox ("Accès interdit!")
        End If
        
    '===================
    'Ajouter***
    ElseIf txt_Matricule.Text = "Auto" Then
        '===================
        'USER ACCES***
        If (CHECK_ACCES("Ins_Utilisateur", LInt_UserId) = False) Then
            MsgBox "Insertion n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Ajouter des utilisateurs", vbInformation, App.ProductName
            Exit Sub
        End If
        
        If txt_mp.Text = txt_Cmp.Text Then
            Set rs = LObj_Find.GetRow_UsersByPwd(ErrNumber, ErrDescription, ErrSourceDetail, w, CNB)
                If ErrNumber <> 0 Then
                        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                        ErrNumber = 0
                        Exit Sub
                End If
                If Not rs.EOF Then
                    MsgBox ("Ce mot de passe existe déjà!")
                    Exit Sub
                Else
                        If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation, App.ProductName) = vbYes Then
                            LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "F_Utilisateur")
                                If ErrNumber <> 0 Then
                                   MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion, App.ProductName
                                   ErrNumber = 0
                                   Exit Sub
                                End If
                         
                            VCode = Format(LInt_NumCompteur, "00000")
                        
                            Call SaveUser(VCode, w)
                        End If
                End If
        Else
            MsgBox ("Mot de passe non conforme!")
            Exit Sub
        End If
        End If
        
    
    Call ViderZone(Frm_Utilisateur)
    Call EnbDisb(True)
    Cmd_ReAjouter.Visible = False
    Lbl_RaAjouter.Visible = False
    CmdSave.Enabled = False
    CmdDelete.Enabled = False
    txt_Matricule.SetFocus
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
'=========================================
'Subilation Modifier / Ajouter***
'=========================================
Private Sub UpdateUser(ByVal Code As String, ByVal MP As String)
    Dim Lobj_Save   As New Utilisateur
    Dim Lrs_User    As New Recordset
    Set Lrs_User = CreateEmptyRS_USER()
    With Lrs_User
        .AddNew
        .Fields("MP") = MP
        .Fields("NomPrn") = txt_Nom.Text
        .Fields("Ins_BC") = Chk_Ins_bc.Value
        .Fields("Maj_BC") = Chk_Maj_bc.Value
        .Fields("Supp_BC") = chk_Supp_BC.Value
        .Fields("Consult_BC") = Chk_consult_bc.Value
        .Fields("Ins_BV") = chk_ins_bv.Value
        .Fields("Maj_BV") = Chk_Maj_BV.Value
        .Fields("Supp_BV") = Chk_supp_bv.Value
        .Fields("Consult_BV") = Chk_consult_bv.Value
        .Fields("Consult_Alerte") = Chk_consult_alrt.Value
        .Fields("Ins_BCR") = chk_ins_bcr.Value
        .Fields("Maj_BCR") = Chk_maj_bcr.Value
        .Fields("Supp_BCR") = Chk_supp_bcr.Value
        .Fields("Consult_BCR") = Chk_consult_bcr.Value
        .Fields("Ins_PR") = chk_ins_pr.Value
        .Fields("Maj_PR") = chk_maj_pr.Value
        .Fields("Supp_PR") = chk_supp_pr.Value
        .Fields("Consult_PR") = Chk_consult_pr.Value
        .Fields("Ins_FF") = chk_ins_ff.Value
        .Fields("Maj_FF") = chk_maj_ff.Value
        .Fields("Supp_FF") = chk_supp_ff.Value
        .Fields("Consult_FF") = chk_consult_ff.Value
        .Fields("Consult_SC") = chk_SC.Value
        .Fields("Ins_FT") = Chk_Ins_FT.Value
        .Fields("Maj_FT") = Chk_Maj_FT.Value
        .Fields("Supp_FT") = chk_supp_FT.Value
        .Fields("Consult_FT") = chk_consult_FT.Value
        .Fields("Consult_Sup") = chk_Consult_sup.Value
        .Fields("Ins_Vehicule") = chk_ins_VH.Value
        .Fields("Maj_vehicule") = chk_maj_vh.Value
        .Fields("Supp_vehicule") = Chk_supp_vh.Value
        .Fields("Consult_vehicule") = Chk_consult_vh.Value
        .Fields("Ins_Fournisseur") = chk_ins_fr.Value
        .Fields("Maj_Fournisseur") = chk_maj_fr.Value
        .Fields("Supp_Fournisseur") = chk_supp_fr.Value
        .Fields("Conslt_Fournisseur") = chk_consult_fr.Value
        .Fields("Ins_TC") = chk_ins_tc.Value
        .Fields("Maj_TC") = Chk_maj_tc.Value
        .Fields("Supp_TC") = Chk_supp_tc.Value
        .Fields("Consult_TC") = Chk_consult_tc.Value
        .Fields("Ins_TV") = chk_ins_tv.Value
        .Fields("Maj_TV") = chk_maj_tv.Value
        .Fields("supp_TV") = chk_supp_tv.Value
        .Fields("Consult_TV") = chk_consult_tv.Value
        .Fields("Ins_Destination") = chk_ins_dest.Value
        .Fields("Maj_Destination") = chk_maj_dest.Value
        .Fields("Supp_Destination") = chk_supp_dest.Value
        .Fields("Consult_Destination") = chk_consult_dest.Value
        .Fields("Ins_Produit") = chk_ins_prod.Value
        .Fields("Maj_produit") = chk_maj_prod.Value
        .Fields("Supp_Produit") = chk_supp_prod.Value
        .Fields("Consult_Produit") = chk_consult_prod.Value
        .Fields("Ins_Personnel") = chk_ins_per.Value
        .Fields("Maj_Personnel") = chk_maj_per.Value
        .Fields("Supp_personnel") = chk_supp_per.Value
        .Fields("Consult_personnel") = chk_consult_per.Value
        .Fields("Ins_Utilisateur") = chk_ins_user.Value
        .Fields("Maj_Utilisateur") = chk_maj_user.Value
        .Fields("Supp_Utilisateur") = chk_supp_user.Value
        .Fields("Consult_Utilisateur") = chk_consult_user.Value
        .Fields("Actif") = chk_Actif.Value
        .Fields("Maj_Disp") = chk_Maj_Dispo.Value
        .Fields("Maj_Compt") = ChK_Maj_Cmpt.Value
        .Fields("Consult_Compteurs") = Chk_Consult_Compteurs.Value
        .Fields("Ins_PCH") = chk_ins_PCH.Value
        .Fields("Maj_PCH") = chk_maj_PCH.Value
        .Fields("Supp_PCH") = chk_supp_PCH.Value
        .Fields("Consult_PCH") = chk_consult_PCH.Value
        .Fields("Ins_PLING") = chk_ins_PLING.Value
        .Fields("Maj_PLING") = chk_maj_PLING.Value
        .Fields("Supp_PLING") = chk_supp_PLING.Value
        .Fields("Consult_PLING") = chk_consult_PLING.Value
        .Fields("Ins_Conge") = chk_ins_Cng.Value
        .Fields("Maj_Conge") = chk_maj_Cng.Value
        .Fields("Supp_Conge") = chk_supp_Cng.Value
        .Fields("Consult_Conge") = chk_consult_Cng.Value
        .Fields("UserUpdate") = LInt_UserId
    End With
    Call Lobj_Save.UpDate_USER(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_User, Code)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
        Exit Sub
    End If
    Set Lobj_Save = Nothing
    MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
End Sub
Private Sub SaveUser(ByVal Code As String, ByVal MP As String)
    Dim Lobj_Save As New Utilisateur
    Dim Lrs_User As New Recordset
    Set Lrs_User = CreateEmptyRS_USER()
    With Lrs_User
        .AddNew
        .Fields("code") = Code
        .Fields("MP") = MP
        .Fields("NomPrn") = txt_Nom.Text
        .Fields("Ins_BC") = Chk_Ins_bc.Value
        .Fields("Maj_BC") = Chk_Maj_bc.Value
        .Fields("Supp_BC") = chk_Supp_BC.Value
        .Fields("Consult_BC") = Chk_consult_bc.Value
        .Fields("Ins_BV") = chk_ins_bv.Value
        .Fields("Maj_BV") = Chk_Maj_BV.Value
        .Fields("Supp_BV") = Chk_supp_bv.Value
        .Fields("Consult_BV") = Chk_consult_bv.Value
        .Fields("Consult_Alerte") = Chk_consult_alrt.Value
        .Fields("Ins_BCR") = chk_ins_bcr.Value
        .Fields("Maj_BCR") = Chk_maj_bcr.Value
        .Fields("Supp_BCR") = Chk_supp_bcr.Value
        .Fields("Consult_BCR") = Chk_consult_bcr.Value
        .Fields("Ins_PR") = chk_ins_pr.Value
        .Fields("Maj_PR") = chk_maj_pr.Value
        .Fields("Supp_PR") = chk_supp_pr.Value
        .Fields("Consult_PR") = Chk_consult_pr.Value
        .Fields("Ins_FF") = chk_ins_ff.Value
        .Fields("Maj_FF") = chk_maj_ff.Value
        .Fields("Supp_FF") = chk_supp_ff.Value
        .Fields("Consult_FF") = chk_consult_ff.Value
        .Fields("Consult_SC") = chk_SC.Value
        .Fields("Ins_FT") = Chk_Ins_FT.Value
        .Fields("Maj_FT") = Chk_Maj_FT.Value
        .Fields("Supp_FT") = chk_supp_FT.Value
        .Fields("Consult_FT") = chk_consult_FT.Value
        .Fields("Consult_Sup") = chk_Consult_sup.Value
        .Fields("Ins_Vehicule") = chk_ins_VH.Value
        .Fields("Maj_vehicule") = chk_maj_vh.Value
        .Fields("Supp_vehicule") = Chk_supp_vh.Value
        .Fields("Consult_vehicule") = Chk_consult_vh.Value
        .Fields("Ins_Fournisseur") = chk_ins_fr.Value
        .Fields("Maj_Fournisseur") = chk_maj_fr.Value
        .Fields("Supp_Fournisseur") = chk_supp_fr.Value
        .Fields("Conslt_Fournisseur") = chk_consult_fr.Value
        .Fields("Ins_TC") = chk_ins_tc.Value
        .Fields("Maj_TC") = Chk_maj_tc.Value
        .Fields("Supp_TC") = Chk_supp_tc.Value
        .Fields("Consult_TC") = Chk_consult_tc.Value
        .Fields("Ins_TV") = chk_ins_tv.Value
        .Fields("Maj_TV") = chk_maj_tv.Value
        .Fields("supp_TV") = chk_supp_tv.Value
        .Fields("Consult_TV") = chk_consult_tv.Value
        .Fields("Ins_Destination") = chk_ins_dest.Value
        .Fields("Maj_Destination") = chk_maj_dest.Value
        .Fields("Supp_Destination") = chk_supp_dest.Value
        .Fields("Consult_Destination") = chk_consult_dest.Value
        .Fields("Ins_Produit") = chk_ins_prod.Value
        .Fields("Maj_produit") = chk_maj_prod.Value
        .Fields("Supp_Produit") = chk_supp_prod.Value
        .Fields("Consult_Produit") = chk_consult_prod.Value
        .Fields("Ins_Personnel") = chk_ins_per.Value
        .Fields("Maj_Personnel") = chk_maj_per.Value
        .Fields("Supp_personnel") = chk_supp_per.Value
        .Fields("Consult_personnel") = chk_consult_per.Value
        .Fields("Ins_Utilisateur") = chk_ins_user.Value
        .Fields("Maj_Utilisateur") = chk_maj_user.Value
        .Fields("Supp_Utilisateur") = chk_supp_user.Value
        .Fields("Consult_Utilisateur") = chk_consult_user.Value
        .Fields("Actif") = chk_Actif.Value
        .Fields("Maj_Disp") = chk_Maj_Dispo.Value
        .Fields("Maj_Compt") = ChK_Maj_Cmpt.Value
        .Fields("Consult_Compteurs") = Chk_Consult_Compteurs.Value
        .Fields("Ins_PCH") = chk_ins_PCH.Value
        .Fields("Maj_PCH") = chk_maj_PCH.Value
        .Fields("Supp_PCH") = chk_supp_PCH.Value
        .Fields("Consult_PCH") = chk_consult_PCH.Value
        .Fields("Ins_PLING") = chk_ins_PLING.Value
        .Fields("Maj_PLING") = chk_maj_PLING.Value
        .Fields("Supp_PLING") = chk_supp_PLING.Value
        .Fields("Consult_PLING") = chk_consult_PLING.Value
        .Fields("Ins_Conge") = chk_ins_Cng.Value
        .Fields("Maj_Conge") = chk_maj_Cng.Value
        .Fields("Supp_Conge") = chk_supp_Cng.Value
        .Fields("Consult_Conge") = chk_consult_Cng.Value
        .Fields("UserInsert") = LInt_UserId
    End With
    Call Lobj_Save.Save_USER(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_User)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
        Exit Sub
    End If
    Set Lobj_Save = Nothing
    MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
End Sub

Public Sub InitialGrid()
 With SGrid1

    .HideGroupingBox = True
    .AllowGrouping = True

    .GroupRowBackColor = vbWindowBackground
    .GroupRowForeColor = vbWindowText
    
    .GridLineColor = vbWindowBackground
    .GridFillLineColor = vbWindowBackground
    .GridLines = True
    
    .SelectionAlphaBlend = True
    .SelectionOutline = True
    .DrawFocusRectangle = False
    
    .AddColumn "Code", "Matricule", , , 60, , , , , , , CCLSortNumeric
    .AddColumn "NomPrn", "Nom et Prenom", , , 140
    .AddColumn "Actif", "Actif", , , 60
  
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
End With
End Sub
