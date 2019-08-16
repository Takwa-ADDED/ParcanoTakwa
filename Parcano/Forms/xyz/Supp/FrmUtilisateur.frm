VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmUtilisateur 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Parcano"
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmUtilisateur.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   15240
   Begin VB.CheckBox chk_maj_per 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Maj Personnel"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8160
      TabIndex        =   70
      Top             =   8640
      Width           =   2175
   End
   Begin VB.CheckBox chk_supp_per 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Supprimer Personnel"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8160
      TabIndex        =   69
      Top             =   8880
      Width           =   1935
   End
   Begin VB.CheckBox chk_consult_per 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Consulter Personnel"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8160
      TabIndex        =   68
      Top             =   9120
      Width           =   1935
   End
   Begin VB.CheckBox chk_maj_prod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Maj Produit"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   67
      Top             =   8640
      Width           =   2055
   End
   Begin VB.CheckBox chk_supp_prod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Supprimer Produit"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   66
      Top             =   8880
      Width           =   2055
   End
   Begin VB.CheckBox chk_consult_prod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Consulter Produit"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   65
      Top             =   9120
      Width           =   1935
   End
   Begin VB.CheckBox chk_consult_lub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Consulter Lubrifiant"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   64
      Top             =   9120
      Width           =   1935
   End
   Begin VB.CheckBox chk_supp_lub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Supprimer Lubrifiant"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   63
      Top             =   8880
      Width           =   1935
   End
   Begin VB.CheckBox chk_maj_lub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Maj Lubrifiant"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   62
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CheckBox chk_maj_dest 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Maj Destination"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   61
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CheckBox chk_supp_dest 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Supprimer Destination"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   60
      Top             =   8880
      Width           =   1935
   End
   Begin VB.CheckBox chk_consult_dest 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Consulter Destination"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   59
      Top             =   9120
      Width           =   1935
   End
   Begin VB.CheckBox chk_Consult_sup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Consulter En/Hors Service"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10320
      TabIndex        =   58
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CheckBox chk_consult_FT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Consulter FT"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10320
      TabIndex        =   57
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CheckBox chk_supp_FT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Supprimer FT"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10320
      TabIndex        =   47
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CheckBox chk_Actif 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   46
      Top             =   3000
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Liste des accées"
      ForeColor       =   &H000040C0&
      Height          =   7575
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   15135
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   12600
         TabIndex        =   98
         Top             =   3360
         Width           =   2175
         Begin VB.CheckBox chk_consult_PCH 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consulter Programme"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CheckBox chk_supp_PCH 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Supprimer Programme"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   101
            Top             =   960
            Width           =   1935
         End
         Begin VB.CheckBox chk_maj_PCH 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Maj Programme"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   600
            Width           =   1815
         End
         Begin VB.CheckBox chk_ins_PCH 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Inserer Programme"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Programme Chauffeurs"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   240
            TabIndex        =   102
            Top             =   0
            Width           =   1665
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   12600
         TabIndex        =   71
         Top             =   240
         Width           =   2175
         Begin VB.CheckBox chk_consult_user 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consulter Utilisateur"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   104
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CheckBox chk_ins_user 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Inserer Utilisateur"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox chk_maj_user 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Maj Utilisateur"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   600
            Width           =   1815
         End
         Begin VB.CheckBox chk_supp_user 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Supprimer Utilisateur"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Utilisateur"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   240
            TabIndex        =   73
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   7560
         TabIndex        =   51
         Top             =   240
         Width           =   2415
         Begin VB.CheckBox chk_SC 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consulter stat Carburant"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   240
            TabIndex        =   55
            Top             =   360
            Width           =   2055
         End
         Begin VB.CheckBox chk_SR 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consulter Stat Reparation"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   240
            TabIndex        =   54
            Top             =   840
            Width           =   1935
         End
         Begin VB.CheckBox chk_EHS 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consulter stat service"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   1680
            Width           =   2055
         End
         Begin VB.CheckBox chk_ST 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consult Stat Traffic"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Statistiques"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   240
            TabIndex        =   56
            Top             =   0
            Width           =   840
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   5040
         TabIndex        =   48
         Top             =   240
         Width           =   2415
         Begin VB.CheckBox chk_ins_ff 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Inserer Facture"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   360
            Width           =   2055
         End
         Begin VB.CheckBox chk_consult_ff 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consulter Fature"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CheckBox chk_maj_ff 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Maj Facture"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   720
            Width           =   2055
         End
         Begin VB.CheckBox chk_supp_ff 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Supprimer Fature"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Facture Fournisseur"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   240
            TabIndex        =   50
            Top             =   0
            Width           =   1440
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   10080
         TabIndex        =   41
         Top             =   240
         Width           =   2415
         Begin VB.CheckBox Chk_Consult_Compteurs 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consulter Compteurs"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   2520
            Width           =   2175
         End
         Begin VB.CheckBox chk_Maj_Dispo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Maj En/Hors Service"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   2160
            Width           =   2175
         End
         Begin VB.CheckBox Chk_Maj_FT 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Maj FT"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   2175
         End
         Begin VB.CheckBox Chk_Ins_FT 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Inserer FT"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   " Traffic"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   240
            TabIndex        =   44
            Top             =   0
            Width           =   510
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   2400
         TabIndex        =   31
         Top             =   240
         Width           =   2415
         Begin VB.CheckBox chk_ins_pr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Insérer  PR"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   1440
            Width           =   2055
         End
         Begin VB.CheckBox chk_ins_bcr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Inserer B.Cmd.R"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   360
            Width           =   2055
         End
         Begin VB.CheckBox chk_maj_pr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Maj PR"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   1680
            Width           =   2055
         End
         Begin VB.CheckBox Chk_consult_pr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consulter PR"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   2160
            Width           =   1935
         End
         Begin VB.CheckBox chk_supp_pr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Supprimer PR"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   1920
            Width           =   1935
         End
         Begin VB.CheckBox Chk_consult_bcr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consulter B.Cmd.R"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   1080
            Width           =   2055
         End
         Begin VB.CheckBox Chk_supp_bcr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Supprimer B.Cmd.R"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   840
            Width           =   1935
         End
         Begin VB.CheckBox Chk_maj_bcr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Maj B.Cmd.R"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Réparation"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   240
            TabIndex        =   37
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2295
         Begin VB.CheckBox chk_ins_bv 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Inserer Bon Vidange"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   78
            Top             =   1440
            Width           =   1935
         End
         Begin VB.CheckBox chk_Supp_BC 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Supprimer BC"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   74
            Top             =   840
            Width           =   1935
         End
         Begin VB.CheckBox Chk_supp_bv 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Supprimer BV"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   1920
            Width           =   1935
         End
         Begin VB.CheckBox Chk_Ins_bc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Insérer Bon Carburant"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox Chk_Maj_bc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Maj Bon Carburant"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   600
            Width           =   1935
         End
         Begin VB.CheckBox Chk_consult_bc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consulter BC"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox Chk_Maj_BV 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Maj Bon Vidange"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   1680
            Width           =   1935
         End
         Begin VB.CheckBox Chk_consult_bv 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consulter BV"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   2160
            Width           =   1935
         End
         Begin VB.CheckBox Chk_consult_alrt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consultation G.Alerte"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   2520
            Width           =   1935
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "BC | BV | Alerte"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   360
            TabIndex        =   30
            Top             =   0
            Width           =   1110
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   120
         TabIndex        =   14
         Top             =   3240
         Width           =   12375
         Begin VB.CheckBox chk_consult_Superv 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consulte Supervision"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   10560
            TabIndex        =   106
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CheckBox ChK_Maj_Cmpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Maj Compteur"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   10560
            TabIndex        =   96
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chk_consult_tc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consulter Type carburant"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5040
            TabIndex        =   94
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CheckBox chk_maj_vh 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Majvehicule"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   92
            Top             =   600
            Width           =   1935
         End
         Begin VB.CheckBox chk_supp_fr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Supprimer Fournisseur"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   91
            Top             =   840
            Width           =   2295
         End
         Begin VB.CheckBox chk_maj_fr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Maj Fournisseur"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   90
            Top             =   600
            Width           =   2295
         End
         Begin VB.CheckBox chk_ins_VH 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Inserer vehicule"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   89
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox chk_ins_per 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Inserer Personnel"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7920
            TabIndex        =   88
            Top             =   1800
            Width           =   2175
         End
         Begin VB.CheckBox chk_ins_prod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Inserer Produit"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5040
            TabIndex        =   87
            Top             =   1800
            Width           =   2055
         End
         Begin VB.CheckBox chk_ins_tv 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Inserer Type Vidange"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7920
            TabIndex        =   86
            Top             =   360
            Width           =   2415
         End
         Begin VB.CheckBox chk_ins_tc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "InsererType Carburant"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5040
            TabIndex        =   85
            Top             =   360
            Width           =   2655
         End
         Begin VB.CheckBox chk_ins_fr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Inserer Fournisseur"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   84
            Top             =   360
            Width           =   2295
         End
         Begin VB.CheckBox chk_ins_lub 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Ins Lubrifiant"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   83
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CheckBox chk_ins_dest 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Ins  Destination"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   82
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CheckBox chk_consult_tv 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consult Type Vidange"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7920
            TabIndex        =   40
            Top             =   1080
            Width           =   2295
         End
         Begin VB.CheckBox chk_supp_tv 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Supp Type Vidange"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7920
            TabIndex        =   22
            Top             =   840
            Width           =   2055
         End
         Begin VB.CheckBox Chk_supp_tc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Supprimer Type carburant"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5040
            TabIndex        =   21
            Top             =   840
            Width           =   2535
         End
         Begin VB.CheckBox chk_maj_tv 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Maj Type Vidange"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7920
            TabIndex        =   20
            Top             =   600
            Width           =   2415
         End
         Begin VB.CheckBox Chk_maj_tc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Maj Type Carburant"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5040
            TabIndex        =   19
            Top             =   600
            Width           =   2655
         End
         Begin VB.CheckBox chk_consult_fr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consulter Fournisseur"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   18
            Top             =   1080
            Width           =   2175
         End
         Begin VB.CheckBox Chk_consult_vh 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Consulter Vehicule"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox Chk_supp_vh 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Supprimer Vehicule"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   840
            Width           =   1935
         End
         Begin VB.Line Line22 
            X1              =   10440
            X2              =   10440
            Y1              =   1680
            Y2              =   2880
         End
         Begin VB.Line Line21 
            X1              =   10440
            X2              =   12240
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Line20 
            X1              =   10440
            X2              =   12240
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line19 
            X1              =   10440
            X2              =   10440
            Y1              =   240
            Y2              =   1440
         End
         Begin VB.Line Line18 
            X1              =   10200
            X2              =   7800
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Line17 
            X1              =   7800
            X2              =   7800
            Y1              =   1800
            Y2              =   2880
         End
         Begin VB.Line Line16 
            X1              =   7680
            X2              =   4920
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Line15 
            X1              =   4920
            X2              =   4920
            Y1              =   1800
            Y2              =   2880
         End
         Begin VB.Line Line12 
            X1              =   2400
            X2              =   4800
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Line11 
            X1              =   2400
            X2              =   2400
            Y1              =   1680
            Y2              =   2880
         End
         Begin VB.Line Line10 
            X1              =   2280
            X2              =   120
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Line9 
            X1              =   120
            X2              =   120
            Y1              =   1800
            Y2              =   2880
         End
         Begin VB.Line Line8 
            X1              =   7800
            X2              =   7800
            Y1              =   240
            Y2              =   1440
         End
         Begin VB.Line Line7 
            X1              =   120
            X2              =   120
            Y1              =   360
            Y2              =   1440
         End
         Begin VB.Line Line6 
            X1              =   4920
            X2              =   4920
            Y1              =   360
            Y2              =   1440
         End
         Begin VB.Line Line5 
            X1              =   2280
            X2              =   2280
            Y1              =   360
            Y2              =   1440
         End
         Begin VB.Line Line4 
            X1              =   7800
            X2              =   10200
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line3 
            X1              =   4920
            X2              =   7680
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line2 
            X1              =   2280
            X2              =   4800
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   2160
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Fichiers de base"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   360
            TabIndex        =   15
            Top             =   0
            Width           =   1155
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
      TabIndex        =   8
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
      Picture         =   "FrmUtilisateur.frx":0ECA
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
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   14640
      TabIndex        =   4
      Top             =   120
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
      Picture         =   "FrmUtilisateur.frx":121D
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   13800
      TabIndex        =   5
      Top             =   120
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
      Picture         =   "FrmUtilisateur.frx":139F
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   14160
      TabIndex        =   6
      Top             =   120
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
      Picture         =   "FrmUtilisateur.frx":16F2
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   13440
      TabIndex        =   12
      Top             =   120
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
      Picture         =   "FrmUtilisateur.frx":1A45
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
      TabIndex        =   105
      Top             =   360
      Width           =   2655
   End
   Begin VB.Image PicBox_Header 
      Height          =   1575
      Left            =   -120
      Picture         =   "FrmUtilisateur.frx":1BC7
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
      TabIndex        =   45
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "FrmUtilisateur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAdd_Click()

On Error GoTo Err
Dim rQ As New ADODB.Recordset
Dim SQL As String

SQL = "Select * from utilisateur where Ins_Utilisateur = 1 and code= " & LInt_UserId
rQ.Open SQL, CNB, adOpenDynamic
If rQ.EOF Then
    rQ.Close
    MsgBox "Accès refusé.", vbExclamation
    Exit Sub
End If

If txt_Matricule.Text = "Auto" Then
    If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Call ViderZone(FrmUtilisateur)
'txt_Matricule.Enabled = False
txt_Matricule.Text = "Auto"
txt_Nom.SetFocus

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdDelete_Click()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim vcode As String

On Error GoTo Err
   If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        Else
            txt_Matricule.SetFocus
            Exit Sub
        End If
    End If
    If txt_Matricule.Text <> "Auto" Then
        Dim rQ As New ADODB.Recordset
        SQL = "Select * from utilisateur where MAJ_Utilisateur= 1 and code= " & LInt_UserId
        rQ.Open SQL, CNB, adOpenDynamic
        If rQ.EOF Then
            rQ.Close
            MsgBox "Accès refusé.", vbExclamation
            Exit Sub
        End If
    End If

    If MsgBox("Confirmez vous la suppression", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    vcode = txt_Matricule.Text
    SQL = "Delete from uTILISATEUR where code =" & SQLText(vcode)
    CNB.Execute SQL
    txt_Matricule.SetFocus
    End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdFind_Click()
On Error Resume Next
If txt_Matricule.Text = "Auto" Then
    If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Unload FrmFind
With FrmFind
    .StrSource = "Utilisateur"
    .Show
End With
End Sub

Private Sub cmdFindMatricule_Click()
Unload FrmFind_Actif
With FrmFind_Actif
    .StrSource = "Utilisateur"
    .Show
End With
End Sub

Private Sub CmdSave_Click()

Dim SQL As String
Dim vcode As String
Dim w
Dim A

On Error GoTo Err
    If txt_Matricule.Text <> "Auto" Then
        Dim rQ As New ADODB.Recordset
        SQL = "Select * from utilisateur where MAJ_Utilisateur= 1 and code= " & LInt_UserId
        rQ.Open SQL, CNB, adOpenDynamic
        If rQ.EOF Then
            rQ.Close
            MsgBox "Accès refusé.", vbExclamation
            Exit Sub
        End If
    End If
    
    If Left(CheckMandatory(FrmUtilisateur), 1) = 1 Then
       Exit Sub
    End If
    
    If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    'Delete l'enregistrement

    CNB.BeginTrans

    vcode = txt_Matricule.Text
    SQL = "Delete from Utilisateur where code =" & SQLText(vcode)
    CNB.Execute SQL
    
    If vcode = "Auto" Then
    LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "F_Utilisateur")
    If ErrNumber <> 0 Then
       MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
       ErrNumber = 0
       Exit Sub
    End If
    Set LObj_Compteur = Nothing
    'Insertion enregistrement assiette
    txt_Matricule.Text = Format(LInt_NumCompteur, "00000")
    End If
    
    A = UCase(txt_mp.Text)
    w = ""
    For i = 1 To Len(A)
       w = w & Asc(Mid(A, i, 1))
    Next
        
    'Insertion enregistrement
    SQL = "Insert into Utilisateur  ("
    SQL = SQL & " Code,MP,NomPrn,Ins_BC , Maj_BC, Supp_BC, Consult_BC, Ins_BV, Maj_BV, Supp_BV, Consult_BV,"
    SQL = SQL & " Consult_Alerte , Ins_BCR, Maj_BCR, Supp_BCR, Consult_BCR, Ins_PR,"
    SQL = SQL & " Maj_PR , Supp_PR, Consult_PR, Ins_FF, Maj_FF, Supp_FF, Consult_FF, Consult_SC,"
    SQL = SQL & " Consult_SR , Consult_ST, Consult_EHS, Ins_FT, Maj_FT, Supp_FT, Consult_FT,"
    SQL = SQL & " Consult_Sup , Ins_Vehicule, Maj_vehicule, Supp_vehicule, Consult_vehicule,"
    SQL = SQL & " Ins_Fournisseur , Maj_Fournisseur, Supp_Fournisseur, Conslt_Fournisseur,"
    SQL = SQL & " Ins_TC , Maj_TC, Supp_TC, Consult_TC, Ins_TV, Maj_TV, supp_TV,"
    SQL = SQL & " Consult_TV , Ins_Destination, Maj_Destination, Supp_Destination,"
    SQL = SQL & " Consult_Destination , Ins_Lub, Maj_Lub, Supp_Lub, Consult_Lub,"
    SQL = SQL & " Ins_Produit , Maj_produit, Supp_Produit, Consult_Produit, Ins_Personnel,"
    SQL = SQL & " Maj_Personnel , Supp_personnel, Consult_personnel, Ins_Utilisateur,"
    SQL = SQL & " Maj_Utilisateur , Supp_Utilisateur, Consult_Utilisateur, Actif, Maj_Disp, Maj_Compt, Consult_Compteurs,"
    SQL = SQL & " Ins_PCH,Maj_PCH , Supp_PCH, Consult_PCH," 'Consult_Supervision"
    SQL = SQL & " )Values ("
    SQL = SQL & SQLText(txt_Matricule.Text)
    SQL = SQL & "," & SQLText(w)
    SQL = SQL & "," & SQLText(txt_Nom.Text)
    
    SQL = SQL & "," & Chk_Ins_bc.Value
    SQL = SQL & "," & Chk_Maj_bc.Value
    SQL = SQL & "," & chk_Supp_BC.Value
    SQL = SQL & "," & Chk_consult_bc.Value
    
    SQL = SQL & "," & chk_ins_bv.Value
    SQL = SQL & "," & Chk_Maj_BV.Value
    SQL = SQL & "," & Chk_supp_bv.Value
    SQL = SQL & "," & Chk_consult_bv.Value
    
    SQL = SQL & "," & Chk_consult_alrt.Value
    
    SQL = SQL & "," & chk_ins_bcr.Value
    SQL = SQL & "," & Chk_maj_bcr.Value
    SQL = SQL & "," & Chk_supp_bcr.Value
    SQL = SQL & "," & Chk_consult_bcr.Value
    
    SQL = SQL & "," & chk_ins_pr.Value
    SQL = SQL & "," & chk_maj_pr.Value
    SQL = SQL & "," & chk_supp_pr.Value
    SQL = SQL & "," & Chk_consult_pr.Value
    
    SQL = SQL & "," & chk_ins_ff.Value
    SQL = SQL & "," & chk_maj_ff.Value
    SQL = SQL & "," & chk_supp_ff.Value
    SQL = SQL & "," & chk_consult_ff.Value
    
    SQL = SQL & "," & chk_SC.Value
    SQL = SQL & "," & chk_SR.Value
    SQL = SQL & "," & chk_ST.Value
    SQL = SQL & "," & chk_EHS.Value
    
    SQL = SQL & "," & Chk_Ins_FT.Value
    SQL = SQL & "," & Chk_Maj_FT.Value
    SQL = SQL & "," & chk_supp_FT.Value
    SQL = SQL & "," & chk_consult_FT.Value
    
    SQL = SQL & "," & chk_Consult_sup.Value
    
    SQL = SQL & "," & chk_ins_VH.Value
    SQL = SQL & "," & chk_maj_vh.Value
    SQL = SQL & "," & Chk_supp_vh.Value
    SQL = SQL & "," & Chk_consult_vh.Value
    
    SQL = SQL & "," & chk_ins_fr.Value
    SQL = SQL & "," & chk_maj_fr.Value
    SQL = SQL & "," & chk_supp_fr.Value
    SQL = SQL & "," & chk_consult_fr.Value
    
    SQL = SQL & "," & chk_ins_tc.Value
    SQL = SQL & "," & Chk_maj_tc.Value
    SQL = SQL & "," & Chk_supp_tc.Value
    SQL = SQL & "," & chk_consult_tc.Value
    
    SQL = SQL & "," & chk_ins_tv.Value
    SQL = SQL & "," & chk_maj_tv.Value
    SQL = SQL & "," & chk_supp_tv.Value
    SQL = SQL & "," & chk_consult_tv.Value
    
    SQL = SQL & "," & chk_ins_dest.Value
    SQL = SQL & "," & chk_maj_dest.Value
    SQL = SQL & "," & chk_supp_dest.Value
    SQL = SQL & "," & chk_consult_dest.Value
    
    SQL = SQL & "," & chk_ins_lub.Value
    SQL = SQL & "," & chk_maj_lub.Value
    SQL = SQL & "," & chk_supp_lub.Value
    SQL = SQL & "," & chk_consult_lub.Value
    
    SQL = SQL & "," & chk_ins_prod.Value
    SQL = SQL & "," & chk_maj_prod.Value
    SQL = SQL & "," & chk_supp_prod.Value
    SQL = SQL & "," & chk_consult_prod.Value
    
    SQL = SQL & "," & chk_ins_per.Value
    SQL = SQL & "," & chk_maj_per.Value
    SQL = SQL & "," & chk_supp_per.Value
    SQL = SQL & "," & chk_consult_per.Value
    
    SQL = SQL & "," & chk_ins_user.Value
    SQL = SQL & "," & chk_maj_user.Value
    SQL = SQL & "," & chk_supp_user.Value
    SQL = SQL & "," & chk_consult_user.Value
    
    SQL = SQL & "," & chk_Actif.Value
    
    SQL = SQL & "," & chk_Maj_Dispo.Value
    
    SQL = SQL & "," & ChK_Maj_Cmpt.Value
    
    SQL = SQL & "," & Chk_Consult_Compteurs.Value
    
    SQL = SQL & "," & chk_ins_PCH.Value
    SQL = SQL & "," & chk_maj_PCH.Value
    SQL = SQL & "," & chk_supp_PCH.Value
    SQL = SQL & "," & chk_consult_PCH.Value
    
    SQL = SQL & ")"
    CNB.Execute SQL
    CNB.CommitTrans
    MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
    txt_Matricule.SetFocus
    End If
    
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub Form_Load()
 
Dim large As Integer
Dim haut As Integer
large = Screen.Width
haut = Screen.Height
Me.Left = 0
Me.Top = 0
Me.Width = large
Me.Height = haut
Me.WindowState = 2
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Dim WidthForm As Integer
    WidthForm = Frm_Main.ACB_Main.Width + 1200
        PicBox_Header.Width = WidthForm - 1000
        CmdAdd.Left = WidthForm - 3500
        CmdDelete.Left = WidthForm - 3100
        CmdFind.Left = WidthForm - 2700
        CmdSave.Left = WidthForm - 2300

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo erreur
   Dim i As Integer
   Dim MSG ' Déclare la variable.
   ' Définit le texte du message.
   MSG = "Voulez-vous vraiment quitter?"
   ' Si l'utilisateur clique sur Non, met fin à l'événement QueryUnload.
   If MsgBox(MSG, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
      Cancel = True
   Else
   Unload Me
   End If
   
   Exit Sub
erreur:
   MsgBox Err.Description, 48
End Sub


Private Sub txt_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Matricule_LostFocus()
Dim SQL As String
Dim rs As New ADODB.Recordset

On Error GoTo Err
If Len(Trim(txt_Matricule.Text)) > 0 Then
SQL = "Select * from personnel where code = " & SQLText(txt_Matricule.Text)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_Nom.Text = rs("Libelle")
    txt_CIN.Text = rs("CIN")
    txt_fonction.Text = rs("Fonction")
    txt_Telephone.Text = rs("telephone")
    txt_mobile.Text = rs("mobile")
    txt_permie.Text = rs("permie")
    cda_DateLivrPermi.Text = rs("datlivr")
    txt_lieuPermi.Text = rs("lieulivr")
'Else
'    MsgBox "Code introuvable", vbInformation
'    txt_Matricule.SetFocus
'    Exit Sub
End If
rs.Close
End If

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub


Private Sub txt_Nom_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_Nom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub


Public Sub AfficheRow(vcode As String)

Dim SQL As String
Dim rs As New ADODB.Recordset
Dim w
Dim k

Call ViderZone(FrmUtilisateur)

SQL = "Select * from utilisateur where code = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_Matricule.Text = rs("Code")
    txt_Nom.Text = rs("NOMPRN")
    w = rs.Fields("mp").Value
    k = ""
    For i = 1 To Len(w) Step 2
        k = k & Chr(Mid(w, i, 2))
    Next
    txt_mp.Text = k
    txt_Cmp.Text = k
        
     Chk_Ins_bc.Value = rs("Ins_BC")
     Chk_Maj_bc.Value = rs("Maj_BC")
     chk_Supp_BC.Value = rs("Supp_BC")
     Chk_consult_bc.Value = rs("Consult_BC")
    
     chk_ins_bv.Value = rs("Ins_BV")
     Chk_Maj_BV.Value = rs("Maj_BV")
     Chk_supp_bv.Value = rs("Supp_BV")
     Chk_consult_bv.Value = rs("Consult_BV")
    
     Chk_consult_alrt.Value = rs("Consult_Alerte")
    
     chk_ins_bcr.Value = rs("Ins_BCR")
     Chk_maj_bcr.Value = rs("Maj_BCR")
     Chk_supp_bcr.Value = rs("Supp_BCR")
     Chk_consult_bcr.Value = rs("Consult_BCR")
    
     chk_ins_pr.Value = rs("InS_PR")
     chk_maj_pr.Value = rs("Maj_PR")
     chk_supp_pr.Value = rs("Supp_PR")
     Chk_consult_pr.Value = rs("Consult_PR")
    
     chk_ins_ff.Value = rs("Ins_FF")
     chk_maj_ff.Value = rs("Maj_FF")
     chk_supp_ff.Value = rs("Supp_FF")
     chk_consult_ff.Value = rs("Consult_FF")
    
     chk_SC.Value = rs("Consult_SC")
     chk_SR.Value = rs("Consult_SR")
     chk_ST.Value = rs("Consult_ST")
     chk_EHS.Value = rs("Consult_EHS")
    
     Chk_Ins_FT.Value = rs("Ins_FT")
     Chk_Maj_FT.Value = rs("Maj_FT")
     chk_supp_FT.Value = rs("Supp_FT")
     chk_consult_FT.Value = rs("Consult_FT")
    
     chk_Consult_sup.Value = rs("Consult_SUp")
    
     chk_ins_VH.Value = rs("Ins_Vehicule")
    chk_maj_vh.Value = rs("Maj_vehicule")
     Chk_supp_vh.Value = rs("Supp_vehicule")
     Chk_consult_vh.Value = rs("Consult_vehicule")
    
     chk_ins_fr.Value = rs("Ins_Fournisseur")
     chk_maj_fr.Value = rs("Maj_Fournisseur")
     chk_supp_fr.Value = rs("Supp_Fournisseur")
     chk_consult_fr.Value = rs("Conslt_Fournisseur")
    
     chk_ins_tc.Value = rs("Ins_TC")
     Chk_maj_tc.Value = rs("Maj_TC")
     Chk_supp_tc.Value = rs("Supp_TC")
     chk_consult_tc.Value = rs("Consult_TC")
    
     chk_ins_tv.Value = rs("Ins_TV")
     chk_maj_tv.Value = rs("Maj_TV")
     chk_supp_tv.Value = rs("supp_TV")
     chk_consult_tv.Value = rs("Consult_TV")
    
     chk_ins_dest.Value = rs("Ins_Destination")
     chk_maj_dest.Value = rs("Maj_Destination")
     chk_supp_dest.Value = rs("Supp_Destination")
     chk_consult_dest.Value = rs("Consult_Destination")
    
     chk_ins_lub.Value = rs("Ins_Lub")
     chk_maj_lub.Value = rs("Maj_Lub")
     chk_supp_lub.Value = rs("Supp_Lub")
     chk_consult_lub.Value = rs("Consult_Lub")
    
     chk_ins_prod.Value = rs("Ins_Produit")
     chk_maj_prod.Value = rs("Maj_produit")
     chk_supp_prod.Value = rs("Supp_Produit")
     chk_consult_prod.Value = rs("Consult_Produit")
    
     chk_ins_per.Value = rs("Ins_Personnel")
     chk_maj_per.Value = rs("Maj_Personnel")
     chk_supp_per.Value = rs("Supp_personnel")
     chk_consult_per.Value = rs("Consult_personnel")
    
     chk_ins_user.Value = rs("Ins_Utilisateur")
     chk_maj_user.Value = rs("Maj_Utilisateur")
     chk_supp_user.Value = rs("Supp_Utilisateur")
     chk_consult_user.Value = rs("Consult_Utilisateur")
     
     chk_ins_PCH.Value = rs("Ins_PCH")
     chk_maj_PCH.Value = rs("Maj_PCH")
     chk_supp_PCH.Value = rs("Supp_PCH")
     chk_consult_PCH.Value = rs("Consult_PCH")

     chk_Actif.Value = rs("Actif")
     
     chk_Maj_Dispo.Value = rs("Maj_Disp")
     
     Chk_Consult_Compteurs.Value = rs("Consult_Compteurs")
     
     ChK_Maj_Cmpt.Value = rs("Maj_Compt")
        
Else
    MsgBox "Code introuvable", vbInformation
    txt_Matricule.SetFocus
End If
rs.Close

End Sub
