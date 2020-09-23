VERSION 5.00
Object = "{F3C40093-AF46-47A4-8D36-C0F61A2590EE}#1.0#0"; "TAB CONTROL.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C56A31&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5415
   ClientLeft      =   -30
   ClientTop       =   -450
   ClientWidth     =   7215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin ClanMng.MacButton MacButton2 
      Height          =   255
      Left            =   6120
      TabIndex        =   40
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      BTYPE           =   4
      TX              =   "Help"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421504
      FCOL            =   0
   End
   Begin ClanMng.MacButton MacButton1 
      Height          =   255
      Left            =   6840
      TabIndex        =   38
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BTYPE           =   4
      TX              =   "X"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421504
      FCOL            =   0
   End
   Begin prjXTab.XTab XTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7858
      TabCount        =   5
      TabCaption(0)   =   "General"
      TabContCtrlCnt(0)=   8
      Tab(0)ContCtrlCap(1)=   "txtClanName"
      Tab(0)ContCtrlCap(2)=   "txtCLanTag"
      Tab(0)ContCtrlCap(3)=   "cmdSaveClanName"
      Tab(0)ContCtrlCap(4)=   "cmdSaveClanTag"
      Tab(0)ContCtrlCap(5)=   "Label1"
      Tab(0)ContCtrlCap(6)=   "Label2"
      Tab(0)ContCtrlCap(7)=   "Label3"
      Tab(0)ContCtrlCap(8)=   "lblNoClanMembers"
      TabCaption(1)   =   "Members"
      TabContCtrlCnt(1)=   16
      Tab(1)ContCtrlCap(1)=   "Timer1"
      Tab(1)ContCtrlCap(2)=   "lstMembers"
      Tab(1)ContCtrlCap(3)=   "txtAddMember"
      Tab(1)ContCtrlCap(4)=   "one"
      Tab(1)ContCtrlCap(5)=   "two"
      Tab(1)ContCtrlCap(6)=   "three"
      Tab(1)ContCtrlCap(7)=   "four"
      Tab(1)ContCtrlCap(8)=   "five"
      Tab(1)ContCtrlCap(9)=   "txtAddMemberClanTag"
      Tab(1)ContCtrlCap(10)=   "cmdDelMember"
      Tab(1)ContCtrlCap(11)=   "cmdAddMember"
      Tab(1)ContCtrlCap(12)=   "cmdSaveMembers"
      Tab(1)ContCtrlCap(13)=   "Label4"
      Tab(1)ContCtrlCap(14)=   "Label5"
      Tab(1)ContCtrlCap(15)=   "Label6"
      Tab(1)ContCtrlCap(16)=   "Label7"
      TabCaption(2)   =   "Clan Wars"
      TabContCtrlCnt(2)=   9
      Tab(2)ContCtrlCap(1)=   "lstClanWars"
      Tab(2)ContCtrlCap(2)=   "txtAddClanWarName"
      Tab(2)ContCtrlCap(3)=   "txtAddClanWarTag"
      Tab(2)ContCtrlCap(4)=   "cmdDelClanWar"
      Tab(2)ContCtrlCap(5)=   "cmdAddClanWar"
      Tab(2)ContCtrlCap(6)=   "cmdSaveClanWar"
      Tab(2)ContCtrlCap(7)=   "Label8"
      Tab(2)ContCtrlCap(8)=   "Label9"
      Tab(2)ContCtrlCap(9)=   "Label10"
      TabCaption(3)   =   "Notes"
      TabContCtrlCnt(3)=   3
      Tab(3)ContCtrlCap(1)=   "txtNotes"
      Tab(3)ContCtrlCap(2)=   "cmdSaveNotes"
      Tab(3)ContCtrlCap(3)=   "Label11"
      TabCaption(4)   =   "About"
      TabContCtrlCnt(4)=   1
      Tab(4)ContCtrlCap(1)=   "Label12"
      TabStyle        =   1
      TabTheme        =   3
      ActiveTabBackStartColor=   12937777
      ActiveTabBackEndColor=   12937777
      InActiveTabBackStartColor=   12937777
      InActiveTabBackEndColor=   12937777
      ActiveTabForeColor=   16777215
      InActiveTabForeColor=   16777215
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   12632256
      BottomRightInnerBorderColor=   12937777
      DisabledTabBackColor=   12937777
      DisabledTabForeColor=   16777215
      HoverColorInverted=   12937777
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   -68760
         Top             =   3600
      End
      Begin VB.TextBox txtNotes 
         BackColor       =   &H00C56A31&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2895
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   840
         Width           =   6495
      End
      Begin ClanMng.Button cmdSaveNotes 
         Height          =   420
         Left            =   -74760
         TabIndex        =   34
         Top             =   3840
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   741
         Icon            =   "Form1.frx":0052
         Style           =   1
         Caption         =   "Save"
         IconAlign       =   1
         iNonThemeStyle  =   0
         USeCustomColors =   -1  'True
         BackColor       =   12937777
         HighlightColor  =   0
         FontColor       =   16777215
         FontHighlightColor=   16777215
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ListBox lstClanWars 
         BackColor       =   &H00C56A31&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1230
         Left            =   -74760
         TabIndex        =   30
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txtAddClanWarName 
         BackColor       =   &H00C56A31&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -74760
         TabIndex        =   29
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox txtAddClanWarTag 
         BackColor       =   &H00C56A31&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -72960
         TabIndex        =   28
         Top             =   2520
         Width           =   1815
      End
      Begin ClanMng.Button cmdDelClanWar 
         Height          =   1575
         Left            =   -71040
         TabIndex        =   27
         Top             =   780
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   2778
         Icon            =   "Form1.frx":006E
         Style           =   1
         Caption         =   "Delete"
         IconAlign       =   1
         iNonThemeStyle  =   0
         USeCustomColors =   -1  'True
         BackColor       =   12937777
         HighlightColor  =   0
         FontColor       =   16777215
         FontHighlightColor=   16777215
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ClanMng.Button cmdAddClanWar 
         Height          =   495
         Left            =   -71040
         TabIndex        =   26
         Top             =   2460
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         Icon            =   "Form1.frx":008A
         Style           =   1
         Caption         =   "Add"
         IconAlign       =   1
         iNonThemeStyle  =   0
         USeCustomColors =   -1  'True
         BackColor       =   12937777
         HighlightColor  =   0
         FontColor       =   16777215
         FontHighlightColor=   16777215
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ClanMng.Button cmdSaveClanWar 
         Height          =   420
         Left            =   -74760
         TabIndex        =   25
         Top             =   3060
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   741
         Icon            =   "Form1.frx":00A6
         Style           =   1
         Caption         =   "Save"
         IconAlign       =   1
         iNonThemeStyle  =   0
         USeCustomColors =   -1  'True
         BackColor       =   12937777
         HighlightColor  =   0
         FontColor       =   16777215
         FontHighlightColor=   16777215
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ListBox lstMembers 
         BackColor       =   &H00C56A31&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1230
         Left            =   -74760
         TabIndex        =   20
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txtAddMember 
         BackColor       =   &H00C56A31&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -74760
         TabIndex        =   19
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox one 
         BackColor       =   &H00C56A31&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -70200
         TabIndex        =   18
         Text            =   "[CL]"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox two 
         BackColor       =   &H00C56A31&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -70200
         TabIndex        =   17
         Text            =   "[CL2]"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox three 
         BackColor       =   &H00C56A31&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -70200
         TabIndex        =   16
         Text            =   "[K]"
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox four 
         BackColor       =   &H00C56A31&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -70200
         TabIndex        =   15
         Text            =   "[P]"
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox five 
         BackColor       =   &H00C56A31&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -70200
         TabIndex        =   14
         Text            =   "[I]"
         Top             =   2640
         Width           =   1815
      End
      Begin VB.ComboBox txtAddMemberClanTag 
         BackColor       =   &H00C56A31&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "Form1.frx":00C2
         Left            =   -72960
         List            =   "Form1.frx":00C4
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2520
         Width           =   1815
      End
      Begin ClanMng.Button cmdDelMember 
         Height          =   1575
         Left            =   -71040
         TabIndex        =   12
         Top             =   780
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   2778
         Icon            =   "Form1.frx":00C6
         Style           =   1
         Caption         =   "Delete"
         IconAlign       =   1
         iNonThemeStyle  =   0
         USeCustomColors =   -1  'True
         BackColor       =   12937777
         HighlightColor  =   0
         FontColor       =   16777215
         FontHighlightColor=   16777215
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ClanMng.Button cmdAddMember 
         Height          =   495
         Left            =   -71040
         TabIndex        =   11
         Top             =   2460
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         Icon            =   "Form1.frx":00E2
         Style           =   1
         Caption         =   "Add"
         IconAlign       =   1
         iNonThemeStyle  =   0
         USeCustomColors =   -1  'True
         BackColor       =   12937777
         HighlightColor  =   0
         FontColor       =   16777215
         FontHighlightColor=   16777215
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ClanMng.Button cmdSaveMembers 
         Height          =   420
         Left            =   -74760
         TabIndex        =   10
         Top             =   3060
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   741
         Icon            =   "Form1.frx":00FE
         Style           =   1
         Caption         =   "Save"
         IconAlign       =   1
         iNonThemeStyle  =   0
         USeCustomColors =   -1  'True
         BackColor       =   12937777
         HighlightColor  =   0
         FontColor       =   16777215
         FontHighlightColor=   16777215
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtClanName 
         BackColor       =   &H00C56A31&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   660
         Width           =   1695
      End
      Begin VB.TextBox txtCLanTag 
         BackColor       =   &H00C56A31&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   1140
         Width           =   1695
      End
      Begin ClanMng.Button cmdSaveClanName 
         Height          =   420
         Left            =   3120
         TabIndex        =   3
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "Form1.frx":011A
         Style           =   1
         Caption         =   "Save"
         IconAlign       =   1
         iNonThemeStyle  =   0
         USeCustomColors =   -1  'True
         BackColor       =   12937777
         HighlightColor  =   0
         FontColor       =   16777215
         FontHighlightColor=   16777215
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ClanMng.Button cmdSaveClanTag 
         Height          =   420
         Left            =   3120
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "Form1.frx":0136
         Style           =   1
         Caption         =   "Save"
         IconAlign       =   1
         iNonThemeStyle  =   0
         USeCustomColors =   -1  'True
         BackColor       =   12937777
         HighlightColor  =   0
         FontColor       =   16777215
         FontHighlightColor=   16777215
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C56A31&
         Caption         =   "Clan Manager                                        Created by Jurek Rutkowski Copyright 2004  JurekWare"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   -72960
         TabIndex        =   37
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Notes:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -74760
         TabIndex        =   36
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C56A31&
         Caption         =   "Current Clan Wars:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -74760
         TabIndex        =   33
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C56A31&
         Caption         =   "Clan Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -74760
         TabIndex        =   32
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C56A31&
         Caption         =   "Clan Tag:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -72960
         TabIndex        =   31
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C56A31&
         Caption         =   "Members:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C56A31&
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -74760
         TabIndex        =   23
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C56A31&
         Caption         =   "Rank:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -72960
         TabIndex        =   22
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C56A31&
         Caption         =   "Clan Ranks:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -70200
         TabIndex        =   21
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C56A31&
         Caption         =   "Clan Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C56A31&
         Caption         =   "Clan Tag:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C56A31&
         Caption         =   "Total Members:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1620
         Width           =   1335
      End
      Begin VB.Label lblNoClanMembers 
         BackColor       =   &H00C56A31&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   1620
         Width           =   135
      End
   End
   Begin ClanMng.Button cmdExit 
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   741
      Icon            =   "Form1.frx":0152
      Style           =   1
      Caption         =   "Exit"
      IconAlign       =   1
      iNonThemeStyle  =   0
      USeCustomColors =   -1  'True
      BackColor       =   12937777
      HighlightColor  =   0
      FontColor       =   16777215
      FontHighlightColor=   16777215
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   3
      Height          =   5415
      Left            =   0
      Top             =   0
      Width           =   7215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Clan Manager"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   0
      Picture         =   "Form1.frx":016E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7215
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About the author"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddClanWar_Click()
   'Add Clan War
   lstClanWars.AddItem "Clan Name: " & txtAddClanWarName.Text & " | Clan Tag: " & txtAddClanWarTag.Text
   txtAddClanWarName.Text = ""
   txtAddClanWarTag.Text = ""
End Sub

Private Sub cmdAddMember_Click()
   'Add Member
   lstMembers.AddItem "Name: " & txtAddMember.Text & " | Rank: " & txtAddMemberClanTag.Text
   txtAddMember.Text = ""
End Sub

Private Sub cmdDelClanWar_Click()
   If lstClanWars.ListCount > 0 Then
      If lstClanWars.ListIndex > -1 Then
         lstClanWars.RemoveItem (lstClanWars.ListIndex)
      End If
   End If
End Sub

Private Sub cmdDelMember_Click()
   If lstMembers.ListCount > 0 Then
      If lstMembers.ListIndex > -1 Then
         lstMembers.RemoveItem (lstMembers.ListIndex)
      End If
   End If
End Sub

Private Sub cmdExit_Click()
   End
End Sub

Private Sub cmdSaveClanName_Click()
SaveSetting "ClanMng", "General", "Clan Name", txtClanName
MsgBox "Clan Name Succesfully Saved!", vbOKOnly, "Saved"
End Sub

Private Sub cmdSaveClanTag_Click()
SaveSetting "ClanMng", "General", "Clan Tag", txtCLanTag
MsgBox "Clan Tag Succesfully Saved!", vbOKOnly, "Saved"
End Sub

Private Sub cmdSaveClanWar_Click()
   Open App.Path & "\ClanWars.dat" For Output As #1
   For I = 0 To lstClanWars.ListCount - 1
      Print #1, lstClanWars.List(I)
   Next
   Close #1
   
   MsgBox "Members Succesfully Saved!", vbOKOnly, "Saved"
End Sub

Private Sub cmdSaveMembers_Click()
   Open App.Path & "\Members.dat" For Output As #1
   For I = 0 To lstMembers.ListCount - 1
      Print #1, lstMembers.List(I)
   Next
   Close #1
   
   SaveSetting "ClanMng", "Members", "1", one.Text
   SaveSetting "ClanMng", "Members", "2", two.Text
   SaveSetting "ClanMng", "Members", "3", three.Text
   SaveSetting "ClanMng", "Members", "4", four.Text
   SaveSetting "ClanMng", "Members", "5", five.Text
   
   MsgBox "Members Succesfully Saved!", vbOKOnly, "Saved"
End Sub

Private Sub cmdSaveNotes_Click()
   SaveSetting "ClanMng", "Notes", "Text", txtNotes.Text
   
   MsgBox "Notes Succesfully Saved!", vbOKOnly, "Saved"
End Sub

Private Sub five_Change()
   txtAddMemberClanTag.Clear
   txtAddMemberClanTag.AddItem one.Text
   txtAddMemberClanTag.AddItem two.Text
   txtAddMemberClanTag.AddItem three.Text
   txtAddMemberClanTag.AddItem four.Text
   txtAddMemberClanTag.AddItem five.Text
End Sub

Private Sub Form_Load()
   txtClanName = GetSetting("ClanMng", "General", "Clan Name")
   txtCLanTag = GetSetting("ClanMng", "General", "Clan Tag")
   

   Open App.Path & "\Members.dat" For Input As #1
   Do Until EOF(1)
      Line Input #1, members
      lstMembers.AddItem members
   Loop
   Close #1
   
   Open App.Path & "\ClanWars.dat" For Input As #1
   Do Until EOF(1)
      Line Input #1, clanwars
      lstClanWars.AddItem clanwars
   Loop
   Close #1
   
   
   one = GetSetting("ClanMng", "Members", "1")
   two = GetSetting("ClanMng", "Members", "2")
   three = GetSetting("ClanMng", "Members", "3")
   four = GetSetting("ClanMng", "Members", "4")
   five = GetSetting("ClanMng", "Members", "5")

   txtAddMemberClanTag.Clear
   txtAddMemberClanTag.AddItem one.Text
   txtAddMemberClanTag.AddItem two.Text
   txtAddMemberClanTag.AddItem three.Text
   txtAddMemberClanTag.AddItem four.Text
   txtAddMemberClanTag.AddItem five.Text
   
   txtNotes.Text = GetSetting("ClanMng", "Notes", "Text")
End Sub

Private Sub four_Change()
   txtAddMemberClanTag.Clear
   txtAddMemberClanTag.AddItem one.Text
   txtAddMemberClanTag.AddItem two.Text
   txtAddMemberClanTag.AddItem three.Text
   txtAddMemberClanTag.AddItem four.Text
   txtAddMemberClanTag.AddItem five.Text
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ReleaseCapture
   SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ReleaseCapture
   SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub MacButton1_Click()
   End
End Sub

Private Sub MacButton2_Click()
   PopupMenu mnuHelp
End Sub

Private Sub mnuHelpAbout_Click()
   MsgBox "Created by Jurek Rutkowski, JurekWare." & vbCrLf & "Do don't copy any meterials of this program or You will be punished by the law."
End Sub

Private Sub one_Change()
   txtAddMemberClanTag.Clear
   txtAddMemberClanTag.AddItem one.Text
   txtAddMemberClanTag.AddItem two.Text
   txtAddMemberClanTag.AddItem three.Text
   txtAddMemberClanTag.AddItem four.Text
   txtAddMemberClanTag.AddItem five.Text
End Sub

Private Sub three_Change()
   txtAddMemberClanTag.Clear
   txtAddMemberClanTag.AddItem one.Text
   txtAddMemberClanTag.AddItem two.Text
   txtAddMemberClanTag.AddItem three.Text
   txtAddMemberClanTag.AddItem four.Text
   txtAddMemberClanTag.AddItem five.Text
End Sub

Private Sub Timer1_Timer()
   lblNoClanMembers.Caption = lstMembers.ListCount
End Sub

Private Sub two_Change()
   txtAddMemberClanTag.Clear
   txtAddMemberClanTag.AddItem one.Text
   txtAddMemberClanTag.AddItem two.Text
   txtAddMemberClanTag.AddItem three.Text
   txtAddMemberClanTag.AddItem four.Text
   txtAddMemberClanTag.AddItem five.Text
End Sub
