VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0"; "MSOUTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB.Print!"
   ClientHeight    =   5520
   ClientLeft      =   2760
   ClientTop       =   1935
   ClientWidth     =   7755
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5520
   ScaleWidth      =   7755
   Begin VB.PictureBox picPaper 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3720
      ScaleHeight     =   285
      ScaleWidth      =   480
      TabIndex        =   140
      TabStop         =   0   'False
      Top             =   4875
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4260
      ScaleHeight     =   5.027
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   8.467
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   4875
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.ListBox lstNames 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   5310
      Sorted          =   -1  'True
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   4815
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.ListBox lstNames 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      IntegralHeight  =   0   'False
      Left            =   5310
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   4995
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "&Restore settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1530
      TabIndex        =   32
      Top             =   4860
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save settings"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   30
      TabIndex        =   31
      Top             =   4860
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   12
      Top             =   4860
      Width           =   1110
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   4755
      Left            =   45
      TabIndex        =   13
      Top             =   30
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   8387
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "VBPrin&t "
      TabPicture(0)   =   "Main.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblUser"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdPrint"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSelectAll"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdClear"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdView"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdPrintSetup"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdAbout"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdHelp"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame(7)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Pre&ferences"
      TabPicture(1)   =   "Main.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame(3)"
      Tab(1).Control(1)=   "TabOptions"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Files and Procedures"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3195
         Left            =   105
         TabIndex        =   100
         Top             =   1020
         Width           =   3330
         Begin MSOutl.Outline Outline 
            Height          =   2940
            Left            =   75
            TabIndex        =   3
            Top             =   180
            Width           =   3180
            _Version        =   65536
            _ExtentX        =   5609
            _ExtentY        =   5186
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            PathSeparator   =   "-"
            PicturePlus     =   "Main.frx":0342
            PictureMinus    =   "Main.frx":043C
            PictureLeaf     =   "Main.frx":0536
            PictureOpen     =   "Main.frx":0E10
            PictureClosed   =   "Main.frx":16EA
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "&Source file"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   7
         Left            =   105
         TabIndex        =   85
         Top             =   345
         Width           =   7425
         Begin VB.CommandButton cmdPrevFile 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7005
            TabIndex        =   2
            Top             =   240
            Width           =   255
         End
         Begin VB.CommandButton cmdPickFile 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6750
            TabIndex        =   1
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtProject 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   0
            Top             =   210
            Width           =   7170
         End
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "&Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4755
         TabIndex        =   8
         Top             =   4320
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "&About..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4755
         TabIndex        =   7
         Top             =   3900
         Width           =   1110
      End
      Begin VB.Frame Frame 
         Caption         =   "Printer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Index           =   2
         Left            =   3525
         TabIndex        =   82
         Top             =   2490
         Width           =   4005
         Begin VB.CheckBox chkPreview 
            Caption         =   "Previe&w"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1890
            TabIndex        =   9
            Top             =   990
            Width           =   930
         End
         Begin VB.Label lblViewSize 
            Alignment       =   1  'Right Justify
            Caption         =   "(Zoom 100%)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2850
            TabIndex        =   97
            Top             =   990
            Width           =   1035
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   9
            X1              =   1755
            X2              =   1755
            Y1              =   1260
            Y2              =   900
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   8
            X1              =   1770
            X2              =   1770
            Y1              =   900
            Y2              =   1275
         End
         Begin VB.Label lblPrinted 
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1215
            TabIndex        =   87
            Top             =   990
            UseMnemonic     =   0   'False
            Width           =   450
         End
         Begin VB.Label Label 
            Caption         =   "Pages printed:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   86
            Top             =   990
            UseMnemonic     =   0   'False
            Width           =   1080
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   15
            X2              =   4000
            Y1              =   885
            Y2              =   885
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   0
            X1              =   30
            X2              =   3980
            Y1              =   870
            Y2              =   870
         End
         Begin VB.Label lblPrinter 
            Caption         =   "None selected -  Select 'Print Setup'"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   120
            TabIndex        =   83
            Top             =   225
            UseMnemonic     =   0   'False
            Width           =   3780
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Files"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Index           =   0
         Left            =   3525
         TabIndex        =   77
         Top             =   1020
         Width           =   4005
         Begin VB.Label lblSelProcs 
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2745
            TabIndex        =   104
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   450
         End
         Begin VB.Label lblProcedures 
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1050
            TabIndex        =   103
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   450
         End
         Begin VB.Label Label 
            Caption         =   "Selected:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   15
            Left            =   1995
            TabIndex        =   102
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   750
         End
         Begin VB.Label Label 
            Caption         =   "Procedures:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   120
            TabIndex        =   101
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   915
         End
         Begin VB.Label lblName 
            Caption         =   "(No item selected)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   89
            Top             =   870
            Width           =   3780
         End
         Begin VB.Label lblType 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   88
            Top             =   1125
            UseMnemonic     =   0   'False
            Width           =   3780
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   3
            X1              =   30
            X2              =   3980
            Y1              =   765
            Y2              =   765
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   2
            X1              =   15
            X2              =   4000
            Y1              =   780
            Y2              =   780
         End
         Begin VB.Label lblFiles 
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1050
            TabIndex        =   81
            Top             =   225
            UseMnemonic     =   0   'False
            Width           =   450
         End
         Begin VB.Label Label 
            Caption         =   "Selected:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   1995
            TabIndex        =   80
            Top             =   225
            UseMnemonic     =   0   'False
            Width           =   750
         End
         Begin VB.Label Label 
            Caption         =   "Files:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   79
            Top             =   225
            UseMnemonic     =   0   'False
            Width           =   915
         End
         Begin VB.Label lblSelFiles 
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2745
            TabIndex        =   78
            Top             =   225
            UseMnemonic     =   0   'False
            Width           =   450
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Page sample"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4245
         Index           =   3
         Left            =   -74895
         TabIndex        =   70
         Top             =   360
         Width           =   3735
         Begin VB.Timer TmrPaint 
            Enabled         =   0   'False
            Interval        =   2000
            Left            =   3195
            Top             =   3135
         End
         Begin VB.PictureBox picPage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2865
            Left            =   165
            ScaleHeight     =   2835
            ScaleWidth      =   2070
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   255
            Width           =   2100
         End
         Begin ComctlLib.Slider sldRight 
            Height          =   225
            Left            =   1440
            TabIndex        =   29
            Top             =   3120
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   397
            _Version        =   327682
            Max             =   70
            SelStart        =   70
            TickFrequency   =   10
            Value           =   70
         End
         Begin ComctlLib.Slider sldBottom 
            Height          =   1185
            Left            =   2265
            TabIndex        =   28
            Top             =   2040
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   2090
            _Version        =   327682
            Orientation     =   1
            Max             =   99
            SelStart        =   99
            TickFrequency   =   10
            Value           =   99
         End
         Begin ComctlLib.Slider sldTop 
            Height          =   1185
            Left            =   2265
            TabIndex        =   27
            Top             =   165
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   2090
            _Version        =   327682
            Orientation     =   1
            Max             =   99
            TickFrequency   =   10
         End
         Begin ComctlLib.Slider sldLeft 
            Height          =   225
            Left            =   60
            TabIndex        =   30
            Top             =   3120
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   397
            _Version        =   327682
            Max             =   70
            TickFrequency   =   10
         End
         Begin VB.Label Label 
            Caption         =   "Orientation:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   90
            TabIndex        =   93
            Top             =   3720
            UseMnemonic     =   0   'False
            Width           =   900
         End
         Begin VB.Label Label 
            Caption         =   "Page size:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   90
            TabIndex        =   92
            Top             =   3960
            UseMnemonic     =   0   'False
            Width           =   900
         End
         Begin VB.Label lblOrient 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1005
            TabIndex        =   91
            Top             =   3720
            UseMnemonic     =   0   'False
            Width           =   2655
         End
         Begin VB.Label lblSize 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1005
            TabIndex        =   90
            Top             =   3960
            UseMnemonic     =   0   'False
            Width           =   2655
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   5
            X1              =   30
            X2              =   3720
            Y1              =   3615
            Y2              =   3615
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   4
            X1              =   15
            X2              =   3720
            Y1              =   3630
            Y2              =   3630
         End
         Begin VB.Label lblMM 
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2565
            TabIndex        =   76
            Top             =   3390
            Width           =   255
         End
         Begin VB.Label lblLeft 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   150
            TabIndex        =   75
            Top             =   3390
            Width           =   330
         End
         Begin VB.Label lblRight 
            Alignment       =   1  'Right Justify
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   1965
            TabIndex        =   74
            Top             =   3390
            Width           =   330
         End
         Begin VB.Label lblBottom 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   2565
            TabIndex        =   73
            Top             =   2970
            Width           =   330
         End
         Begin VB.Label lblTop 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   2565
            TabIndex        =   72
            Top             =   240
            Width           =   330
         End
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "Print &Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6450
         TabIndex        =   10
         Top             =   3900
         Width           =   1110
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3540
         TabIndex        =   6
         Top             =   3900
         Width           =   1110
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear All"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2310
         TabIndex        =   5
         Top             =   4320
         Width           =   1110
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Select &All"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         TabIndex        =   4
         Top             =   4320
         Width           =   1110
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6450
         TabIndex        =   11
         Top             =   4320
         Width           =   1110
      End
      Begin TabDlg.SSTab TabOptions 
         Height          =   4155
         Left            =   -71040
         TabIndex        =   14
         Top             =   450
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   7329
         _Version        =   393216
         TabOrientation  =   3
         Style           =   1
         Tabs            =   6
         TabsPerRow      =   6
         TabHeight       =   520
         WordWrap        =   0   'False
         TabCaption(0)   =   "Options"
         TabPicture(0)   =   "Main.frx":1FC4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "chkSortIndex"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "chkIndex"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "chkControlPage"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "chkProcNames"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "chkProject"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "chkSeparator"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "chkCode"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "chkControlNames"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "chkSortControls"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "chkIcon"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "chkProcPage"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "chkFormIcons"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).ControlCount=   12
         TabCaption(1)   =   "Page"
         TabPicture(1)   =   "Main.frx":1FE0
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "optPagePos(1)"
         Tab(1).Control(1)=   "optPagePos(0)"
         Tab(1).Control(2)=   "chkFooter"
         Tab(1).Control(3)=   "chkResetPage"
         Tab(1).Control(4)=   "chkTime"
         Tab(1).Control(5)=   "chkDate"
         Tab(1).Control(6)=   "chkHeader"
         Tab(1).Control(7)=   "chkPageNumbers"
         Tab(1).Control(8)=   "txtOwner(1)"
         Tab(1).Control(9)=   "txtOwner(0)"
         Tab(1).Control(10)=   "Line(17)"
         Tab(1).Control(11)=   "Line(16)"
         Tab(1).Control(12)=   "Label(8)"
         Tab(1).ControlCount=   13
         TabCaption(2)   =   "User"
         TabPicture(2)   =   "Main.frx":1FFC
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "chkMinimize"
         Tab(2).Control(1)=   "cmdEditor"
         Tab(2).Control(2)=   "cmdPlay"
         Tab(2).Control(3)=   "cboSounds"
         Tab(2).Control(4)=   "cboZoom"
         Tab(2).Control(5)=   "cboExtention"
         Tab(2).Control(6)=   "chkPlayWaves"
         Tab(2).Control(7)=   "Line(21)"
         Tab(2).Control(8)=   "Line(20)"
         Tab(2).Control(9)=   "Line(19)"
         Tab(2).Control(10)=   "Line(18)"
         Tab(2).Control(11)=   "lblSoundFile"
         Tab(2).Control(12)=   "lblZoomDialog"
         Tab(2).Control(13)=   "Line(15)"
         Tab(2).Control(14)=   "Line(14)"
         Tab(2).Control(15)=   "lblExtention"
         Tab(2).Control(16)=   "Line(11)"
         Tab(2).Control(17)=   "Line(10)"
         Tab(2).ControlCount=   18
         TabCaption(3)   =   "Printer"
         TabPicture(3)   =   "Main.frx":2018
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "cmdTextTest"
         Tab(3).Control(1)=   "chkFormFeed"
         Tab(3).Control(2)=   "picContainer(1)"
         Tab(3).Control(3)=   "cboPort"
         Tab(3).Control(4)=   "cboWidth"
         Tab(3).Control(5)=   "cboHeight"
         Tab(3).Control(6)=   "lblPort(5)"
         Tab(3).Control(7)=   "lblPort(6)"
         Tab(3).Control(8)=   "lblRight(1)"
         Tab(3).Control(9)=   "lblBottom(1)"
         Tab(3).Control(10)=   "lblTop(1)"
         Tab(3).Control(11)=   "lblLeft(1)"
         Tab(3).Control(12)=   "lblPort(4)"
         Tab(3).Control(13)=   "lblPort(3)"
         Tab(3).Control(14)=   "lblPort(2)"
         Tab(3).Control(15)=   "lblPort(1)"
         Tab(3).Control(16)=   "lblPort(0)"
         Tab(3).Control(17)=   "Label(16)"
         Tab(3).Control(18)=   "lblOutput(0)"
         Tab(3).Control(19)=   "lblOutput(1)"
         Tab(3).Control(20)=   "lblOutput(2)"
         Tab(3).Control(21)=   "lblOutput(3)"
         Tab(3).Control(22)=   "lblOutput(4)"
         Tab(3).Control(23)=   "Line(13)"
         Tab(3).Control(24)=   "Line(12)"
         Tab(3).ControlCount=   25
         TabCaption(4)   =   "Fonts"
         TabPicture(4)   =   "Main.frx":2034
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "cmdFont(0)"
         Tab(4).Control(1)=   "cmdFont(1)"
         Tab(4).Control(2)=   "cmdFont(2)"
         Tab(4).Control(3)=   "cmdFont(3)"
         Tab(4).Control(4)=   "cmdFont(4)"
         Tab(4).Control(5)=   "cmdFont(5)"
         Tab(4).Control(6)=   "cmdFont(6)"
         Tab(4).Control(7)=   "lblFont(0)"
         Tab(4).Control(8)=   "lblFont(1)"
         Tab(4).Control(9)=   "lblFont(2)"
         Tab(4).Control(10)=   "lblFont(3)"
         Tab(4).Control(11)=   "lblFont(4)"
         Tab(4).Control(12)=   "lblFont(5)"
         Tab(4).Control(13)=   "lblFont(6)"
         Tab(4).ControlCount=   14
         TabCaption(5)   =   "Rego"
         TabPicture(5)   =   "Main.frx":2050
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "cmdRegPrint"
         Tab(5).Control(1)=   "cmdHowTo"
         Tab(5).Control(2)=   "cmdRegister"
         Tab(5).Control(3)=   "txtRegName"
         Tab(5).Control(4)=   "txtRegKey"
         Tab(5).Control(5)=   "Label(14)"
         Tab(5).Control(6)=   "Label(13)"
         Tab(5).Control(7)=   "Label(12)"
         Tab(5).Control(8)=   "Label(11)"
         Tab(5).Control(9)=   "lblVersion"
         Tab(5).Control(10)=   "ImageLogo"
         Tab(5).Control(11)=   "Line(7)"
         Tab(5).Control(12)=   "Line(6)"
         Tab(5).Control(13)=   "Label(10)"
         Tab(5).Control(14)=   "Label(9)"
         Tab(5).ControlCount=   15
         Begin VB.CheckBox chkMinimize 
            Caption         =   "Minimize main form while printing"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74910
            TabIndex        =   144
            Top             =   2430
            Value           =   1  'Checked
            Width           =   3060
         End
         Begin VB.CommandButton cmdEditor 
            Caption         =   "&Load file into text editor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -74880
            TabIndex        =   143
            Top             =   3720
            Width           =   1875
         End
         Begin VB.CommandButton cmdRegPrint 
            Caption         =   "Print &form"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -73845
            TabIndex        =   68
            Top             =   3720
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.CommandButton cmdPlay 
            Caption         =   "&Play"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -72600
            TabIndex        =   46
            Top             =   1155
            Width           =   750
         End
         Begin VB.ComboBox cboSounds 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -74910
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   1155
            Width           =   2220
         End
         Begin VB.ComboBox cboZoom 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Main.frx":206C
            Left            =   -72600
            List            =   "Main.frx":2085
            TabIndex        =   47
            Text            =   "100%"
            Top             =   1920
            Width           =   750
         End
         Begin VB.CommandButton cmdTextTest 
            Caption         =   "Test text"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -74880
            TabIndex        =   57
            Top             =   3720
            Width           =   1110
         End
         Begin VB.OptionButton optPagePos 
            Caption         =   "Place number in footer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   -74325
            TabIndex        =   38
            Top             =   1140
            Width           =   2490
         End
         Begin VB.OptionButton optPagePos 
            Caption         =   "Place number in header"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   -74325
            TabIndex        =   37
            Top             =   885
            Value           =   -1  'True
            Width           =   2490
         End
         Begin VB.CheckBox chkFooter 
            Caption         =   "&Footers"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -73365
            TabIndex        =   34
            Top             =   120
            Value           =   1  'Checked
            Width           =   1530
         End
         Begin VB.CheckBox chkResetPage 
            Caption         =   "&Reset number for each file"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -74325
            TabIndex        =   36
            Top             =   630
            Width           =   2490
         End
         Begin VB.CheckBox chkTime 
            Caption         =   "&Time Stamp"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74625
            TabIndex        =   40
            Top             =   1650
            Value           =   1  'Checked
            Width           =   2790
         End
         Begin VB.CheckBox chkDate 
            Caption         =   "&Date Stamp"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74625
            TabIndex        =   39
            Top             =   1395
            Value           =   1  'Checked
            Width           =   2790
         End
         Begin VB.CheckBox chkHeader 
            Caption         =   "&Headers"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74895
            TabIndex        =   33
            Top             =   120
            Value           =   1  'Checked
            Width           =   1410
         End
         Begin VB.CheckBox chkPageNumbers 
            Caption         =   "Page &Numbers"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74625
            TabIndex        =   35
            Top             =   375
            Value           =   1  'Checked
            Width           =   2790
         End
         Begin VB.ComboBox cboExtention 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Main.frx":20AF
            Left            =   -74910
            List            =   "Main.frx":20B1
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   360
            Width           =   3060
         End
         Begin VB.CheckBox chkPlayWaves 
            Caption         =   "Play sound events"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74910
            TabIndex        =   44
            Top             =   870
            Value           =   1  'Checked
            Width           =   3060
         End
         Begin VB.CheckBox chkFormFeed 
            Alignment       =   1  'Right Justify
            Caption         =   "Force &formfeed"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -74550
            TabIndex        =   56
            Tag             =   "  "
            Top             =   2535
            Value           =   1  'Checked
            Width           =   1725
         End
         Begin VB.PictureBox picContainer 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Index           =   1
            Left            =   -74790
            ScaleHeight     =   1080
            ScaleWidth      =   2925
            TabIndex        =   120
            Top             =   345
            Width           =   2925
            Begin VB.CommandButton cmdPickRtf 
               Caption         =   "..."
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2580
               TabIndex        =   52
               Top             =   510
               Width           =   255
            End
            Begin VB.TextBox txtRTFfile 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   270
               TabIndex        =   51
               Text            =   "Listing.rtf"
               Top             =   480
               Width           =   2595
            End
            Begin VB.OptionButton optOutput 
               Caption         =   "Direct to &port (text only)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   0
               TabIndex        =   50
               Top             =   855
               Width           =   2550
            End
            Begin VB.OptionButton optOutput 
               Caption         =   "Rich-Text &file (no images)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   0
               TabIndex        =   49
               Top             =   255
               Width           =   2535
            End
            Begin VB.OptionButton optOutput 
               Caption         =   "Use &windows (graphics)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   0
               TabIndex        =   48
               Top             =   0
               Value           =   -1  'True
               Width           =   2565
            End
         End
         Begin VB.ComboBox cboPort 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Main.frx":20B3
            Left            =   -73560
            List            =   "Main.frx":20C0
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Tag             =   "  "
            Top             =   1455
            Width           =   735
         End
         Begin VB.ComboBox cboWidth 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Main.frx":20D6
            Left            =   -73560
            List            =   "Main.frx":20E0
            TabIndex        =   55
            Tag             =   "  "
            Text            =   "80"
            Top             =   2175
            Width           =   735
         End
         Begin VB.ComboBox cboHeight 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Main.frx":20ED
            Left            =   -73560
            List            =   "Main.frx":20F7
            TabIndex        =   54
            Tag             =   "  "
            Text            =   "66"
            Top             =   1815
            Width           =   735
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Change"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   -72600
            TabIndex        =   58
            Top             =   135
            Width           =   750
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Change"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   -72600
            TabIndex        =   59
            Top             =   525
            Width           =   750
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Change"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   -72600
            TabIndex        =   60
            Top             =   915
            Width           =   750
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Change"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   -72600
            TabIndex        =   61
            Top             =   1305
            Width           =   750
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Change"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   -72600
            TabIndex        =   62
            Top             =   1710
            Width           =   750
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Change"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   -72600
            TabIndex        =   63
            Top             =   2115
            Width           =   750
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Change"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   -72600
            TabIndex        =   64
            Top             =   2520
            Width           =   750
         End
         Begin VB.CommandButton cmdHowTo 
            Caption         =   "&How to..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -74880
            TabIndex        =   67
            Top             =   3720
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.CommandButton cmdRegister 
            Caption         =   "&Register"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -72810
            TabIndex        =   69
            Top             =   3720
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.TextBox txtRegName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -74880
            MaxLength       =   64
            TabIndex        =   65
            Text            =   "Freeware"
            Top             =   2790
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.TextBox txtRegKey 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -74880
            Locked          =   -1  'True
            TabIndex        =   66
            Top             =   3345
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.CheckBox chkFormIcons 
            Caption         =   "Print all for&m icons on seperate page(s)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   16
            Top             =   375
            Width           =   3060
         End
         Begin VB.TextBox txtOwner 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   -74910
            TabIndex        =   42
            Text            =   "Written by Inner Control Business Management"
            Top             =   2580
            Width           =   3060
         End
         Begin VB.TextBox txtOwner 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   -74910
            TabIndex        =   41
            Text            =   "This listing is created with VBPrint"
            Top             =   2235
            Width           =   3060
         End
         Begin VB.CheckBox chkProcPage 
            Caption         =   "One procedure per pa&ge"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   375
            TabIndex        =   23
            Top             =   2160
            Width           =   2790
         End
         Begin VB.CheckBox chkIcon 
            Caption         =   "Form Ic&on (picture)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   17
            Top             =   630
            Value           =   1  'Checked
            Width           =   3060
         End
         Begin VB.CheckBox chkSortControls 
            Caption         =   "Sort controls by name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   375
            TabIndex        =   19
            Top             =   1140
            Value           =   1  'Checked
            Width           =   2790
         End
         Begin VB.CheckBox chkControlNames 
            Caption         =   "&Form Control Names"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   18
            Top             =   885
            Value           =   1  'Checked
            Width           =   3060
         End
         Begin VB.CheckBox chkCode 
            Caption         =   "&Code"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   21
            Top             =   1650
            Value           =   1  'Checked
            Width           =   3060
         End
         Begin VB.CheckBox chkSeparator 
            Caption         =   "Separator &Line"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   375
            TabIndex        =   24
            Top             =   2415
            Value           =   1  'Checked
            Width           =   2790
         End
         Begin VB.CheckBox chkProject 
            Caption         =   "Project &Information"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   15
            Top             =   120
            Value           =   1  'Checked
            Width           =   3060
         End
         Begin VB.CheckBox chkProcNames 
            Caption         =   "Procedures n&ames only"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   375
            TabIndex        =   22
            Top             =   1905
            Width           =   2790
         End
         Begin VB.CheckBox chkControlPage 
            Caption         =   "Print names on separate pa&ge(s)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   375
            TabIndex        =   20
            Top             =   1395
            Width           =   2790
         End
         Begin VB.CheckBox chkIndex 
            Caption         =   "&Index page"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   25
            Top             =   2670
            Width           =   3060
         End
         Begin VB.CheckBox chkSortIndex 
            Caption         =   "Sort index by name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   375
            TabIndex        =   26
            Top             =   2925
            Width           =   2790
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   21
            X1              =   -74970
            X2              =   -71745
            Y1              =   2685
            Y2              =   2685
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   20
            X1              =   -75000
            X2              =   -71760
            Y1              =   2700
            Y2              =   2700
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   19
            X1              =   -75000
            X2              =   -71760
            Y1              =   2340
            Y2              =   2340
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   18
            X1              =   -74970
            X2              =   -71745
            Y1              =   2325
            Y2              =   2325
         End
         Begin VB.Label lblSoundFile 
            Caption         =   "Wave file: n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -74865
            TabIndex        =   142
            Top             =   1545
            Width           =   2955
         End
         Begin VB.Label lblZoomDialog 
            Caption         =   "Default preview page zoom:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -74865
            TabIndex        =   127
            Top             =   1965
            Width           =   2235
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   15
            X1              =   -74985
            X2              =   -71760
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   14
            X1              =   -75015
            X2              =   -71775
            Y1              =   1815
            Y2              =   1815
         End
         Begin VB.Label lblPort 
            Caption         =   "lines"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   -72675
            TabIndex        =   139
            Top             =   3180
            Width           =   405
         End
         Begin VB.Label lblPort 
            Caption         =   "characters"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   -72675
            TabIndex        =   138
            Tag             =   "  "
            Top             =   3435
            Width           =   795
         End
         Begin VB.Label lblRight 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   -73080
            TabIndex        =   137
            Tag             =   "  "
            Top             =   3435
            Width           =   330
         End
         Begin VB.Label lblBottom 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   -73080
            TabIndex        =   136
            Tag             =   "  "
            Top             =   3180
            Width           =   330
         End
         Begin VB.Label lblTop 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   -74340
            TabIndex        =   135
            Tag             =   "  "
            Top             =   3180
            Width           =   330
         End
         Begin VB.Label lblLeft 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   -74340
            TabIndex        =   134
            Tag             =   "  "
            Top             =   3435
            Width           =   330
         End
         Begin VB.Label lblPort 
            Caption         =   "Right:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   -73710
            TabIndex        =   133
            Tag             =   "  "
            Top             =   3435
            UseMnemonic     =   0   'False
            Width           =   600
         End
         Begin VB.Label lblPort 
            Caption         =   "Left:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   -74745
            TabIndex        =   132
            Tag             =   "  "
            Top             =   3435
            UseMnemonic     =   0   'False
            Width           =   375
         End
         Begin VB.Label lblPort 
            Caption         =   "Bottom:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   -73710
            TabIndex        =   131
            Tag             =   "  "
            Top             =   3180
            UseMnemonic     =   0   'False
            Width           =   600
         End
         Begin VB.Label lblPort 
            Caption         =   "Top:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   -74775
            TabIndex        =   130
            Tag             =   "  "
            Top             =   3180
            UseMnemonic     =   0   'False
            Width           =   375
         End
         Begin VB.Label lblPort 
            Caption         =   "Text page margins:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   -74880
            TabIndex        =   129
            Tag             =   "  "
            Top             =   2925
            UseMnemonic     =   0   'False
            Width           =   2985
         End
         Begin VB.Label lblExtention 
            Caption         =   "Default extention for 'Select file' dialog:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -74865
            TabIndex        =   128
            Top             =   105
            Width           =   2985
         End
         Begin VB.Label Label 
            Caption         =   "Printer device:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   16
            Left            =   -74865
            TabIndex        =   126
            Top             =   105
            Width           =   2985
         End
         Begin VB.Label lblOutput 
            Caption         =   "Printer port:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   -74520
            TabIndex        =   125
            Tag             =   "  "
            Top             =   1515
            Width           =   960
         End
         Begin VB.Label lblOutput 
            Caption         =   "Page width:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   -74520
            TabIndex        =   124
            Tag             =   "  "
            Top             =   2220
            Width           =   960
         End
         Begin VB.Label lblOutput 
            Caption         =   "Page height:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   -74520
            TabIndex        =   123
            Tag             =   "  "
            Top             =   1860
            Width           =   960
         End
         Begin VB.Label lblOutput 
            Caption         =   "lines"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   -72675
            TabIndex        =   122
            Tag             =   "  "
            Top             =   1860
            Width           =   405
         End
         Begin VB.Label lblOutput 
            Caption         =   "characters"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   -72675
            TabIndex        =   121
            Tag             =   "  "
            Top             =   2220
            Width           =   795
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   13
            Tag             =   "  "
            X1              =   -74985
            X2              =   -71745
            Y1              =   2820
            Y2              =   2820
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   12
            Tag             =   "  "
            X1              =   -74985
            X2              =   -71760
            Y1              =   2805
            Y2              =   2805
         End
         Begin VB.Label lblFont 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Procedures"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   -74895
            TabIndex        =   119
            Top             =   135
            UseMnemonic     =   0   'False
            Width           =   2220
         End
         Begin VB.Label lblFont 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Code"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   -74895
            TabIndex        =   118
            Top             =   525
            Width           =   2220
         End
         Begin VB.Label lblFont 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Comments"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   -74895
            TabIndex        =   117
            Top             =   915
            Width           =   2220
         End
         Begin VB.Label lblFont 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Headers"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   3
            Left            =   -74895
            TabIndex        =   116
            Top             =   1305
            UseMnemonic     =   0   'False
            Width           =   2220
         End
         Begin VB.Label lblFont 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Footers"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   4
            Left            =   -74895
            TabIndex        =   115
            Top             =   1710
            UseMnemonic     =   0   'False
            Width           =   2220
         End
         Begin VB.Label lblFont 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Directives"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   5
            Left            =   -74895
            TabIndex        =   114
            Top             =   2115
            Width           =   2220
         End
         Begin VB.Label lblFont 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Titles"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   6
            Left            =   -74895
            TabIndex        =   113
            Top             =   2520
            UseMnemonic     =   0   'False
            Width           =   2220
         End
         Begin VB.Label Label 
            Caption         =   "All rights reserved."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   210
            Index           =   14
            Left            =   -74835
            TabIndex        =   112
            Top             =   1560
            Width           =   2940
         End
         Begin VB.Label Label 
            Caption         =   "Inner Control Business Management."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   420
            Index           =   13
            Left            =   -73515
            TabIndex        =   111
            Top             =   1140
            Width           =   1665
         End
         Begin VB.Label Label 
            Caption         =   "VB.Print! is a freeware product. Please feel free to distribute this software with others."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   12
            Left            =   -74835
            TabIndex        =   110
            Top             =   1815
            Width           =   2880
         End
         Begin VB.Label Label 
            Caption         =   "Copyright© 1997-2000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   405
            Index           =   11
            Left            =   -74835
            TabIndex        =   109
            Top             =   1140
            Width           =   1260
         End
         Begin VB.Label lblVersion 
            Alignment       =   2  'Center
            Caption         =   "Version: 1.0.0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -74835
            TabIndex        =   108
            Top             =   900
            Width           =   2940
         End
         Begin VB.Image ImageLogo 
            Height          =   735
            Left            =   -74865
            Picture         =   "Main.frx":2103
            Stretch         =   -1  'True
            Top             =   120
            Width           =   3000
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   7
            Visible         =   0   'False
            X1              =   -75000
            X2              =   -71760
            Y1              =   2490
            Y2              =   2490
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   6
            Visible         =   0   'False
            X1              =   -74985
            X2              =   -71760
            Y1              =   2475
            Y2              =   2475
         End
         Begin VB.Label Label 
            Caption         =   "Registration unlock key code:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   10
            Left            =   -74865
            TabIndex        =   107
            Top             =   3120
            Visible         =   0   'False
            Width           =   3000
         End
         Begin VB.Label Label 
            Caption         =   "User name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   9
            Left            =   -74865
            TabIndex        =   106
            Top             =   2550
            Visible         =   0   'False
            Width           =   3000
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   17
            X1              =   -74985
            X2              =   -71745
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   16
            X1              =   -74985
            X2              =   -71760
            Y1              =   1905
            Y2              =   1905
         End
         Begin VB.Label Label 
            Caption         =   "Footer text (eg. Code owner details)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   -74865
            TabIndex        =   105
            Top             =   1980
            Width           =   2985
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   11
            X1              =   -75000
            X2              =   -71760
            Y1              =   780
            Y2              =   780
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   10
            X1              =   -74985
            X2              =   -71760
            Y1              =   765
            Y2              =   765
         End
      End
      Begin VB.Label lblUser 
         Alignment       =   1  'Right Justify
         Caption         =   "Freeware"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1995
         TabIndex        =   94
         Top             =   45
         Width           =   5580
      End
   End
   Begin VB.TextBox txtWrap 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4815
      MultiLine       =   -1  'True
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   4875
      Visible         =   0   'False
      Width           =   450
   End
   Begin RichTextLib.RichTextBox RTBox 
      Height          =   285
      Left            =   3165
      TabIndex        =   141
      TabStop         =   0   'False
      Top             =   4875
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   503
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      DisableNoScroll =   -1  'True
      TextRTF         =   $"Main.frx":941D
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   84
      Top             =   5220
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   529
      Style           =   1
      SimpleText      =   "Ready"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6105
      Top             =   4785
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "Files"
      Visible         =   0   'False
      Begin VB.Menu mnuRecentFile 
         Caption         =   "(Empty)"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuRecentBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopCancel 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Comments here belong to the declarations section - these are just here for testing.

Dim sLoadedFile As String
Dim bBinding As Boolean
Dim bZoomRefresh As Boolean

' These are some comments it the top - belonging to the next
' procedure. A empty line above marks the seperation.
' So long no statements are below these comments, these will belong to the procedure.

Private Sub Form_Load()
   On Error Resume Next

   lblVersion = "Version: " & Format(App.Major, "###0") & "." & Format(App.Minor, "###0") & "." & Format(App.Revision, "###0")

   bBinding = True

   ' Set the controls...
   cboPort.ListIndex = 0

   cboExtention.AddItem "VB files (*.vbp;*.frm;*.bas;*.cls;*.ctl;*.pag;*.dob)"
   cboExtention.AddItem "Project files (*.vbp)"
   cboExtention.AddItem "Form files (*.frm)"
   cboExtention.AddItem "Module files (*.bas)"
   cboExtention.AddItem "Class files (*.cls)"
   cboExtention.AddItem "User Control files (*.ctl)"
   cboExtention.AddItem "Property Page files (*.pag)"
   cboExtention.AddItem "User Document files (*.dob)"
   cboExtention.AddItem "All files (*.*)"
   cboExtention.ListIndex = 0

   cboSounds.AddItem "Files accessed"
   cboSounds.AddItem "Analysing"
   cboSounds.AddItem "Error"
   cboSounds.AddItem "Exit application"
   cboSounds.AddItem "Ok"
   cboSounds.AddItem "Ready"
   cboSounds.AddItem "Sorry, ..."
   cboSounds.AddItem "Standby"
   cboSounds.AddItem "Startup"
   cboSounds.AddItem "Thank you"
   cboSounds.ListIndex = 0

   chkPlayWaves.Value = IIf(GetIniString(sIniFile, "Options", "WaveSounds", "1") = "1", vbChecked, vbUnchecked)
   chkMinimize.Value = IIf(GetIniString(sIniFile, "Options", "Minimize", "1") = "1", vbChecked, vbUnchecked)

   ' Read INI file and set the recent menu file list control array appropriately.
   GetRecentFiles

   RestoreFromINI
   ButtonsState
   ShowPrinterInfo

   bBinding = False

   CentreForm Me
   Me.Show
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If cmdSave.Enabled Then
      Select Case MsgBox("You made some changes to the settings. Do you wish to save these settings?", vbYesNoCancel + vbQuestion + vbDefaultButton2, "Save Settings")
      Case vbYes
         cmdSave_Click
      Case vbCancel
         Cancel = True
         Exit Sub
      End Select
   End If

   On Error Resume Next
   Unload frmPrint
   Unload frmPreview

   MakeSound WAVE_EXIT, True
   End                  ' Just in case anything is still running...
End Sub

Private Sub cmdAbout_Click()
   frmAbout.Show vbModal
End Sub

Private Sub cmdHelp_Click()
   On Error Resume Next
   MousePointer = vbHourglass

   Load frmViewFile
   frmViewFile.ShowHelpFile

   MousePointer = vbDefault

   frmViewFile.Show
End Sub

Private Sub lblPrinter_DblClick()
   SSTab.Tab = 1
   TabOptions.Tab = 3
   If optOutput(1) Then
      optOutput(1).SetFocus
   ElseIf optOutput(2) Then
      optOutput(2).SetFocus
   Else
      optOutput(0).SetFocus
   End If
End Sub

Private Sub lblViewSize_DblClick()
   SSTab.Tab = 1
   TabOptions.Tab = 2
   cboZoom.SetFocus
End Sub

' --------------------------------------------------------
Private Sub cmdSave_Click()
   Dim i As Integer
   Dim sText As String
   Dim aFonts As Variant

   On Error GoTo SaveError
   AddIniString sIniFile, "Options", "Header", IIf(chkHeader = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "Footer", IIf(chkFooter = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "PageNumbers", IIf(chkPageNumbers = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "ResetPage", IIf(chkResetPage = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "DateStamp", IIf(chkDate = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "TimeStamp", IIf(chkTime = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "Index", IIf(chkIndex = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "SortIndex", IIf(chkSortIndex = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "ProjectInfo", IIf(chkProject = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "AllIcons", IIf(chkFormIcons = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "FormIcon", IIf(chkIcon = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "ControlNames", IIf(chkControlNames = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "SortControls", IIf(chkSortControls = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "ControlPage", IIf(chkControlPage = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "Code", IIf(chkCode = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "NamesOnly", IIf(chkProcNames = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "ProcPerPage", IIf(chkProcPage = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "SubSeparator", IIf(chkSeparator = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "PagePos", IIf(optPagePos(1), "Footer", "Header")
   AddIniString sIniFile, "Options", "PreviewZoom", cboZoom.Text
   AddIniString sIniFile, "Options", "Extention", cboExtention.ListIndex

   AddIniString sIniFile, "Print", "Device", IIf(optOutput(1), "File", IIf(optOutput(2), "Port", "Driver"))
   AddIniString sIniFile, "Print", "RTFFile", txtRTFfile
   AddIniString sIniFile, "Print", "Port", cboPort
   AddIniString sIniFile, "Print", "Width", cboWidth
   AddIniString sIniFile, "Print", "Height", cboHeight
   AddIniString sIniFile, "Print", "FormFeed", IIf(chkFormFeed = vbChecked, "1", "0")

   AddIniString sIniFile, "Margins", "Top", lblTop(0)
   AddIniString sIniFile, "Margins", "Bottom", lblBottom(0)
   AddIniString sIniFile, "Margins", "Left", lblLeft(0)
   AddIniString sIniFile, "Margins", "Right", lblRight(0)

   AddIniString sIniFile, "Footer", "Line1", txtOwner(0).Text
   AddIniString sIniFile, "Footer", "Line2", txtOwner(1).Text

   aFonts = Array("Procs", "Code", "Comments", "Header", "Footer", "Directive", "Titles")

   For i = 0 To 6
      sText = "Font" & aFonts(i)

      AddIniString sIniFile, sText, "Font", lblFont(i).FontName
      AddIniString sIniFile, sText, "Size", lblFont(i).FontSize
      AddIniString sIniFile, sText, "Color", "&H" & Right("00000000" & Hex$(lblFont(i).ForeColor), 8)
      AddIniString sIniFile, sText, "Bold", IIf(lblFont(i).FontBold, "1", "0")
      AddIniString sIniFile, sText, "Italic", IIf(lblFont(i).FontItalic, "1", "0")
      AddIniString sIniFile, sText, "Strikethru", IIf(lblFont(i).FontStrikethru, "1", "0")
      AddIniString sIniFile, sText, "Underline", IIf(lblFont(i).FontUnderline, "1", "0")
   Next

   SetEnabled cmdSave, False
   SetEnabled cmdRestore, False

   MakeSound WAVE_OK, True
   Exit Sub
SaveError:
   MsgBox "Problem saving current preference settings." & vbCrLf & _
          "Error reported #" & Err.Number & " - " & Err.Description, vbCritical, "Save Error"
End Sub

Private Sub cmdRestore_Click()
   If MsgBox("Current settings will be lost! Are you sure to restore previously saved settings?", vbYesNo + vbQuestion, "Restore Settings") = vbYes Then
      RestoreFromINI
      SetMargins
   End If
End Sub

Private Sub RestoreFromINI()
   Dim sText As String
   Dim nNumber As Integer, i As Integer
   Dim aFonts As Variant
   On Error Resume Next

   chkHeader = IIf(GetIniString(sIniFile, "Options", "Header", IIf(chkHeader = vbChecked, "1", "0")) = "1", 1, 0)
   chkFooter = IIf(GetIniString(sIniFile, "Options", "Footer", IIf(chkFooter = vbChecked, "1", "0")) = "1", 1, 0)
   chkPageNumbers = IIf(GetIniString(sIniFile, "Options", "PageNumbers", IIf(chkPageNumbers = vbChecked, "1", "0")) = "1", 1, 0)
   chkResetPage = IIf(GetIniString(sIniFile, "Options", "resetPage", IIf(chkResetPage = vbChecked, "1", "0")) = "1", 1, 0)
   chkDate = IIf(GetIniString(sIniFile, "Options", "DateStamp", IIf(chkDate = vbChecked, "1", "0")) = "1", 1, 0)
   chkTime = IIf(GetIniString(sIniFile, "Options", "TimeStamp", IIf(chkTime = vbChecked, "1", "0")) = "1", 1, 0)
   chkIndex = IIf(GetIniString(sIniFile, "Options", "Index", IIf(chkIndex = vbChecked, "1", "0")) = "1", 1, 0)
   chkSortIndex = IIf(GetIniString(sIniFile, "Options", "SortIndex", IIf(chkSortIndex = vbChecked, "1", "0")) = "1", 1, 0)
   chkProject = IIf(GetIniString(sIniFile, "Options", "ProjectInfo", IIf(chkProject = vbChecked, "1", "0")) = "1", 1, 0)
   chkFormIcons = IIf(GetIniString(sIniFile, "Options", "AllIcons", IIf(chkFormIcons = vbChecked, "1", "0")) = "1", 1, 0)
   chkIcon = IIf(GetIniString(sIniFile, "Options", "FormIcon", IIf(chkIcon = vbChecked, "1", "0")) = "1", 1, 0)
   chkControlNames = IIf(GetIniString(sIniFile, "Options", "ControlNames", IIf(chkControlNames = vbChecked, "1", "0")) = "1", 1, 0)
   chkSortControls = IIf(GetIniString(sIniFile, "Options", "SortControls", IIf(chkSortControls = vbChecked, "1", "0")) = "1", 1, 0)
   chkControlPage = IIf(GetIniString(sIniFile, "Options", "ControlPage", IIf(chkControlPage = vbChecked, "1", "0")) = "1", 1, 0)
   chkCode = IIf(GetIniString(sIniFile, "Options", "Code", IIf(chkCode = vbChecked, "1", "0")) = "1", 1, 0)
   chkProcNames = IIf(GetIniString(sIniFile, "Options", "NamesOnly", IIf(chkProcNames = vbChecked, "1", "0")) = "1", 1, 0)
   chkProcPage = IIf(GetIniString(sIniFile, "Options", "ProcPerPage", IIf(chkProcPage = vbChecked, "1", "0")) = "1", 1, 0)
   chkSeparator = IIf(GetIniString(sIniFile, "Options", "SubSeparator", IIf(chkSeparator = vbChecked, "1", "0")) = "1", 1, 0)
   If UCase$(Left(GetIniString(sIniFile, "Options", "PagePos", "Header"), 1)) = "F" Then
      optPagePos(1) = True
   Else
      optPagePos(0) = True
   End If
   nNumber = Int(Val(GetIniString(sIniFile, "Options", "PreviewZoom", "100%")))
   If nNumber <= 0 Then
      cboZoom.Text = "Fit"
   Else
      cboZoom.Text = Format(nNumber, "##0") & "%"
   End If
   nNumber = Val(GetIniString(sIniFile, "Options", "Extention", cboExtention.ListIndex))
   cboExtention.ListIndex = IIf(nNumber < 0 Or nNumber > 8, 0, nNumber)

   sText = UCase$(Left(GetIniString(sIniFile, "Print", "Device", "Driver"), 1))
   If sText = "F" Then
      optOutput(1) = True
   ElseIf sText = "P" Then
      optOutput(2) = True
   Else
      optOutput(0) = True
   End If
   txtRTFfile = GetIniString(sIniFile, "Print", "RTFFile", txtRTFfile)
   cboPort = GetIniString(sIniFile, "Print", "Port", cboPort)
   cboWidth = GetIniString(sIniFile, "Print", "Width", cboWidth)
   cboHeight = GetIniString(sIniFile, "Print", "Height", cboHeight)
   chkFormFeed = IIf(GetIniString(sIniFile, "Print", "FormFeed", IIf(chkFormFeed = vbChecked, "1", "0")) = "1", 1, 0)

   lblTop(0) = GetIniString(sIniFile, "Margins", "Top", lblTop(0))
   lblBottom(0) = GetIniString(sIniFile, "Margins", "Bottom", lblBottom(0))
   lblLeft(0) = GetIniString(sIniFile, "Margins", "Left", lblLeft(0))
   lblRight(0) = GetIniString(sIniFile, "Margins", "Right", lblRight(0))
   AssignSliderValues

   txtOwner(0).Text = GetIniString(sIniFile, "Footer", "Line1", "Created with 'VB.Print!' - VB Source code printing utility")
   txtOwner(1).Text = GetIniString(sIniFile, "Footer", "Line2", "'VB.Print!' is written by Inner Control Business Management (Australia)")

   aFonts = Array("Procs", "Code", "Comments", "Header", "Footer", "Directive", "Titles")

   For i = 0 To 6
      sText = "Font" & aFonts(i)

      lblFont(i).FontName = GetIniString(sIniFile, sText, "Font", lblFont(i).FontName)
      lblFont(i).FontSize = Val(GetIniString(sIniFile, sText, "Size", lblFont(i).FontSize))
      lblFont(i).ForeColor = Abs(Val(GetIniString(sIniFile, sText, "Color", "&H" & Hex(lblFont(i).ForeColor))))
      lblFont(i).FontBold = (GetIniString(sIniFile, sText, "Bold", IIf(lblFont(i).FontBold, "1", "0")) = "1")
      lblFont(i).FontItalic = (GetIniString(sIniFile, sText, "Italic", IIf(lblFont(i).FontItalic, "1", "0")) = "1")
      lblFont(i).FontStrikethru = (GetIniString(sIniFile, sText, "Strikethru", IIf(lblFont(i).FontStrikethru, "1", "0")) = "1")
      lblFont(i).FontUnderline = (GetIniString(sIniFile, sText, "Underline", IIf(lblFont(i).FontUnderline, "1", "0")) = "1")
   Next

   SetEnabled cmdSave, False
   SetEnabled cmdRestore, False

End Sub

Private Sub chkPlayWaves_Click()
   AddIniString sIniFile, "Options", "WaveSounds", IIf(chkPlayWaves = vbChecked, "1", "0")
End Sub

Private Sub chkMinimize_Click()
   AddIniString sIniFile, "Options", "Minimize", IIf(chkMinimize = vbChecked, "1", "0")
End Sub

Private Sub cmdEditor_Click()
   On Error GoTo EditorCancelled

   CommonDialog.DialogTitle = "Open"
   CommonDialog.Filter = "Rich Text Format (*.rtf)|*.rtf|Text document (*.txt)|*.txt|All files (*.*)|*.*"
   CommonDialog.FilterIndex = 1

   CommonDialog.CancelError = True
   CommonDialog.Flags = cdlOFNHideReadOnly
   CommonDialog.FileName = ""

   CommonDialog.ShowOpen

   CommonDialog.CancelError = False

   MousePointer = vbHourglass

   Load frmViewFile
   frmViewFile.SetFileName CommonDialog.FileName
   frmViewFile.InitView

   MousePointer = vbDefault

   frmViewFile.Show

EditorCancelled:
   CommonDialog.CancelError = False
End Sub

Private Sub cmdView_Click()
   On Error Resume Next

   If Outline.ListCount < 1 Or Outline.ListIndex < 0 Then Exit Sub

   Dim sText As String
   Dim i As Integer, n As Integer

   MousePointer = vbHourglass

   Load frmViewFile

   Select Case ItemRef(Outline.ListIndex).ProcPoint
   Case Is = 0
      sText = ""
      For i = 1 To Mdl(ItemRef(Outline.ListIndex).FilePoint).CtrlCount

         If Mdl(ItemRef(Outline.ListIndex).FilePoint).Control(i).Elements > 1 Then
            sText = sText & Pad(Mdl(ItemRef(Outline.ListIndex).FilePoint).Control(i).Name, 20) & " " & _
                            Pad(Mdl(ItemRef(Outline.ListIndex).FilePoint).Control(i).Type, 20) & " " & _
                            Pad(Mdl(ItemRef(Outline.ListIndex).FilePoint).Control(i).Library, 20) & " " & _
                            "Elements: " & Mdl(ItemRef(Outline.ListIndex).FilePoint).Control(i).Elements
         Else
            sText = sText & Pad(Mdl(ItemRef(Outline.ListIndex).FilePoint).Control(i).Name, 20) & " " & _
                            Pad(Mdl(ItemRef(Outline.ListIndex).FilePoint).Control(i).Type, 20) & " " & _
                            Mdl(ItemRef(Outline.ListIndex).FilePoint).Control(i).Library
         End If
         sText = sText & vbCrLf
      Next
      sText = sText & vbCrLf & _
              "   Total control names: " & Mdl(ItemRef(Outline.ListIndex).FilePoint).CtrlCount & vbCrLf & _
              "Total control elements: " & Mdl(ItemRef(Outline.ListIndex).FilePoint).CtrlElements

      frmViewFile.SetText "Form Controls", sText
    
   Case Is > 0
      Dim sUpper As String
      sText = ""
      n = 1       ' Remove empty line(s) in top
      For i = 1 To Mdl(ItemRef(Outline.ListIndex).FilePoint).Proc(ItemRef(Outline.ListIndex).ProcPoint).Lines
         If Not EmptyString(Mdl(ItemRef(Outline.ListIndex).FilePoint).Proc(ItemRef(Outline.ListIndex).ProcPoint).Code(i)) Then
            n = i
            Exit For
         End If
      Next
      frmViewFile.Caption = Mdl(ItemRef(Outline.ListIndex).FilePoint).Proc(ItemRef(Outline.ListIndex).ProcPoint).IndexName & " - View"
      frmViewFile.SetFont FONT_CODE

      For i = n To Mdl(ItemRef(Outline.ListIndex).FilePoint).Proc(ItemRef(Outline.ListIndex).ProcPoint).Lines
         
         sText = Mdl(ItemRef(Outline.ListIndex).FilePoint).Proc(ItemRef(Outline.ListIndex).ProcPoint).Code(i) & vbCrLf
         sUpper = UCase$(Trim$(sText))

         If MatchString(sUpper, "'") Then                      ' Comments
            frmViewFile.SetFont FONT_COMMENTS
            frmViewFile.SetLine sText
            frmViewFile.SetFont FONT_CODE

         ElseIf MatchString(sUpper, "#") Then                  ' Compiler directive
            frmViewFile.SetFont FONT_DIRECTIVE
            frmViewFile.SetLine sText
            frmViewFile.SetFont FONT_CODE

         ElseIf IsProcedure(sUpper) Then                       ' Only happens once (I hope) - and not in declaration section
            frmViewFile.SetFont FONT_PROCS
            frmViewFile.SetLine sText
            frmViewFile.SetFont FONT_CODE
         Else                                                  ' Just some code or space
            frmViewFile.SetLine sText
         End If
      Next
 
   Case Else
      frmViewFile.SetFileName Mdl(ItemRef(Outline.ListIndex).FilePoint).PathFile
   End Select

   frmViewFile.InitView

   MousePointer = vbDefault

   frmViewFile.Show
End Sub

' --------------------------------------------------------

Private Sub cboSounds_Change()
   cboSounds_Click
End Sub

Private Sub cboSounds_Click()
   lblSoundFile.Caption = "Wave file: " & GetSoundFileName(cboSounds.ListIndex)
End Sub

Private Sub cmdPlay_Click()
   If bBinding Then Exit Sub
   If cboSounds.ListIndex < 0 Then Exit Sub
   MakeSound cboSounds.ListIndex, False, True
End Sub

' --------------------------------------------------------

' Test layout of text printer
Private Sub cmdTextTest_Click()
   On Error GoTo TestPrintError

   SetEnabled cmdTextTest, False

   Me.MousePointer = vbHourglass
   DoEvents
   frmMain.Enabled = False
   On Error GoTo 0

   MakeSound WAVE_STANDBY

   TestTextPrint

TestPrintError:
   frmMain.Enabled = True
   SetEnabled cmdTextTest, True
   Me.MousePointer = vbDefault
End Sub

Private Sub cmdPrintSetup_Click()
   CommonDialog.CancelError = False
   CommonDialog.Flags = cdlPDPrintSetup
   CommonDialog.ShowPrinter

   DoEvents
   ShowPrinterInfo
End Sub

Public Sub SetRTFfile(sFile As String)
   bBinding = True
   frmMain.txtRTFfile = sFile
   bBinding = False
End Sub

Private Sub cmdPrint_Click()
   Dim nFormState As Integer
   nFormState = Me.WindowState

   On Error GoTo PrintDialogCancelled

   SetEnabled cmdPrint, False

   If optOutput(1) Then                ' RTF
      Dim sRTFFile As String
      Me.MousePointer = vbHourglass
      DoEvents
      frmMain.Enabled = False
      On Error GoTo 0

      If frmMain.chkPreview <> vbChecked Then
         sRTFFile = frmMain.txtRTFfile
         If Not FileOverwriteDialog(sRTFFile, CommonDialog, "RTF files (*.rtf)|*.rtf|All files (*.*)|*.*", ".rtf") Then
            GoTo PrintDialogCancelled
         End If
         SetRTFfile sRTFFile
      End If
   
   ElseIf optOutput(2) Then            ' Port
      Me.MousePointer = vbHourglass
      DoEvents
      frmMain.Enabled = False
      On Error GoTo 0

   Else
      If chkPreview <> vbChecked Then
         CommonDialog.Flags = cdlPDHidePrintToFile Or cdlPDNoSelection Or cdlPDUseDevModeCopies Or cdlPDPageNums 'Or cdlPDPageNums 'Or cdlPDNoPageNums
'         CommonDialog.FromPage = txtFromPage
'         CommonDialog.ToPage = txtToPage
'         CommonDialog.Min = 1
'         CommonDialog.Max = 1

         CommonDialog.FromPage = 1
         CommonDialog.ToPage = 1
         CommonDialog.Min = 1
         CommonDialog.Max = 1

         CommonDialog.CancelError = True

         CommonDialog.ShowPrinter

         CommonDialog.CancelError = False
      End If

      Me.MousePointer = vbHourglass
      DoEvents
      If chkPreview <> vbChecked Then ShowPrinterInfo
      frmMain.Enabled = False
      On Error GoTo 0
   End If

   MakeSound WAVE_STANDBY

   If GetIniString(sIniFile, "Options", "Minimize", "1") = "1" Then
      WindowState = vbMinimized
   End If

   ' This is it - The whole application turns around this routine.
   PrintControl

PrintDialogCancelled:
   lstNames(0).Clear
   lstNames(1).Clear
   picImage.Picture = LoadPicture()
   frmMain.Enabled = True
   SetEnabled cmdPrint, True
   CommonDialog.CancelError = False
   Me.MousePointer = vbDefault

   If WindowState = vbMinimized Then
      DoEvents
      Do While Page.Show: DoEvents: Loop
      If WindowState = vbMinimized Then WindowState = nFormState
   End If
End Sub

Private Sub ShowPrinterInfo(Optional bRefreshSample)
   Me.MousePointer = vbHourglass

   If optOutput(2) Then
      lblPrinter = "Plain text direct to port " & cboPort
      lblOrient = "n/a"
      lblSize = Format(cboWidth, "###0") & " (chars) by " & Format(cboHeight, "###0") & " (lines)"

   Else
      If optOutput(1) Then
         lblPrinter = "RTF file " & txtRTFfile
      Else
         lblPrinter = Printer.DeviceName & " on " & Printer.Port
      End If

      Select Case Printer.Orientation
      Case vbPRORPortrait
         lblOrient = "Portrait"
      Case vbPRORLandscape
         lblOrient = "Landscape"
      Case Else
         lblOrient = "n/a"
      End Select

      'Printer.ScaleLeft, Printer.ScaleTop
      Printer.ScaleMode = vbMillimeters
      If optOutput(1) Then
         lblSize = Format(Printer.ScaleWidth - (Val(lblLeft(0)) + Val(lblRight(0))), "###0") & " mm by " & Format(Printer.ScaleHeight - (Val(lblTop(0)) + Val(lblBottom(0))), "###0") & " mm"
      Else
         lblSize = Format(Printer.ScaleWidth, "###0") & " mm by " & Format(Printer.ScaleHeight, "###0") & " mm"
      End If
   End If

   If IsMissing(bRefreshSample) Then bRefreshSample = True

   If bRefreshSample Then
      ' Don't change the execution order !!
      SetSampleOrientation
      SetSliders
      If Page.Show Then
         PaintSamplePage
      Else
         PrintSamplePage
      End If
   End If

   Me.MousePointer = vbDefault
End Sub

' --------------------------------------------------------

Private Sub txtProject_Change()
   If Len(Trim(txtProject)) = 0 Then
      If Outline.ListCount > 0 Then
         ClearOutline
         ButtonsState
      End If
   Else
      If Not FileExist(txtProject) Then
         If Outline.ListCount > 0 Then
            ClearOutline
            ButtonsState
         End If
      End If
   End If
End Sub

Private Sub txtProject_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txtProject)) > 0 Then
         GetProjectDetails
      End If
   End If
End Sub

Private Sub cmdPrevFile_Click()
'   Dim sText As String
'   sText = GetIniString(sIniFile, "Options", "LastFile", "")
'   If Len(Trim(sText)) > 0 Then
'      txtProject = sText
'      GetProjectDetails
'   End If
   PopupMenu mnuFiles
End Sub

Private Sub mnuRecentFile_Click(Index As Integer)
   If Len(Trim$(mnuRecentFile(Index).Caption)) > 0 Then
      If txtProject <> mnuRecentFile(Index).Caption Then
         txtProject = mnuRecentFile(Index).Caption
         GetProjectDetails
      End If
   End If
End Sub

Private Sub cmdPickFile_Click()
   On Error GoTo PickFileCancelled

   CommonDialog.DialogTitle = "Open project, form, module or class file"
   CommonDialog.Filter = "VB files (*.vbp;*.frm;*.bas;*.cls;*.ctl;*.pag;*.dob)|*.vbp;*.frm;*.bas;*.cls;*.ctl;*.pag;*.dob|" & _
                         "Project files (*.vbp)|*.vbp|" & _
                         "Form files (*.frm)|*.frm|" & _
                         "Module files (*.bas)|*.bas|" & _
                         "Class files (*.cls)|*.cls|" & _
                         "User Control files (*.ctl)|*.ctl|" & _
                         "Property Page files (*.pag)|*.pag|" & _
                         "User Document files (*.dob)|*.dob|" & _
                         "All files (*.*)|*.*"

   CommonDialog.FilterIndex = cboExtention.ListIndex + 1

   CommonDialog.CancelError = True
   CommonDialog.Flags = cdlOFNHideReadOnly
   CommonDialog.FileName = txtProject

   CommonDialog.ShowOpen

   CommonDialog.CancelError = False

   txtProject = CommonDialog.FileName

   GetProjectDetails

PickFileCancelled:
   CommonDialog.CancelError = False
End Sub

Private Sub lstForms_Click()
   ButtonsState
End Sub

' --- Preferences tab area -----------------------------------------------------

Private Sub SSTab_Click(PreviousTab As Integer)
   SetVisible cmdSave, (SSTab.Tab = 1)
   SetVisible cmdRestore, (SSTab.Tab = 1)
End Sub

Private Sub chkHeader_Click()
   SetEnabled chkPageNumbers, (chkHeader = vbChecked Or chkFooter = vbChecked)
   SetEnabled chkResetPage, (chkPageNumbers.Enabled And (chkPageNumbers = vbChecked))
   SetEnabled optPagePos(0), (chkPageNumbers.Enabled And (chkPageNumbers = vbChecked))
   SetEnabled optPagePos(1), (chkPageNumbers.Enabled And (chkPageNumbers = vbChecked))
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkFooter_Click()
   SetEnabled chkPageNumbers, (chkHeader = vbChecked Or chkFooter = vbChecked)
   SetEnabled chkResetPage, (chkPageNumbers.Enabled And (chkPageNumbers = vbChecked))
   SetEnabled optPagePos(0), (chkPageNumbers.Enabled And (chkPageNumbers = vbChecked))
   SetEnabled optPagePos(1), (chkPageNumbers.Enabled And (chkPageNumbers = vbChecked))
   SetEnabled chkDate, (chkFooter = vbChecked)
   SetEnabled chkTime, (chkFooter = vbChecked)
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkPageNumbers_Click()
   SetEnabled chkResetPage, (chkPageNumbers = vbChecked)
   SetEnabled optPagePos(0), (chkPageNumbers = vbChecked)
   SetEnabled optPagePos(1), (chkPageNumbers = vbChecked)
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkResetPage_Click()
   EnabledStorage
End Sub

Private Sub optPagePos_Click(Index As Integer)
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkDate_Click()
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkTime_Click()
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkIndex_Click()
   SetEnabled chkSortIndex, (chkIndex = vbChecked)
   EnabledStorage
End Sub

Private Sub chkSortIndex_Click()
   EnabledStorage
End Sub

Private Sub chkProject_Click()
   ButtonsState
   EnabledStorage
End Sub

Private Sub chkFormIcons_Click()
   EnabledStorage
End Sub

Private Sub chkIcon_Click()
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkControlNames_Click()
   SetEnabled chkSortControls, (chkControlNames = vbChecked)
   SetEnabled chkControlPage, (chkControlNames = vbChecked)
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkSortControls_Click()
   EnabledStorage
End Sub

Private Sub chkControlPage_Click()
   EnabledStorage
End Sub

Private Sub chkCode_Click()
   SetEnabled chkProcNames, (chkCode = vbChecked)
   SetEnabled chkProcPage, (chkCode = vbChecked)
   SetEnabled chkSeparator, (chkCode = vbChecked)
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkProcNames_Click()
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkProcPage_Click()
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkSeparator_Click()
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub cboZoom_Click()
   cboZoom_Change
End Sub

Private Sub cboZoom_Change()
   If bZoomRefresh Then Exit Sub

   Dim nFactor As Single
   nFactor = Int(Val(cboZoom.Text))
   If nFactor <= 0 Then
      lblViewSize = "(Zoom to Fit)"
   Else
      nFactor = nFactor / 100
      lblViewSize = "(Zoom " & Format(nFactor * 100, "##0") & "%)"
   End If
   EnabledStorage
End Sub

Private Sub cboZoom_LostFocus()
   Dim nFactor As Single
   nFactor = Int(Val(cboZoom.Text))

   bZoomRefresh = True
   If nFactor <= 0 Then
      cboZoom.Text = "Fit"
   Else
      nFactor = nFactor / 100
      cboZoom.Text = Format(nFactor * 100, "##0") & "%"
   End If
   bZoomRefresh = False
End Sub

Private Sub txtOwner_LostFocus(Index As Integer)
   PaintSamplePage
End Sub

Private Sub txtOwner_Change(Index As Integer)
   EnabledStorage
End Sub

Private Sub optOutput_Click(Index As Integer)
   If Index = 2 Then
      ' Port
      SetEnabled txtRTFfile, False
      SetEnabled cmdPickRtf, False

      SetEnabled chkFormIcons, False
      SetEnabled chkIcon, False

      SetEnabled lblOutput(0), True
      SetEnabled lblOutput(1), True
      SetEnabled lblOutput(2), True
      SetEnabled lblOutput(3), True
      SetEnabled lblOutput(4), True
      SetEnabled cboPort, True
      SetEnabled cboWidth, True
      SetEnabled cboHeight, True
      SetEnabled chkFormFeed, True

      SetEnabled lblPort(0), True
      SetEnabled lblPort(1), True
      SetEnabled lblPort(2), True
      SetEnabled lblPort(3), True
      SetEnabled lblPort(4), True
      SetEnabled lblPort(5), True
      SetEnabled lblPort(6), True

      SetEnabled lblTop(1), True
      SetEnabled lblBottom(1), True
      SetEnabled lblLeft(1), True
      SetEnabled lblRight(1), True

      SetEnabled cmdTextTest, True

   Else
      If Index = 1 Then
         ' RTF
         SetEnabled txtRTFfile, True
         SetEnabled cmdPickRtf, True

         SetEnabled chkFormIcons, False
         SetEnabled chkIcon, False

      Else
         ' Driver
         SetEnabled txtRTFfile, False
         SetEnabled cmdPickRtf, False

         SetEnabled chkFormIcons, True
         SetEnabled chkIcon, True
      End If

      SetEnabled lblOutput(0), False
      SetEnabled lblOutput(1), False
      SetEnabled lblOutput(2), False
      SetEnabled lblOutput(3), False
      SetEnabled lblOutput(4), False
      SetEnabled cboPort, False
      SetEnabled cboWidth, False
      SetEnabled cboHeight, False
      SetEnabled chkFormFeed, False

      SetEnabled lblPort(0), False
      SetEnabled lblPort(1), False
      SetEnabled lblPort(2), False
      SetEnabled lblPort(3), False
      SetEnabled lblPort(4), False
      SetEnabled lblPort(5), False
      SetEnabled lblPort(6), False

      SetEnabled lblTop(1), False
      SetEnabled lblBottom(1), False
      SetEnabled lblLeft(1), False
      SetEnabled lblRight(1), False

      SetEnabled cmdTextTest, False
   End If

   If bBinding Then Exit Sub
   ShowPrinterInfo False
   ButtonsState
   EnabledStorage
End Sub

Private Sub cmdPickRtf_Click()
   On Error GoTo PickRTFCancelled

   CommonDialog.DialogTitle = "Save RTF file as ..."
   CommonDialog.Filter = "RTF files (*.rtf)|*.rtf|All files (*.*)|*.*"
   CommonDialog.FilterIndex = 1
   CommonDialog.DefaultExt = ".rtf"

   CommonDialog.CancelError = True
   CommonDialog.Flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist

   CommonDialog.FileName = txtRTFfile

   CommonDialog.ShowSave

   txtRTFfile = CommonDialog.FileName

PickRTFCancelled:
   CommonDialog.CancelError = False
End Sub

Private Sub txtRTFfile_Change()
   If bBinding Then Exit Sub
   ShowPrinterInfo False
   EnabledStorage
End Sub

Private Sub cboPort_Click()
   If bBinding Then Exit Sub
   ShowPrinterInfo False
   EnabledStorage
End Sub

Private Sub cboPort_Change()
   If bBinding Then Exit Sub
   ShowPrinterInfo False
   EnabledStorage
End Sub

Private Sub cboHeight_KeyPress(KeyAscii As Integer)
   KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub cboHeight_Click()
   If bBinding Then Exit Sub
   ShowPrinterInfo False
   EnabledStorage
End Sub

Private Sub cboHeight_Change()
   If bBinding Then Exit Sub
   ShowPrinterInfo False
   EnabledStorage
End Sub

Private Sub cboWidth_KeyPress(KeyAscii As Integer)
   KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub cboWidth_Click()
   If bBinding Then Exit Sub
   ShowPrinterInfo False
   EnabledStorage
End Sub

Private Sub cboWidth_Change()
   If bBinding Then Exit Sub
   ShowPrinterInfo False
   EnabledStorage
End Sub

Private Sub chkFormFeed_Click()
   EnabledStorage
End Sub

Private Sub cboExtention_Click()
   EnabledStorage
End Sub

Private Sub sldLeft_Change()
   EnabledStorage
   SetMargins
End Sub

Private Sub sldRight_Change()
   EnabledStorage
   SetMargins
End Sub

Private Sub sldTop_Change()
   EnabledStorage
   SetMargins
End Sub

Private Sub sldBottom_Change()
   EnabledStorage
   SetMargins
End Sub

' Only use third of width and length of page for margins
Private Sub SetSliders()
   Dim nMax As Integer, nValue As Integer

   Printer.ScaleMode = vbMillimeters

   nMax = Printer.ScaleWidth / 3
   sldLeft.Max = nMax

   nValue = sldRight.Max - sldRight.Value
   sldRight.Max = nMax
   sldRight.Value = nMax - nValue

   nMax = Printer.ScaleHeight / 3
   sldTop.Max = nMax

   nValue = sldBottom.Max - sldBottom.Value
   sldBottom.Max = nMax
   sldBottom.Value = nMax - nValue
End Sub

Private Sub AssignSliderValues()
   sldLeft.Value = lblLeft(0)
   sldRight.Value = sldRight.Max - lblRight(0)
   sldTop.Value = lblTop(0)
   sldBottom.Value = sldBottom.Max - lblBottom(0)

   SetTextMargins
End Sub

' Margins are in millimeters
Private Sub SetMargins()
   If bBinding Then Exit Sub
   lblLeft(0) = sldLeft.Value
   lblRight(0) = sldRight.Max - sldRight.Value
   lblTop(0) = sldTop.Value
   lblBottom(0) = sldBottom.Max - sldBottom.Value

   If optOutput(1) Then
      lblSize = Format(Printer.ScaleWidth - (Val(lblLeft(0)) + Val(lblRight(0))), "###0") & " mm by " & Format(Printer.ScaleHeight - (Val(lblTop(0)) + Val(lblBottom(0))), "###0") & " mm"
   End If

   SetTextMargins

   PaintSamplePage
End Sub

' The source margins are millimeters. Convert them to characters
Private Sub SetTextMargins()
   Dim nCharFactor As Double, nLineFactor As Double
   nCharFactor = 1200 / 567
   nLineFactor = 2400 / 567
   lblTop(1) = RoundToInt(lblTop(0) / nLineFactor)
   lblBottom(1) = RoundToInt(lblBottom(0) / nLineFactor)
   lblLeft(1) = RoundToInt(lblLeft(0) / nCharFactor)
   lblRight(1) = RoundToInt(lblRight(0) / nCharFactor)
End Sub

' Only accepts positive numbers
Private Function RoundToInt(ByVal nValue As Double) As Integer
   If nValue < 0 Then
      RoundToInt = CInt(nValue)
   Else
      Dim nFraction As Double
      nFraction = nValue - Int(nValue)
      If nFraction < 0.5 Then
         RoundToInt = Int(nValue)
      Else
         RoundToInt = Int(nValue) + 1
      End If
   End If
End Function

Private Sub cmdFont_Click(Index As Integer)
   
   On Error GoTo FontSelectCancel

   CommonDialog.CancelError = True
   CommonDialog.Color = lblFont(Index).ForeColor
   CommonDialog.FontBold = lblFont(Index).FontBold
   CommonDialog.FontItalic = lblFont(Index).FontItalic
   CommonDialog.FontStrikethru = lblFont(Index).FontStrikethru
   CommonDialog.FontUnderline = lblFont(Index).FontUnderline
   CommonDialog.FontName = lblFont(Index).FontName
   CommonDialog.FontSize = lblFont(Index).FontSize
   CommonDialog.Flags = cdlCFEffects Or cdlCFForceFontExist Or cdlCFPrinterFonts Or cdlCFScalableOnly ' Or cdlCFBoth

   CommonDialog.ShowFont

   lblFont(Index).ForeColor = CommonDialog.Color
   lblFont(Index).FontBold = CommonDialog.FontBold
   lblFont(Index).FontItalic = CommonDialog.FontItalic
   lblFont(Index).FontStrikethru = CommonDialog.FontStrikethru
   lblFont(Index).FontUnderline = CommonDialog.FontUnderline
   lblFont(Index).FontName = CommonDialog.FontName
   lblFont(Index).FontSize = CommonDialog.FontSize

   EnabledStorage

   ' Paint the sample page
   PaintSamplePage

FontSelectCancel:
   CommonDialog.CancelError = False
End Sub

Private Sub EnabledStorage()
   If bBinding Then Exit Sub
   SetEnabled cmdSave, True
   SetEnabled cmdRestore, FileExist(sIniFile)
End Sub

' Factor change value: 765
Private Sub SetSampleOrientation()
   Dim nFactor As Integer, nScale As Integer

   If Printer.Orientation = vbPRORLandscape Then   ' Landscape
      nFactor = 765
      nScale = 270
   Else
      nFactor = 0
      nScale = 0
   End If

   picPage.Width = 2100 + nFactor
   picPage.Height = 2865 - nFactor

   'object.Move Left, Top, Width, Heigh
   lblMM.Move 2565 + nFactor, 3390 - nFactor

   sldLeft.Move 60, 3120 - nFactor, 915 + nScale, 225
   lblLeft(0).Move 150, 3390 - nFactor

   sldRight.Move 1440 + (nFactor - nScale), 3120 - nFactor, 915 + nScale, 225
   lblRight(0).Move 1965 + nFactor, 3390 - nFactor

   sldTop.Move 2265 + nFactor, 165, 225, 1185 - nScale
   lblTop(0).Move 2565 + nFactor, 225

   sldBottom.Move 2265 + nFactor, 2040 - (nFactor - nScale), 225, 1185 - nScale
   lblBottom(0).Move 2565 + nFactor, 2970 - nFactor

End Sub

' Refresh sample now
Private Sub picPage_DblClick()
   If Page.Show Then Exit Sub
   TmrPaint.Enabled = False
   PrintSamplePage
End Sub

' Don't refresh sample just yet - wait for 2 seconds (because it takes a little time to update)
Private Sub PaintSamplePage()
   If bBinding Then Exit Sub
   InvalidateSamplePage
   TmrPaint.Enabled = False
   TmrPaint.Enabled = True
End Sub

' Timer event fired, the 2 seconds must be over - refresh sample page
Private Sub TmrPaint_Timer()
   If Page.Show Then Exit Sub
   TmrPaint.Enabled = False
   PrintSamplePage
End Sub

' --- Status messages (mini help) area -----------------------------------------------------

Private Sub cboExtention_GotFocus()
   lblExtention_MouseMove 0, 0, 0, 0
End Sub

Private Sub cboPort_GotFocus()
   lblOutput_MouseMove 0, 0, 0, 0, 0
End Sub

Private Sub cboHeight_GotFocus()
   lblOutput_MouseMove 2, 0, 0, 0, 0
End Sub

Private Sub cboWidth_GotFocus()
   lblOutput_MouseMove 1, 0, 0, 0, 0
End Sub

Private Sub cboZoom_GotFocus()
   lblZoomDialog_MouseMove 0, 0, 0, 0
End Sub

Private Sub SetStatusText(sText As String)
   If StatusBar.SimpleText <> sText Then StatusBar.SimpleText = sText
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Ready"
End Sub

Private Sub Frame_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Ready"
End Sub

Private Sub Label_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Ready"
End Sub

Private Sub SSTab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Ready"
End Sub

Private Sub TabOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Ready"
End Sub

Private Sub txtProject_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Enter a project, form, module or class file name"
End Sub

Private Sub lblPrinter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Current printer selected - Double click on here to go directly to the selection area"
End Sub

Private Sub lblViewSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Page size of preview - Double click on here to go directly to the selection area"
End Sub

Private Sub lblFont_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Font display sample"
End Sub

Private Sub lstForms_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Select file(s) to print or view"
End Sub

Private Sub picPage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Displays a rough sample of a file listing"
End Sub

Private Sub sldBottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Adjust bottom margin in millimeters"
End Sub

Private Sub sldLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Adjust left margin in millimeters"
End Sub

Private Sub sldRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Adjust right margin in millimeters"
End Sub

Private Sub sldTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Adjust top margin in millimeters"
End Sub

Private Sub chkFormIcons_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Print icons of all forms (in project) on seperate page(s)"
End Sub

Private Sub chkIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Print form icon prior to the control names and code"
End Sub

Private Sub txtOwner_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Footer text (mainly to identify the code owner)"
End Sub

Private Sub chkCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Enable/Disable code printing"
End Sub

Private Sub chkProcPage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "When enabled each procedure will be printed on a seperate page"
End Sub

Private Sub chkControlNames_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Enable/Disable form control object names listing"
End Sub

Private Sub chkControlPage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "When enabled new page when control names are finished"
End Sub

Private Sub chkDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Enable/Disable current system date in footer"
End Sub

Private Sub chkTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Enable/Disable current system time in footer"
End Sub

Private Sub chkFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Enable/Disable footer section (shows time/date and footer text)"
End Sub

Private Sub chkHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Enable/Disable header section (shows file name)"
End Sub

Private Sub chkIndex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Enable/Disable index page (eg. procedures)"
End Sub

Private Sub chkSortIndex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Enable/Disable index sort on names (otherwise by page number)"
End Sub

Private Sub chkSortControls_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Enable/Disable control objects sort on names"
End Sub

Private Sub chkPageNumbers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Enable/Disable page number display"
End Sub

Private Sub chkResetPage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "When enabled resets page number to 1 for each file"
End Sub

Private Sub optPagePos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Index = 0 Then
      SetStatusText "When enabled prints page number in the header section"
   Else
      SetStatusText "When enabled prints page number in the footer section"
   End If
End Sub

Private Sub optOutput_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Index = 1 Then
      SetStatusText "When enabled sends print data to a RTF file (usefull to fancy it up in a wordprocessor)"
   ElseIf Index = 2 Then
      SetStatusText "When enabled sends print data directly to printer port (usefull fast text print-outs)"
   Else
      SetStatusText "When enabled uses the windows printer driver for printing"
   End If
End Sub

Private Sub lblOutput_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Index = 0 Then
      SetStatusText "Specify which printer port the text printer is connected to your computer"
   ElseIf Index = 1 Then
      SetStatusText "Specify page width in characters (mostly 80 for narrow or 132 for wide printers)"
   ElseIf Index = 2 Then
      SetStatusText "Specify page height in lines (A4 format [297 mm] is 70 lines)"
   End If
End Sub

Private Sub lblPort_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Displays the conversion calculation of the margins (from millimeters)"
End Sub

Private Sub cmdTextTest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Usefull to check if your page dimension are proper (Prints to selected printer port)"
End Sub

Private Sub chkFormFeed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Forces a form feed control character (Ascii 12) once page is printed"
End Sub

Private Sub chkProject_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Enable/Disable project information section"
End Sub

Private Sub chkSeparator_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Enable/Disable horinzontal separation line between areas"
End Sub

Private Sub chkPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Toggle between print or preview file listings"
End Sub

Private Sub chkProcNames_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "When enabled only shows procedure names (no code)"
End Sub

Private Sub lblExtention_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Select default file extention in file selection dialog"
End Sub

Private Sub cmdAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Shows the about window"
End Sub

Private Sub cmdClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Resets the selection status of all files"
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Exit VB.Print!"
End Sub

Private Sub cmdFont_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Font change (adjust font, size, colour, attributes)"
End Sub

Private Sub cmdHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "On-line help"
End Sub

Private Sub cmdPickFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Select a file from a list"
End Sub

Private Sub cmdPrevFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Select file from recent files list"
End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Print/Preview listing of selected files"
End Sub

Private Sub cmdPrintSetup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Print setup: Printer, paper and orientation"
End Sub

Private Sub cmdRestore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Cancel changed settings and retrieves last saved settings"
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Save changed settings to file"
End Sub

Private Sub cmdSelectAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Selects all files for printing"
End Sub

Private Sub cmdView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "View the contents of the current selected file"
End Sub

Private Sub chkPlayWaves_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Enable or disable wave sounds to specific events"
End Sub

Private Sub cboZoom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Set the default zoom size of the preview area (magnification factor)"
End Sub

Private Sub lblZoomDialog_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SetStatusText "Default preview page zoom factor as a percentage"
End Sub


' ----------------------------------------------------------------------------------------------

Private Sub ClearOutline()
   MdCount = 0
   PrCount = 0
   MdSelected = 0
   PrSelected = 0
   sLoadedFile = ""
   Outline.Clear
   Erase Mdl
   ReDim ItemRef(0)
End Sub

Private Sub SetOutline(sFile As String, nType As Integer)
   Dim nIndex As Integer, i As Integer, nListIndex As Integer
   Dim sName As String

   nIndex = AnalyseFile(sFile, nType)

   If nIndex > -1 Then

      sName = Mdl(nIndex).File
      If Mdl(nIndex).ProcCount > 0 Then
         sName = sName & "     [" & Mdl(nIndex).ProcCount & "]"
      End If

      ' List File
      Outline.AddItem sName
      nListIndex = Outline.ListCount - 1
      Outline.Indent(nListIndex) = 1
      Outline.PictureType(nListIndex) = outClosed
      Mdl(nIndex).ListIndex = nListIndex
      MakeReference nListIndex, nIndex, -1

      ' Controls
      If Mdl(nIndex).CtrlCount > 0 Then
         Outline.AddItem "(Controls)" & "     [" & Mdl(nIndex).CtrlCount & "]"
         nListIndex = Outline.ListCount - 1
         Outline.Indent(nListIndex) = 2
         Outline.PictureType(nListIndex) = outClosed
         Mdl(nIndex).CtrlLIndex = nListIndex
         MakeReference nListIndex, nIndex, 0
      End If

      ' Declaration and procedures
      If Mdl(nIndex).ProcCount > 0 Then
         For i = 1 To Mdl(nIndex).ProcCount
            Outline.AddItem Mdl(nIndex).Proc(i).Name
            nListIndex = Outline.ListCount - 1
            Outline.Indent(nListIndex) = 2
            Outline.PictureType(nListIndex) = outClosed
            Mdl(nIndex).Proc(i).ListIndex = nListIndex
            MakeReference nListIndex, nIndex, i
         Next
      End If
   End If

   ShowCounts
   Refresh

End Sub

Private Sub MakeReference(nListIndex As Integer, nFileIndex As Integer, nProcIndex As Integer)
   ReDim Preserve ItemRef(nListIndex)
   ItemRef(nListIndex).FilePoint = nFileIndex
   ItemRef(nListIndex).ProcPoint = nProcIndex      ' -1 = File, 0 = Controls, 1 > = Procedures
End Sub

Private Sub GetProjectDetails()
   On Error Resume Next
   If Outline.ListCount > 0 Then ClearOutline

   If Not FileExist(txtProject) Then
      ButtonsState
      MsgBox "File not found. Please specify other file", vbInformation, "Not found"
      Exit Sub
   End If

   If txtProject = sLoadedFile Then
      ButtonsState
      Exit Sub
   End If

   Me.MousePointer = vbHourglass
   SetStatusText "Analysing " & ExtractFileName(txtProject) & " ..."
   StatusBar.Refresh

   MakeSound WAVE_ANALYSE

   ' Get the extention to obtain type
   Select Case UCase$(ExtractFileExt(txtProject))
   Case "VBP"
      ' Open the VBP file and extract the files...
      Dim nMark As Integer, nHandle As Integer
      Dim sString As String, sFile As String, sPath As String
      sPath = ExtractPath(txtProject)

      nHandle = FreeFile
      Open txtProject For Input Access Read Shared As #nHandle
      
      Do While Not EOF(nHandle)  ' Loop until end of file.
         Line Input #nHandle, sString
      
         If UCase$(Left(sString, 4)) = "FORM" Then
            nMark = InStr(sString, "=")
            If nMark > 0 Then
               sFile = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
               SetOutline sFile, MT_FORM
            End If

         ElseIf UCase$(Left(sString, 6)) = "MODULE" Then
            nMark = InStr(sString, ";")
            If nMark > 0 Then
               sFile = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
               SetOutline sFile, MT_MODULE
            End If

         ElseIf UCase$(Left(sString, 5)) = "CLASS" Then
            nMark = InStr(sString, ";")
            If nMark > 0 Then
               sFile = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
               SetOutline sFile, MT_CLASS
            End If

         ElseIf UCase$(Left(sString, 11)) = "USERCONTROL" Then
            nMark = InStr(sString, "=")
            If nMark > 0 Then
               sFile = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
               SetOutline sFile, MT_CONTROL
            End If

         ElseIf UCase$(Left(sString, 12)) = "PROPERTYPAGE" Then
            nMark = InStr(sString, "=")
            If nMark > 0 Then
               sFile = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
               SetOutline sFile, MT_PROPERTY
            End If

         ElseIf UCase$(Left(sString, 12)) = "USERDOCUMENT" Then
            nMark = InStr(sString, "=")
            If nMark > 0 Then
               sFile = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
               SetOutline sFile, MT_DOCUMENT
            End If

         End If
      Loop
      
      Close #nHandle

   Case "FRM"
      SetOutline txtProject, MT_FORM

   Case "BAS"
      SetOutline txtProject, MT_MODULE

   Case "CLS"
      SetOutline txtProject, MT_CLASS

   Case "CTL"
      SetOutline txtProject, MT_CONTROL

   Case "PAG"
      SetOutline txtProject, MT_PROPERTY

   Case "DOB"
      SetOutline txtProject, MT_DOCUMENT

   End Select

   sLoadedFile = txtProject
'   AddIniString sIniFile, "Options", "LastFile", txtProject
   UpdateRecentFiles txtProject

   ButtonsState

   SetStatusText "Ready"

   MakeSound WAVE_ACCESSED

   Outline.SetFocus

   Me.MousePointer = vbDefault

End Sub

Private Sub Outline_Click()
   ButtonsState
End Sub

Private Sub CountSelected()
   Dim i As Integer, j As Integer, n As Integer

   MdSelected = 0
   PrSelected = 0

   On Error GoTo SelCountError

   For i = 1 To MdCount
      If Mdl(i).Selected = vbUnchecked Then
         Mdl(i).SelCount = 0

      Else  ' File checked or semi-checked
         n = 0

         If Mdl(i).CtrlSelect = vbChecked Then n = n + 1
         
         ' If file is somehow selected, it can have procedures selected too
         For j = 1 To Mdl(i).ProcCount
            If Mdl(i).Proc(j).Selected = vbChecked Then
               PrSelected = PrSelected + 1
               n = n + 1
            End If
         Next

         If n = 0 Then
            Mdl(i).SelCount = 0
            ' Something must be wrong. No procedures/control selected, but file is?
            Mdl(i).Selected = vbUnchecked
         Else
            Mdl(i).SelCount = n
            MdSelected = MdSelected + 1
         End If
      End If
   Next

SelCountError:
End Sub

Private Sub ShowCounts()
   lblFiles = IIf(MdCount = 0, "None", MdCount)
   lblProcedures = IIf(PrCount = 0, "None", PrCount)

   lblSelFiles = IIf(MdSelected = 0, "None", MdSelected)
   lblSelProcs = IIf(PrSelected = 0, "None", PrSelected)
End Sub

Private Sub cmdClear_Click()
   Dim i As Integer, j As Integer, nMax As Integer
   On Error Resume Next
   Me.MousePointer = vbHourglass

   If Outline.ListIndex < 0 Or Outline.Indent(Outline.ListIndex) = 1 Then
      ' File... Clear all entries

      For i = 0 To MdCount
         Mdl(i).Selected = vbUnchecked
         Mdl(i).CtrlSelect = vbUnchecked

         ' If file is somehow selected, it must have procedures selected too
         For j = 1 To Mdl(i).ProcCount
            Mdl(i).Proc(j).Selected = vbUnchecked
         Next
      Next

      ' Remove all "checked" boxes
      For i = 0 To (Outline.ListCount - 1)
         Outline.PictureType(i) = outClosed
      Next

   Else
      ' Procedure... Clear relatives only
      i = ItemRef(Outline.ListIndex).FilePoint     ' Get the file (parent) array pointer
      SetChildrenTick i, vbUnchecked

      Mdl(i).Selected = vbUnchecked
      Outline.PictureType(Mdl(i).ListIndex) = outClosed
   End If

   CountSelected
   ButtonsState
   ShowCounts
   Me.MousePointer = vbDefault
End Sub

Private Sub cmdSelectAll_Click()
   Dim i As Integer, j As Integer
   On Error Resume Next
   Me.MousePointer = vbHourglass

   If Outline.ListIndex < 0 Or Outline.Indent(Outline.ListIndex) = 1 Then
      ' File... Clear all entries

      For i = 0 To MdCount
         Mdl(i).Selected = vbChecked
         Mdl(i).CtrlSelect = vbChecked

         ' If file is somehow selected, it must have procedures selected too
         For j = 1 To Mdl(i).ProcCount
            Mdl(i).Proc(j).Selected = vbChecked
         Next
      Next

      ' Set all boxes to "checked"
      For i = 0 To (Outline.ListCount - 1)
         Outline.PictureType(i) = outOpen
      Next

   Else
      ' Procedure... Set relatives only
      i = ItemRef(Outline.ListIndex).FilePoint     ' Get the file (parent) array pointer
      SetChildrenTick i, vbChecked

      Mdl(i).Selected = vbChecked
      Outline.PictureType(Mdl(i).ListIndex) = outOpen
   End If
   
   CountSelected
   ButtonsState
   ShowCounts
   Me.MousePointer = vbDefault
End Sub

Private Sub Outline_PictureClick(ListIndex As Integer)
   Outline.MousePointer = vbHourglass

   Select Case ItemRef(ListIndex).ProcPoint
   Case Is < 0
      ' It's a file (parent)...
      If Mdl(ItemRef(ListIndex).FilePoint).Selected = vbUnchecked Then
         Mdl(ItemRef(ListIndex).FilePoint).Selected = vbChecked
         SetChildrenTick ItemRef(ListIndex).FilePoint, vbChecked     ' Select all children too
         Outline.PictureType(ListIndex) = outOpen
      Else
         Mdl(ItemRef(ListIndex).FilePoint).Selected = vbUnchecked
         SetChildrenTick ItemRef(ListIndex).FilePoint, vbUnchecked   ' Unselect all children too
         Outline.PictureType(ListIndex) = outClosed
      End If

   Case Is = 0
      ' It's the controls...
      If Mdl(ItemRef(ListIndex).FilePoint).CtrlSelect = vbUnchecked Then
         Mdl(ItemRef(ListIndex).FilePoint).CtrlSelect = vbChecked
         Outline.PictureType(ListIndex) = outOpen
      Else
         Mdl(ItemRef(ListIndex).FilePoint).CtrlSelect = vbUnchecked
         Outline.PictureType(ListIndex) = outClosed
      End If
      SetParentTick ItemRef(ListIndex).FilePoint

   Case Is > 0
      ' It's a procedure or declaration...
      If Mdl(ItemRef(ListIndex).FilePoint).Proc(ItemRef(ListIndex).ProcPoint).Selected = vbUnchecked Then
         Mdl(ItemRef(ListIndex).FilePoint).Proc(ItemRef(ListIndex).ProcPoint).Selected = vbChecked
         Outline.PictureType(ListIndex) = outOpen
      Else
         Mdl(ItemRef(ListIndex).FilePoint).Proc(ItemRef(ListIndex).ProcPoint).Selected = vbUnchecked
         Outline.PictureType(ListIndex) = outClosed
      End If
      SetParentTick ItemRef(ListIndex).FilePoint

   End Select

   CountSelected
   ButtonsState
   ShowCounts
   Outline.MousePointer = vbDefault
End Sub

Private Sub SetParentTick(nFileIndex As Integer)
   Dim nImage As Integer, i As Integer

   ' Get the first child's status
   If Mdl(nFileIndex).CtrlCount > 0 Then
      ' Controls is considered a child.
      nImage = IIf(Mdl(nFileIndex).CtrlSelect = vbChecked, outOpen, outClosed)
   ElseIf Mdl(nFileIndex).ProcCount > 0 Then
      nImage = IIf(Mdl(nFileIndex).Proc(1).Selected = vbChecked, outOpen, outClosed)
   Else
      nImage = IIf(Mdl(nFileIndex).Selected = vbChecked, outOpen, outClosed)
   End If

   If Mdl(nFileIndex).ProcCount > 0 Then
      For i = 1 To Mdl(nFileIndex).ProcCount
         If Mdl(nFileIndex).Proc(i).Selected = vbChecked And nImage = outClosed Then
            nImage = outLeaf     ' Grey it...
            Exit For
         ElseIf Mdl(nFileIndex).Proc(i).Selected = vbUnchecked And nImage = outOpen Then
            nImage = outLeaf     ' Grey it...
            Exit For
         End If
      Next
   End If

   Select Case nImage
   Case outOpen, outLeaf
      Mdl(nFileIndex).Selected = True
   Case outClosed
      Mdl(nFileIndex).Selected = False
   End Select

   Outline.PictureType(Mdl(nFileIndex).ListIndex) = nImage
End Sub

Private Sub SetChildrenTick(nFileIndex As Integer, nSelect As Integer)
   Dim nImage As Integer, i As Integer
   nImage = IIf(nSelect = vbChecked, outOpen, outClosed)

   If Mdl(nFileIndex).CtrlCount > 0 Then
      ' Controls item is considered a child.
      Mdl(nFileIndex).CtrlSelect = nSelect
      Outline.PictureType(Mdl(nFileIndex).CtrlLIndex) = nImage
   End If

   If Mdl(nFileIndex).ProcCount > 0 Then
      For i = 1 To Mdl(nFileIndex).ProcCount
         Mdl(nFileIndex).Proc(i).Selected = nSelect
         Outline.PictureType(Mdl(nFileIndex).Proc(i).ListIndex) = nImage
      Next
   End If
End Sub

Private Sub ButtonsState()
   On Error Resume Next

   Dim nIndex As Integer
   If Outline.ListCount = 0 Then
      nIndex = -1
   Else
      nIndex = Outline.ListIndex
   End If

   ShowCounts

   If Outline.ListCount < 1 Or nIndex < 0 Then
      lblName = "(No item selected)"
      lblType = ""
   Else
      Select Case ItemRef(nIndex).ProcPoint
      Case Is < 0
         ' File...
         Select Case Mdl(ItemRef(nIndex).FilePoint).Type
         Case MT_FORM
            lblName = LongDirFix(Mdl(ItemRef(nIndex).FilePoint).PathFile, 30)
            If Len(Trim(Mdl(ItemRef(nIndex).FilePoint).Name)) = 0 Then
               lblType = "Form"
            Else
               lblType = "Form - " & Mdl(ItemRef(nIndex).FilePoint).Name
            End If
         Case MT_MODULE
            lblName = LongDirFix(Mdl(ItemRef(nIndex).FilePoint).PathFile, 30)
            If Len(Trim(Mdl(ItemRef(nIndex).FilePoint).Name)) = 0 Then
               lblType = "Module"
            Else
               lblType = "Module - " & Mdl(ItemRef(nIndex).FilePoint).Name
            End If
         Case MT_CLASS
            lblName = LongDirFix(Mdl(ItemRef(nIndex).FilePoint).PathFile, 30)
            If Len(Trim(Mdl(ItemRef(nIndex).FilePoint).Name)) = 0 Then
               lblType = "Class"
            Else
               lblType = "Class - " & Mdl(ItemRef(nIndex).FilePoint).Name
            End If
         Case MT_CONTROL
            lblName = LongDirFix(Mdl(ItemRef(nIndex).FilePoint).PathFile, 30)
            If Len(Trim(Mdl(ItemRef(nIndex).FilePoint).Name)) = 0 Then
               lblType = "User Control"
            Else
               lblType = "User Control - " & Mdl(ItemRef(nIndex).FilePoint).Name
            End If
         Case MT_PROPERTY
            lblName = LongDirFix(Mdl(ItemRef(nIndex).FilePoint).PathFile, 30)
            If Len(Trim(Mdl(ItemRef(nIndex).FilePoint).Name)) = 0 Then
               lblType = "Property Page"
            Else
               lblType = "Property Page - " & Mdl(ItemRef(nIndex).FilePoint).Name
            End If
         Case MT_DOCUMENT
            lblName = LongDirFix(Mdl(ItemRef(nIndex).FilePoint).PathFile, 30)
            If Len(Trim(Mdl(ItemRef(nIndex).FilePoint).Name)) = 0 Then
               lblType = "User Document"
            Else
               lblType = "User Document - " & Mdl(ItemRef(nIndex).FilePoint).Name
            End If
         Case Else
            ' Dunno
            lblName = ""
            lblType = ""
         End Select

      Case Is = 0
         ' Controls
         If Mdl(ItemRef(nIndex).FilePoint).CtrlCount > 0 Then
            lblName = LongDirFix(Mdl(ItemRef(nIndex).FilePoint).PathFile, 30)
            lblType = "Form Controls"
         Else
            lblName = ""
            lblType = ""
         End If
      
      Case Is > 0
         ' Procedure
         Select Case Mdl(ItemRef(nIndex).FilePoint).Proc(ItemRef(nIndex).ProcPoint).Type
         Case PT_DECLARE
            lblName = Mdl(ItemRef(nIndex).FilePoint).Proc(ItemRef(nIndex).ProcPoint).Name
            lblType = "Declarations"
         Case PT_PROPERTY
            lblName = Mdl(ItemRef(nIndex).FilePoint).Proc(ItemRef(nIndex).ProcPoint).Name
            lblType = "Property"
         Case PT_SUB
            lblName = Mdl(ItemRef(nIndex).FilePoint).Proc(ItemRef(nIndex).ProcPoint).Name
            lblType = "Sub"
         Case PT_FUNCTION
            lblName = Mdl(ItemRef(nIndex).FilePoint).Proc(ItemRef(nIndex).ProcPoint).Name
            lblType = "Function"
         Case Else
            ' Dunno
            lblName = ""
            lblType = ""
         End Select

      End Select
   End If

   If Outline.ListCount > 0 Then

      SetEnabled cmdView, (nIndex > -1)

      If nIndex < 0 Then
         SetEnabled cmdSelectAll, (MdCount <> MdSelected)
         SetEnabled cmdClear, (MdSelected > 0)
      ElseIf Outline.Indent(nIndex) = 1 Then
         SetEnabled cmdSelectAll, (MdCount <> MdSelected)
         SetEnabled cmdClear, (MdSelected > 0)
      Else
         SetEnabled cmdSelectAll, (Mdl(ItemRef(nIndex).FilePoint).ChildCount <> Mdl(ItemRef(nIndex).FilePoint).SelCount)
         SetEnabled cmdClear, (Mdl(ItemRef(nIndex).FilePoint).SelCount > 0)
      End If

      If MdSelected > 0 Then
         SetEnabled cmdPrint, True
      Else
         SetEnabled cmdPrint, (chkProject = vbChecked And UCase$(ExtractFileExt(txtProject)) = "VBP")
      End If

   Else
      SetEnabled cmdView, False
      SetEnabled cmdClear, False
      SetEnabled cmdSelectAll, False
      SetEnabled cmdPrint, (chkProject = vbChecked And UCase$(ExtractFileExt(txtProject)) = "VBP")
   End If

   SetEnabled cmdPrintSetup, (Not optOutput(2))
   SetEnabled cmdHelp, FileExist(sHelpFile)

End Sub

' Some comments in the footer
' I don't know why, but it's here - What to do with it?
' It actually belongs to the procedure above, but how to connect to it without producing a line
' or page break. - I do need something to test this program with, why not this.
