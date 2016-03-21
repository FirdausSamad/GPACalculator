VERSION 5.00
Begin VB.Form form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15825
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   15825
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15495
      Begin VB.CommandButton cmdExitTotal 
         Caption         =   "Exit"
         Enabled         =   0   'False
         Height          =   495
         Left            =   12000
         TabIndex        =   80
         Top             =   8280
         Width           =   1575
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Enabled         =   0   'False
         Height          =   495
         Left            =   9840
         TabIndex        =   79
         Top             =   8280
         Width           =   1695
      End
      Begin VB.CommandButton cmdCalculate 
         Caption         =   "Calculate"
         Enabled         =   0   'False
         Height          =   495
         Left            =   7680
         TabIndex        =   78
         Top             =   8280
         Width           =   1695
      End
      Begin VB.Frame frmCourseRegister 
         Caption         =   "Course Register"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   9480
         TabIndex        =   52
         Top             =   3000
         Width           =   5775
         Begin VB.TextBox txtCourseName 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   2040
            TabIndex        =   58
            Top             =   360
            Width           =   3375
         End
         Begin VB.ComboBox cmbGrade 
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
            ItemData        =   "Form1.frx":0000
            Left            =   2040
            List            =   "Form1.frx":002E
            TabIndex        =   57
            TabStop         =   0   'False
            Text            =   "(Grade)"
            Top             =   1320
            Width           =   1575
         End
         Begin VB.ComboBox cmbCreditHour 
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
            ItemData        =   "Form1.frx":0069
            Left            =   2040
            List            =   "Form1.frx":007C
            TabIndex        =   56
            Text            =   "(Credit Hour)"
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Form"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2400
            MaskColor       =   &H00808080&
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   2640
            Width           =   1695
         End
         Begin VB.CommandButton cmdAddCourse 
            Caption         =   "Add Course"
            Enabled         =   0   'False
            Height          =   495
            Left            =   480
            MaskColor       =   &H00808080&
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   2640
            Width           =   1695
         End
         Begin VB.TextBox txtCourseCode 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   53
            Top             =   840
            Width           =   3375
         End
         Begin VB.Label lblCourseName 
            Caption         =   "Course Name:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   62
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblCourseCode 
            Caption         =   "Course Code :"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   61
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblGrade 
            Caption         =   "Grade             :"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   60
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label lblCreditHour 
            Caption         =   "Credit Hour    :"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   59
            Top             =   1800
            Width           =   1575
         End
      End
      Begin VB.Frame frmList 
         Caption         =   "List of Subject"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   480
         TabIndex        =   17
         Top             =   6600
         Width           =   6855
         Begin VB.CommandButton cmdClearSubject 
            Caption         =   "Clear"
            Height          =   495
            Left            =   5520
            TabIndex        =   83
            Top             =   2040
            Width           =   1215
         End
         Begin VB.CheckBox chkSubject1 
            Height          =   255
            Left            =   120
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   840
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkSubject2 
            Height          =   255
            Left            =   120
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   1200
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkSubject3 
            Height          =   255
            Left            =   120
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1560
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkSubject4 
            Height          =   255
            Left            =   120
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1920
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkSubject5 
            Height          =   255
            Left            =   120
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   2280
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkSubject6 
            Height          =   255
            Left            =   120
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   2640
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkSubject7 
            Height          =   255
            Left            =   120
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   3000
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton cmdConfirmList 
            Caption         =   "Confirm"
            Enabled         =   0   'False
            Height          =   495
            Left            =   5520
            TabIndex        =   20
            Top             =   2640
            Width           =   1215
         End
         Begin VB.CommandButton cmdCheckAll 
            Caption         =   "Check All"
            Enabled         =   0   'False
            Height          =   495
            Left            =   5520
            TabIndex        =   19
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdUnchecked 
            Caption         =   "Unchecked"
            Enabled         =   0   'False
            Height          =   495
            Left            =   5520
            TabIndex        =   18
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblCheckAll 
            Caption         =   "*Please Check All Subject You Want To Calculate "
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   240
            TabIndex        =   82
            Top             =   3360
            Width           =   4455
         End
         Begin VB.Label lblCredit5 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4440
            TabIndex        =   70
            Top             =   2280
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblCredit6 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4440
            TabIndex        =   69
            Top             =   2640
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblCredit7 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4440
            TabIndex        =   68
            Top             =   3000
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblCredit2 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4440
            TabIndex        =   67
            Top             =   1200
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblCredit3 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4440
            TabIndex        =   66
            Top             =   1560
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblCredit4 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4440
            TabIndex        =   65
            Top             =   1920
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblCredit1 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4440
            TabIndex        =   64
            Top             =   840
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblCreditList 
            Caption         =   "Credit Hours"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4080
            TabIndex        =   63
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblCourseCodeList 
            Caption         =   "Course Code"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1440
            TabIndex        =   51
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblGradeList 
            Caption         =   "Grade"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3000
            TabIndex        =   50
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblNo 
            Caption         =   "No."
            Enabled         =   0   'False
            Height          =   255
            Left            =   600
            TabIndex        =   49
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblNo1 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   48
            Top             =   840
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblSubjectCode1 
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   47
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblGrade1 
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   46
            Top             =   840
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblNo2 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   45
            Top             =   1200
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblSubjectCode2 
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   44
            Top             =   1200
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblGrade5 
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   43
            Top             =   2280
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblNo3 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   42
            Top             =   1560
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblSubjectCode3 
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   41
            Top             =   1560
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblGrade3 
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   40
            Top             =   1560
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblNo4 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   39
            Top             =   1920
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblSubjectCode4 
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   38
            Top             =   1920
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblGrade4 
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   37
            Top             =   1920
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblNo5 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   36
            Top             =   2280
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblSubjectCode5 
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   35
            Top             =   2280
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblNo6 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   34
            Top             =   2640
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblSubjectCode6 
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   33
            Top             =   2640
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblGrade6 
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   32
            Top             =   2640
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblNo7 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   31
            Top             =   3000
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblSubjectCode7 
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   30
            Top             =   3000
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblGrade7 
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   29
            Top             =   3000
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblGrade2 
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   28
            Top             =   1200
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Frame frmStudent 
         Caption         =   "Student Detail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   240
         TabIndex        =   1
         Top             =   3000
         Width           =   9135
         Begin VB.TextBox txtSemester 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1320
            Width           =   5175
         End
         Begin VB.TextBox txtStudentNumber 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   840
            Width           =   5175
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   8
            Top             =   360
            Width           =   5175
         End
         Begin VB.TextBox txtCampus 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1800
            Width           =   5175
         End
         Begin VB.TextBox txtFaculty 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   2280
            Width           =   5175
         End
         Begin VB.TextBox txtProgramme 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   2760
            Width           =   5175
         End
         Begin VB.CommandButton cmdRegister 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Register"
            Height          =   495
            Left            =   7680
            MaskColor       =   &H00808080&
            TabIndex        =   4
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdClearDetail 
            Caption         =   "Clear Form"
            Height          =   495
            Left            =   7680
            TabIndex        =   3
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdExitStudent 
            Caption         =   "Exit"
            Height          =   495
            Left            =   7680
            TabIndex        =   2
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label lblSemester 
            Caption         =   "Semester              :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label lblStudentNumber 
            Caption         =   "Student Number   :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label lblStudentName 
            Caption         =   "Student Name      :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label lblCampus 
            Caption         =   "Campus                 :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   13
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label lblFaculty 
            Caption         =   "Faculty                  :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label lblProgramme 
            Caption         =   "Programme          : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   2760
            Width           =   2175
         End
      End
      Begin VB.Label lblTitle 
         Caption         =   "       UITM MELAKA GPA CALCULATOR"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         TabIndex        =   81
         Top             =   2400
         Width           =   5655
      End
      Begin VB.Label lblGradePointAverage1 
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   12240
         TabIndex        =   77
         Top             =   7560
         Width           =   1815
      End
      Begin VB.Label lblGradePointAverage 
         Caption         =   "GRADE POINT AVERAGE"
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
         Height          =   375
         Left            =   12240
         TabIndex        =   76
         Top             =   7200
         Width           =   2055
      End
      Begin VB.Label lblCreditTotal1 
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   9960
         TabIndex        =   75
         Top             =   7560
         Width           =   1695
      End
      Begin VB.Label lblCreditTotal 
         Caption         =   "CREDIT TOTAL"
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
         Height          =   375
         Left            =   9960
         TabIndex        =   74
         Top             =   7200
         Width           =   1695
      End
      Begin VB.Label lblGradePointTotal1 
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   7680
         TabIndex        =   73
         Top             =   7560
         Width           =   1695
      End
      Begin VB.Label lblGradePointTotal 
         Caption         =   "GRADE POINT TOTAL"
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
         Height          =   375
         Left            =   7680
         TabIndex        =   72
         Top             =   7200
         Width           =   2175
      End
      Begin VB.Label lblCurrentSemester 
         Caption         =   "Current Semester  :"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   71
         Top             =   6720
         Width           =   2175
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1935
         Left            =   120
         Picture         =   "Form1.frx":009B
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3135
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1935
         Left            =   3360
         Picture         =   "Form1.frx":68173
         Stretch         =   -1  'True
         Top             =   240
         Width           =   11895
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim no As Integer
Dim creditHour As Integer
Dim index As Integer
Dim no1 As Integer

Private Sub cmdAddCourse_Click()
'To Add Course
Dim CourseCode As String
Dim Grade As String
Dim credit As Double

'Calculate Credit Hour
creditHour = creditHour + Val(cmbCreditHour.Text)
    no = no + 1
    'To get Course Code ,Grade and Credit from text box
    CourseCode = txtCourseCode.Text
    Grade = cmbGrade.Text
    credit = Val(cmbCreditHour.Text)
                        
                        lblCurrentSemester.Enabled = False
                        lblGradePointTotal.Enabled = False
                        lblGradePointTotal1.Enabled = False
                        lblCreditTotal.Enabled = False
                        lblCreditTotal1.Enabled = False
                        lblGradePointAverage.Enabled = False
                        lblGradePointAverage1.Enabled = False
                        cmdCalculate.Enabled = False
                        cmdPrint.Enabled = False
                        cmdExitTotal.Enabled = False
                        'To Limit the Credit Hour less than 21
                        If creditHour <= 21 Then
                                        'To add Subject. Subject will be add by number
                                        If no = 1 Then
                                        lblNo1.Caption = no
                                        lblSubjectCode1.Caption = CourseCode
                                        lblGrade1.Caption = Grade
                                        lblCredit1.Caption = credit
       
        
                                        chkSubject1.Visible = True
                                        lblNo1.Visible = True
                                        lblSubjectCode1.Visible = True
                                        lblGrade1.Visible = True
                                        lblCredit1.Visible = True
                                            'To add Subject. Subject will be add by number
                                         ElseIf no = 2 Then
                                         lblNo2.Caption = no
                                         lblSubjectCode2.Caption = CourseCode
                                         lblGrade2.Caption = Grade
                                         lblCredit2.Caption = credit
                                          
                                         chkSubject2.Visible = True
                                         lblNo2.Visible = True
                                         lblSubjectCode2.Visible = True
                                         lblGrade2.Visible = True
                                         lblCredit2.Visible = True
      
                                        'To add Subject. Subject will be add by number
                                         ElseIf no = 3 Then
                                         lblNo3.Caption = no
                                         lblSubjectCode3.Caption = CourseCode
                                         lblGrade3.Caption = Grade
                                         lblCredit3.Caption = credit
        
                                         chkSubject3.Visible = True
                                         lblNo3.Visible = True
                                         lblSubjectCode3.Visible = True
                                         lblGrade3.Visible = True
                                         lblCredit3.Visible = True
                                        'To add Subject. Subject will be add by number
                                      ElseIf no = 4 Then
                                         lblNo4.Caption = no
                                         lblSubjectCode4.Caption = CourseCode
                                         lblGrade4.Caption = Grade
                                         lblCredit4.Caption = credit
        
                                         chkSubject4.Visible = True
                                         lblNo4.Visible = True
                                         lblSubjectCode4.Visible = True
                                         lblGrade4.Visible = True
                                         lblCredit4.Visible = True
                                            'To add Subject. Subject will be add by number
                                      ElseIf no = 5 Then
                                        lblNo5.Caption = no
                                        lblSubjectCode5.Caption = CourseCode
                                        lblCredit5.Caption = credit
                                        lblGrade5.Caption = Grade
                                        chkSubject5.Visible = True
                                        lblNo5.Visible = True
                                        lblSubjectCode5.Visible = True
                                        lblGrade5.Visible = True
                                        lblCredit5.Visible = True
                                        'To add Subject. Subject will be add by number
                                     ElseIf no = 6 Then
                                        lblNo6.Caption = no
                                        lblSubjectCode6.Caption = CourseCode
                                        lblGrade6.Caption = Grade
                                        lblCredit6.Caption = credit
        
                                        chkSubject6.Visible = True
                                        lblNo6.Visible = True
                                        lblSubjectCode6.Visible = True
                                        lblGrade6.Visible = True
                                        lblCredit6.Visible = True
                                        'To add Subject. Subject will be add by number
                                     ElseIf no = 7 Then
                                        lblNo7.Caption = no
                                        lblSubjectCode7.Caption = CourseCode
                                        lblGrade7.Caption = Grade
                                        lblCredit7.Caption = credit
        
                                        chkSubject7.Visible = True
                                        lblNo7.Visible = True
                                        lblSubjectCode7.Visible = True
                                        lblGrade7.Visible = True
                                        lblCredit7.Visible = True
                                     Else
                                         MsgBox "You can only add 7 Course.", vbInformation, "Course Added limit"
                                End If
                            Else
                                  MsgBox "Credit hour limit.", vbInformation, "Credit Hour Limit"
                        End If
End Sub

Private Sub cmdCalculate_Click()
'lblCreditTotal1 = lblCredit1
Dim gradePoint1 As Double
Dim gradePoint2 As Double
Dim gradePoint3 As Double
Dim gradePoint4 As Double
Dim gradePoint5 As Double
Dim gradePoint6 As Double
Dim gradePoint7 As Double

Dim dblCreditHour1 As Double
Dim dblCreditHour2 As Double
Dim dblCreditHour3 As Double
Dim dblCreditHour4 As Double
Dim dblCreditHour5 As Double
Dim dblCreditHour6 As Double
Dim dblCreditHour7 As Double

'To find the Checked in List Of Subject
If chkSubject1.Value = Checked Then

'To set The pointer for each grade
        If lblGrade1 = "A+" Then
            gradePoint1 = 4
        ElseIf lblGrade1 = "A" Then
            gradePoint1 = 4
        ElseIf lblGrade1 = "A-" Then
            gradePoint1 = 3.67
        ElseIf lblGrade1 = "B+" Then
            gradePoint1 = 3.33
        ElseIf lblGrade1 = "B" Then
            gradePoint1 = 3
        ElseIf lblGrade1 = "B-" Then
            gradePoint1 = 2.67
        ElseIf lblGrade1 = "C+" Then
            gradePoint1 = 2.33
        ElseIf lblGrade1 = "C" Then
            gradePoint1 = 2
        ElseIf lblGrade1 = "C-" Then
            gradePoint1 = 1.67
        ElseIf lblGrade1 = "D+" Then
            gradePoint1 = 1.33
        ElseIf lblGrade1 = "D" Then
            gradePoint1 = 1
        ElseIf lblGrade1 = "E" Then
            gradePoint1 = 0.67
        ElseIf lblGrade1 = "F" Then
            gradePoint1 = 0
        End If
        'To set the credit hour
        If lblCredit1 = "1" Then
            dblCreditHour1 = 1
        ElseIf lblCredit1 = "2" Then
            dblCreditHour1 = 2
        ElseIf lblCredit1 = "3" Then
            dblCreditHour1 = 3
        ElseIf lblCredit1 = "4" Then
            dblCreditHour1 = 4
        End If
End If
'To find the Checked in List Of Subject
If chkSubject2.Value = Checked Then
        If lblGrade2 = "A+" Then
            gradePoint2 = 4
        ElseIf lblGrade2 = "A" Then
            gradePoint2 = 4
        ElseIf lblGrade2 = "A-" Then
            gradePoint2 = 3.67
        ElseIf lblGrade2 = "B+" Then
            gradePoint2 = 3.33
        ElseIf lblGrade2 = "B" Then
            gradePoint2 = 3
        ElseIf lblGrade2 = "B-" Then
            gradePoint2 = 2.67
        ElseIf lblGrade2 = "C+" Then
            gradePoint2 = 2.33
        ElseIf lblGrade2 = "C" Then
            gradePoint2 = 2
        ElseIf lblGrade2 = "C-" Then
            gradePoint2 = 1.67
        ElseIf lblGrade2 = "D+" Then
            gradePoint2 = 1.33
        ElseIf lblGrade2 = "D" Then
            gradePoint2 = 1
        ElseIf lblGrade2 = "E" Then
            gradePoint2 = 0.67
        ElseIf lblGrade2 = "F" Then
            gradePoint2 = 0
        End If
        
         'To set the credit hour
        If lblCredit2 = "1" Then
            dblCreditHour2 = 1
        ElseIf lblCredit2 = "2" Then
            dblCreditHour2 = 2
        ElseIf lblCredit2 = "3" Then
            dblCreditHour2 = 3
        ElseIf lblCredit2 = "4" Then
            dblCreditHour2 = 4
        End If
        
End If
'To find the Checked in List Of Subject

If chkSubject3.Value = Checked Then
        If lblGrade3 = "A+" Then
            gradePoint3 = 4
        ElseIf lblGrade3 = "A" Then
            gradePoint3 = 4
        ElseIf lblGrade3 = "A-" Then
            gradePoint3 = 3.67
        ElseIf lblGrade3 = "B+" Then
            gradePoint3 = 3.33
        ElseIf lblGrade3 = "B" Then
            gradePoint3 = 3
        ElseIf lblGrade3 = "B-" Then
            gradePoint3 = 2.67
        ElseIf lblGrade3 = "C+" Then
            gradePoint3 = 2.33
        ElseIf lblGrade3 = "C" Then
            gradePoint3 = 2
        ElseIf lblGrade3 = "C-" Then
            gradePoint3 = 1.67
        ElseIf lblGrade3 = "D+" Then
            gradePoint3 = 1.33
        ElseIf lblGrade3 = "D" Then
            gradePoint3 = 1
        ElseIf lblGrade3 = "E" Then
            gradePoint3 = 0.67
        ElseIf lblGrade3 = "F" Then
            gradePoint3 = 0
        End If
        
         'To set the credit hour
        If lblCredit3 = "3" Then
            dblCreditHour3 = 1
        ElseIf lblCredit3 = "2" Then
            dblCreditHour3 = 2
        ElseIf lblCredit3 = "3" Then
            dblCreditHour3 = 3
        ElseIf lblCredit3 = "4" Then
            dblCreditHour3 = 4
        End If
        
End If
'To find the Checked in List Of Subject
If chkSubject4.Value = Checked Then
        If lblGrade4 = "A+" Then
            gradePoint4 = 4
        ElseIf lblGrade4 = "A" Then
            gradePoint4 = 4
        ElseIf lblGrade4 = "A-" Then
            gradePoint4 = 3.67
        ElseIf lblGrade4 = "B+" Then
            gradePoint4 = 3.33
        ElseIf lblGrade4 = "B" Then
            gradePoint4 = 3
        ElseIf lblGrade4 = "B-" Then
            gradePoint4 = 2.67
        ElseIf lblGrade4 = "C+" Then
            gradePoint4 = 2.33
        ElseIf lblGrade4 = "C" Then
            gradePoint4 = 2
        ElseIf lblGrade4 = "C-" Then
            gradePoint4 = 1.67
        ElseIf lblGrade4 = "D+" Then
            gradePoint4 = 1.33
        ElseIf lblGrade4 = "D" Then
            gradePoint4 = 1
        ElseIf lblGrade4 = "E" Then
            gradePoint4 = 0.67
        ElseIf lblGrade4 = "F" Then
            gradePoint4 = 0
        End If
        
         'To set the credit hour
        If lblCredit4 = "1" Then
            dblCreditHour4 = 1
        ElseIf lblCredit4 = "2" Then
            dblCreditHour4 = 2
        ElseIf lblCredit4 = "3" Then
            dblCreditHour4 = 3
        ElseIf lblCredit4 = "4" Then
            dblCreditHour4 = 4
        End If
        
End If
'To find the Checked in List Of Subject
If chkSubject5.Value = Checked Then
        If lblGrade5 = "A+" Then
            gradePoint5 = 4
        ElseIf lblGrade5 = "A" Then
            gradePoint5 = 4
        ElseIf lblGrade5 = "A-" Then
            gradePoint5 = 3.67
        ElseIf lblGrade5 = "B+" Then
            gradePoint5 = 3.33
        ElseIf lblGrade5 = "B" Then
            gradePoint5 = 3
        ElseIf lblGrade5 = "B-" Then
            gradePoint5 = 2.67
        ElseIf lblGrade5 = "C+" Then
            gradePoint5 = 2.33
        ElseIf lblGrade5 = "C" Then
            gradePoint5 = 2
        ElseIf lblGrade5 = "C-" Then
            gradePoint5 = 1.67
        ElseIf lblGrade5 = "D+" Then
            gradePoint5 = 1.33
        ElseIf lblGrade5 = "D" Then
            gradePoint5 = 1
        ElseIf lblGrade5 = "E" Then
            gradePoint5 = 0.67
        ElseIf lblGrade5 = "F" Then
            gradePoint5 = 0
        End If
        
         'To set the credit hour
        If lblCredit5 = "1" Then
            dblCreditHour5 = 1
        ElseIf lblCredit5 = "2" Then
            dblCreditHour5 = 2
        ElseIf lblCredit5 = "3" Then
            dblCreditHour5 = 3
        ElseIf lblCredit5 = "4" Then
            dblCreditHour5 = 4
        End If
        
End If
'To find the Checked in List Of Subject
If chkSubject6.Value = Checked Then
        If lblGrade6 = "A+" Then
            gradePoint6 = 4
        ElseIf lblGrade6 = "A" Then
            gradePoint6 = 4
        ElseIf lblGrade6 = "A-" Then
            gradePoint6 = 3.67
        ElseIf lblGrade6 = "B+" Then
            gradePoint6 = 3.33
        ElseIf lblGrade6 = "B" Then
            gradePoint6 = 3
        ElseIf lblGrade6 = "B-" Then
            gradePoint6 = 2.67
        ElseIf lblGrade6 = "C+" Then
            gradePoint6 = 2.33
        ElseIf lblGrade6 = "C" Then
            gradePoint6 = 2
        ElseIf lblGrade6 = "C-" Then
            gradePoint6 = 1.67
        ElseIf lblGrade6 = "D+" Then
            gradePoint6 = 1.33
        ElseIf lblGrade6 = "D" Then
            gradePoint6 = 1
        ElseIf lblGrade6 = "E" Then
            gradePoint6 = 0.67
        ElseIf lblGrade6 = "F" Then
            gradePoint6 = 0
        End If
        
         'To set the credit hour
        If lblCredit6 = "1" Then
            dblCreditHour6 = 1
        ElseIf lblCredit6 = "2" Then
            dblCreditHour6 = 2
        ElseIf lblCredit6 = "3" Then
            dblCreditHour6 = 3
        ElseIf lblCredit6 = "4" Then
            dblCreditHour6 = 4
        End If
End If
'To find the Checked in List Of Subject
If chkSubject7.Value = Checked Then
        If lblGrade7 = "A+" Then
            gradePoint7 = 4
        ElseIf lblGrade7 = "A" Then
            gradePoint7 = 4
        ElseIf lblGrade7 = "A-" Then
            gradePoint7 = 3.67
        ElseIf lblGrade7 = "B+" Then
            gradePoint7 = 3.33
        ElseIf lblGrade7 = "B" Then
            gradePoint = 3
        ElseIf lblGrade7 = "B-" Then
            gradePoint7 = 2.67
        ElseIf lblGrade7 = "C+" Then
            gradePoint7 = 2.33
        ElseIf lblGrade7 = "C" Then
            gradePoint7 = 2
        ElseIf lblGrade7 = "C-" Then
            gradePoint7 = 1.67
        ElseIf lblGrade7 = "D+" Then
            gradePoint7 = 1.33
        ElseIf lblGrade7 = "D" Then
            gradePoint7 = 1
        ElseIf lblGrade7 = "E" Then
            gradePoint7 = 0.67
        ElseIf lblGrade7 = "F" Then
            gradePoint7 = 0
        End If
        
         'To set the credit hour
        If lblCredit7 = "1" Then
            dblCreditHour7 = 1
        ElseIf lblCredit7 = "2" Then
            dblCreditHour7 = 2
        ElseIf lblCredit7 = "3" Then
            dblCreditHour7 = 3
        ElseIf lblCredit7 = "4" Then
            dblCreditHour7 = 4
        End If
End If

Dim dblGradePointTotal As Double
Dim dblTotalCredit As Double
Dim dblGradePointAverage As Double

'To calculate Grade Point Total
dblGradePointTotal = (gradePoint1 * dblCreditHour1) + (gradePoint2 * dblCreditHour2) + (gradePoint3 * dblCreditHour3) + (gradePoint4 * dblCreditHour4) + (gradePoint5 * dblCreditHour5) + (gradePoint6 * dblCreditHour6) + (gradePoint7 * dblCreditHour7)
lblGradePointTotal1 = dblGradePointTotal
'To calculate Total Credit
dblTotalCredit = dblCreditHour1 + dblCreditHour2 + dblCreditHour3 + dblCreditHour4 + dblCreditHour5 + dblCreditHour6 + dblCreditHour7
lblCreditTotal1 = dblTotalCredit
'To Calculate GPA
If Not dblGradePointTotal = 0 And Not dblTotalCredit = 0 Then
dblGradePointAverage = (dblGradePointTotal / dblTotalCredit)
lblGradePointAverage1 = Round(dblGradePointAverage, 2)
Else
MsgBox "Please register your course.", vbInformation, "No Course register"


End If

End Sub

Private Sub cmdCheckAll_Click()

'To make check all button
chkSubject1.Value = Checked
chkSubject2.Value = Checked
chkSubject3.Value = Checked
chkSubject4.Value = Checked
chkSubject5.Value = Checked
chkSubject6.Value = Checked
chkSubject7.Value = Checked

End Sub

Private Sub cmdClear_Click()
'To clear in add course button
txtCourseName.Text = ""
txtCourseCode.Text = ""
cmbGrade.Text = "(Grade)"
cmbCreditHour.Text = "(Credit Hour)"

txtCourseName.SetFocus

End Sub

Private Sub cmdClearDetail_Click()

                       'To disable the form
                        lblCreditList.Enabled = False
                        frmList.Enabled = False
                        lblCourseName.Enabled = False
                        txtCourseName.Enabled = False
                        lblCourseCode.Enabled = False
                        txtCourseCode.Enabled = False
                        lblGrade.Enabled = False
                        cmbGrade.Enabled = False
                        lblCreditHour.Enabled = False
                        cmbCreditHour.Enabled = False
                        cmdAddCourse.Enabled = False
                        cmdClear.Enabled = False
                        frmList.Enabled = False
                        lblNo.Enabled = False
                        lblCourseCodeList.Enabled = False
                        lblGradeList.Enabled = False
                        cmdConfirmList.Enabled = False
                        
                        cmdCheckAll.Enabled = False
                        cmdUnchecked.Enabled = False
                        
                        lblCurrentSemester.Enabled = False
                        lblGradePointTotal.Enabled = False
                        lblGradePointTotal1.Enabled = False
                        lblCreditTotal.Enabled = False
                        lblCreditTotal1.Enabled = False
                        lblGradePointAverage.Enabled = False
                        lblGradePointAverage1.Enabled = False
                        cmdCalculate.Enabled = False
                        cmdPrint.Enabled = False
                        cmdExitTotal.Enabled = False
                        
                        txtName.Text = ""
                        txtStudentNumber.Text = ""
                        txtSemester.Text = ""
                        txtCampus.Text = ""
                        txtFaculty.Text = ""
                        txtProgramme.Text = ""
                        
                        
                        lblNo1.Caption = ""
                        lblSubjectCode1.Caption = ""
                        lblGrade1.Caption = ""
                        lblCredit1.Caption = ""
                        chkSubject1.Visible = False
                        
                        lblNo2.Caption = ""
                        lblSubjectCode2.Caption = ""
                        lblGrade2.Caption = ""
                        lblCredit2.Caption = ""
                        chkSubject2.Visible = False
                        
                        lblNo3.Caption = ""
                        lblSubjectCode3.Caption = ""
                        lblGrade3.Caption = ""
                        lblCredit3.Caption = ""
                        chkSubject3.Visible = False
                        
                        lblNo4.Caption = ""
                        lblSubjectCode4.Caption = ""
                        lblGrade4.Caption = ""
                        lblCredit4.Caption = ""
                        chkSubject4.Visible = False
                        
                        lblNo5.Caption = ""
                        lblSubjectCode5.Caption = ""
                        lblGrade5.Caption = ""
                        lblCredit5.Caption = ""
                        chkSubject5.Visible = False
                        
                        lblNo6.Caption = ""
                        lblSubjectCode6.Caption = ""
                        lblGrade6.Caption = ""
                        lblCredit6.Caption = ""
                        chkSubject6.Visible = False
                        
                        lblNo7.Caption = ""
                        lblSubjectCode7.Caption = ""
                        lblGrade7.Caption = ""
                        lblCredit7.Caption = ""
                        chkSubject7.Visible = False
                        
                        lblGradePointTotal1.Caption = ""
                        lblCreditTotal1.Caption = ""
                        lblGradePointAverage1.Caption = ""
                        
                        txtName.SetFocus
                        
                        
End Sub





Private Sub cmdClearSubject_Click()
no = 0
creditHour = 0
                                        chkSubject1.Visible = False
                                        lblNo1.Visible = False
                                        lblSubjectCode1.Visible = False
                                        lblGrade1.Visible = False
                                        lblCredit1.Visible = False
                                     
                                        chkSubject2.Visible = False
                                         lblNo2.Visible = False
                                         lblSubjectCode2.Visible = False
                                         lblGrade2.Visible = False
                                         lblCredit2.Visible = False
                                      
                                          chkSubject3.Visible = False
                                         lblNo3.Visible = False
                                         lblSubjectCode3.Visible = False
                                         lblGrade3.Visible = False
                                         lblCredit3.Visible = False
                                       
                                         chkSubject4.Visible = False
                                         lblNo4.Visible = False
                                         lblSubjectCode4.Visible = False
                                         lblGrade4.Visible = False
                                         lblCredit4.Visible = False
                                      
                                         chkSubject5.Visible = False
                                         lblNo5.Visible = False
                                         lblSubjectCode5.Visible = False
                                         lblGrade5.Visible = False
                                         lblCredit5.Visible = False
                                        
                                         chkSubject6.Visible = False
                                         lblNo6.Visible = False
                                         lblSubjectCode6.Visible = False
                                         lblGrade6.Visible = False
                                         lblCredit6.Visible = False
                                         
                                         chkSubject7.Visible = False
                                         lblNo7.Visible = False
                                         lblSubjectCode7.Visible = False
                                         lblGrade7.Visible = False
                                         lblCredit7.Visible = False
                                  
                                        
End Sub

Private Sub cmdConfirmList_Click()




'Conform button
MsgBox "Do You Confirm This List.", vbOKCancel, "Confirmation"
                     
                        lblCurrentSemester.Enabled = True
                        lblGradePointTotal.Enabled = True
                        lblGradePointTotal1.Enabled = True
                        lblCreditTotal.Enabled = True
                        lblCreditTotal1.Enabled = True
                        lblGradePointAverage.Enabled = True
                        lblGradePointAverage1.Enabled = True
                        cmdCalculate.Enabled = True
                        cmdPrint.Enabled = True
                        cmdExitTotal.Enabled = True
End Sub

Private Sub cmdExitStudent_Click()
End
End Sub


Private Sub Command2_Click()

End Sub

Private Sub cmdExitTotal_Click()
End
End Sub

Private Sub cmdPrint_Click()
PrintForm
End Sub

Private Sub cmdRegister_Click()

'To display the error if the form does not fill
If txtName.Text = "" Then
MsgBox "Please Enter Your Name.", vbInformation, "Error"
ElseIf txtStudentNumber.Text = "" Then
MsgBox "Please Enter Your Student Number.", vbInformation, "Error"
ElseIf txtSemester.Text = "" Then
MsgBox "Please Enter Your Semester.", vbInformation, "Error"
ElseIf txtCampus.Text = "" Then
MsgBox "Please Enter Your Campus.", vbInformation, "Error"
ElseIf txtFaculty.Text = "" Then
MsgBox "Please Enter Your Faculty.", vbInformation, "Error"
ElseIf txtProgramme.Text = "" Then
MsgBox "Please Enter Your Programme.", vbInformation, "Error"
End If
                
                    
                
              
'To Enable the form if all the Text entered Success
If Not txtName.Text = "" Then
    If Not txtStudentNumber.Text = "" Then
        If Not txtSemester.Text = "" Then
            If Not txtCampus.Text = "" Then
                If Not txtFaculty.Text = "" Then
                    If Not txtProgramme.Text = "" Then
                        frmCourseRegister.Enabled = True
                        lblCourseName.Enabled = True
                        txtCourseName.Enabled = True
                        lblCourseCode.Enabled = True
                        txtCourseCode.Enabled = True
                        lblGrade.Enabled = True
                        cmbGrade.Enabled = True
                        lblCreditHour.Enabled = True
                        cmbCreditHour.Enabled = True
                        cmdAddCourse.Enabled = True
                        cmdClear.Enabled = True
                        frmList.Enabled = True
                        lblNo.Enabled = True
                        lblCourseCodeList.Enabled = True
                        lblGradeList.Enabled = True
                        cmdConfirmList.Enabled = True
                        cmdCheckAll.Enabled = True
                        cmdUnchecked.Enabled = True
                        lblCreditList.Enabled = True
                        
                         lblCurrentSemester.Enabled = False
                        lblGradePointTotal.Enabled = False
                        lblGradePointTotal1.Enabled = False
                        lblCreditTotal.Enabled = False
                        lblCreditTotal1.Enabled = False
                        lblGradePointAverage.Enabled = False
                        lblGradePointAverage1.Enabled = False
                        cmdCalculate.Enabled = False
                        cmdPrint.Enabled = False
                        cmdExitTotal.Enabled = False
                        
                        MsgBox "Student Register Success. You can start register your course.", vbOKOnly, "Success"
                    End If
                End If
              End If
          End If
        End If
    End If
End Sub

Private Sub cmdUnchecked_Click()
'To unchecked all
chkSubject1.Value = Unchecked
chkSubject2.Value = Unchecked
chkSubject3.Value = Unchecked
chkSubject4.Value = Unchecked
chkSubject5.Value = Unchecked
chkSubject6.Value = Unchecked
chkSubject7.Value = Unchecked
End Sub

