VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   555
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3600
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pSmall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   960
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer Timer 
      Interval        =   1
      Left            =   1560
      Top             =   120
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading File Types"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   495
      Left            =   53
      TabIndex        =   0
      Top             =   30
      Width           =   3495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''#####################################################''
''##                                                 ##''
''##  Created By BelgiumBoy_007                      ##''
''##                                                 ##''
''##  Visit BartNet @ www.bartnet.be for more Codes  ##''
''##                                                 ##''
''##  Copyright 2003 BartNet Corp.                   ##''
''##                                                 ##''
''#####################################################''

Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const MAX_PATH_LENGTH As Long = 260
Private Const GOOD_RETURN_CODE As Long = 0
Private Const STARTS_WITH_A_PERIOD As Long = 46
Private Const SH_USEFILEATTRIBUTES As Long = &H10
Private Const SH_TYPENAME As Long = &H400
Private Const SH_SHELLICONSIZE = &H4
Private Const SH_SYSICONINDEX = &H4000
Private Const SH_DISPLAYNAME = &H200
Private Const SH_EXETYPE = &H2000
Private Const BASIC_SH_FLAGS = SH_TYPENAME Or SH_SHELLICONSIZE Or SH_SYSICONINDEX Or SH_DISPLAYNAME Or SH_EXETYPE
Private Const SH_LARGEICON = &H0
Private Const SH_SMALLICON = &H1
Private Const REG_SZ = (1)
Private Const REG_EXPAND_SZ = (2)
Private Const ILD_TRANSPARENT = &H1
Private Const KEY_QUERY_VALUE = &H1
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Private Type FILETIME
    dwLowDateTime   As Long
    dwHighDateTime  As Long
End Type

Private Type SHFILEINFO
    hIcon           As Long
    iIcon           As Long
    dwAttributes    As Long
    szDisplayName   As String * MAX_PATH_LENGTH
    szTypeName      As String * 80
End Type

Private Declare Function RegEnumKeyEx Lib "advapi32" Alias "RegEnumKeyExA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal y As Long, ByVal flags As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Private Item As ListItem
Private bResize As Boolean
Private lSortCol As Long
Private iPos As Integer
Private lResult As Long
Private lrc As Long
Private rc1 As Long
Private rc2 As Long
Private rc3 As Long
Private cch As Long
Private lType As Long
Private vValue As Variant
Private sValue As String
Private sKey As String
Private sImageList1Key As String
Private iSmall As ListImage
Private Sub Form_Load()
    DoEvents
End Sub

Private Sub Timer_Timer()
On Error Resume Next
    Dim Index As Long
    Dim Subkey As String * MAX_PATH_LENGTH
    Dim KeyClass As String * MAX_PATH_LENGTH
    Dim TheTime   As FILETIME
    Dim Icon As Long
    Dim Icon2 As Long
    Dim Info As SHFILEINFO
    Dim FileTypeName As String
    Dim FileExtension As String
    
    Screen.MousePointer = vbHourglass
    
    Do While RegEnumKeyEx(HKEY_CLASSES_ROOT, Index, Subkey, MAX_PATH_LENGTH, 0, KeyClass, MAX_PATH_LENGTH, TheTime) = GOOD_RETURN_CODE
        If Asc(Subkey) = STARTS_WITH_A_PERIOD Then
            Icon2 = SHGetFileInfo(Subkey, FILE_ATTRIBUTE_NORMAL, Info, Len(Info), SH_USEFILEATTRIBUTES Or BASIC_SH_FLAGS Or SH_LARGEICON)
            Icon = SHGetFileInfo(Subkey, FILE_ATTRIBUTE_NORMAL, Info, Len(Info), SH_USEFILEATTRIBUTES Or BASIC_SH_FLAGS Or SH_SMALLICON)
            FileTypeName = TrimNull(Info.szTypeName)
            FileExtension = TrimNull(Subkey)
            FileExtension = Right(FileExtension, Len(FileExtension) - 1)
            pSmall.Picture = LoadPicture()
            Call ImageList_Draw(Icon, Info.iIcon, pSmall.hDC, 0, 0, ILD_TRANSPARENT)
            pSmall.Picture = pSmall.Image
            sImageList1Key = "#" & FileExtension & "#"
            Form1.ImageList.ListImages.Add , sImageList1Key, pSmall.Picture
            Set Item = Form1.ListView.ListItems.Add(, , , Form1.ImageList.ListImages.Item(sImageList1Key).Key, Form1.ImageList.ListImages.Item(sImageList1Key).Key)
            If FileExtension = "" Then FileExtension = "*"
            Item.SubItems(1) = "*." & FileExtension
            Item.SubItems(2) = FileTypeName
       End If
       Index = Index + 1
       
       DoEvents
    Loop
    
    Screen.MousePointer = vbDefault
    
    Unload Me
    Form1.Show
    DoEvents
End Sub
Private Function TrimNull(startstr As String) As String
On Error Resume Next
    iPos = InStr(startstr, Chr$(0))
    If iPos Then
       TrimNull = Left$(startstr, iPos - 1)
       Exit Function
    End If
    TrimNull = startstr
End Function
