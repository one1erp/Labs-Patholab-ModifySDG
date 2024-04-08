VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.UserControl ModifySDGCtrl 
   ClientHeight    =   8910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   15360
   Begin VB.Frame FrameSDGHeader 
      Caption         =   "ש"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   0
      Width           =   15015
      Begin VB.TextBox RequestText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10800
         TabIndex        =   15
         Top             =   250
         Width           =   2300
      End
      Begin VB.Label LblRequestNum 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   ":מס. דרישה"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13620
         TabIndex        =   17
         Top             =   300
         Width           =   1260
      End
      Begin VB.Label RequestLabel 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6600
         TabIndex        =   16
         Top             =   300
         Width           =   3972
      End
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "בצע"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   13
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton CloseButton 
      Caption         =   "סגור"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   12
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Frame FrameSDGDetail 
      Caption         =   "פרטי דרישה"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   15015
      Begin VB.PictureBox Frame1 
         Height          =   6615
         Left            =   120
         ScaleHeight     =   6555
         ScaleWidth      =   14685
         TabIndex        =   1
         Top             =   240
         Width           =   14750
         Begin VB.PictureBox Frame2 
            BorderStyle     =   0  'None
            Height          =   625
            Left            =   120
            ScaleHeight     =   630
            ScaleWidth      =   14295
            TabIndex        =   3
            Top             =   45
            Width           =   14295
            Begin VB.ComboBox CmbFieldType 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   0
               Left            =   960
               RightToLeft     =   -1  'True
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   90
               Visible         =   0   'False
               Width           =   7155
            End
            Begin VB.TextBox TxtFieldType 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   0
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Top             =   83
               Visible         =   0   'False
               Width           =   7155
            End
            Begin VB.CheckBox CheckFieldType 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   7920
               TabIndex        =   6
               Top             =   120
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.CommandButton CmdFind 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   5
               Top             =   120
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.CommandButton CmdOpenCalander 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   4
               Top             =   120
               Visible         =   0   'False
               Width           =   465
            End
            Begin MSACAL.Calendar Calendar 
               Height          =   3015
               Left            =   1080
               TabIndex        =   18
               Top             =   0
               Visible         =   0   'False
               Width           =   4455
               _Version        =   524288
               _ExtentX        =   7858
               _ExtentY        =   5318
               _StockProps     =   1
               BackColor       =   -2147483633
               Year            =   2004
               Month           =   10
               Day             =   18
               DayLength       =   1
               MonthLength     =   1
               DayFontColor    =   0
               FirstDay        =   7
               GridCellEffect  =   1
               GridFontColor   =   10485760
               GridLinesColor  =   -2147483632
               ShowDateSelectors=   -1  'True
               ShowDays        =   -1  'True
               ShowHorizontalGrid=   -1  'True
               ShowTitle       =   -1  'True
               ShowVerticalGrid=   -1  'True
               TitleFontColor  =   10485760
               ValueIsNull     =   0   'False
               BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label LblFieldTitle 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   420
               Index           =   0
               Left            =   8580
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   90
               Visible         =   0   'False
               Width           =   5500
            End
            Begin VB.Label LblFieldType 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   420
               Index           =   0
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   10
               Top             =   90
               Visible         =   0   'False
               Width           =   7155
            End
            Begin VB.Label LblFindObject 
               Appearance      =   0  'Flat
               BackColor       =   &H80000013&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   420
               Index           =   0
               Left            =   10080
               TabIndex        =   9
               Top             =   90
               Visible         =   0   'False
               Width           =   1215
            End
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   6555
            Left            =   14430
            TabIndex        =   2
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
      End
   End
End
Attribute VB_Name = "ModifySDGCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements LSExtensionWindowLib.IExtensionWindow
Implements LSExtensionWindowLib.IExtensionWindow2

Option Explicit

'הגדרת צבעים גלובליים
Private Const RED = &HFF&
Private Const WHITE = &HFFFFFF
Private Const BLACK = &H80000008
Private Const FONT_SIZE = 12

Private ProcessXML As LSSERVICEPROVIDERLib.NautilusProcessXML
Private NtlsCon As LSSERVICEPROVIDERLib.NautilusDBConnection
Private NtlsSite As LSExtensionWindowLib.IExtensionWindowSite2
Private NtlsUser As LSSERVICEPROVIDERLib.NautilusUser

Private Con As ADODB.Connection
Private Sdg As ADODB.Recordset
Private SdgFlag As Boolean
Private OpenedRequest As Boolean
Private SaveLine As Integer
Private sdg_log As New SdgLog.CreateLog
Private sdg_log_desc As String

Public Type LineRec
    FieldName As String
    FieldType As String
    ID As String
    SDGModifyID As String
    ParentID As String
    ComboIDs As Collection
    ComboQuery As String
    TextQuery As String
End Type
Private LinesRec As New Collection

Public RunFromWindow As Boolean
Public Event CloseClicked()

'05.09.2006
'holds the valid role-names to activate this program:
Private strPrivilagedOperators As String

Private Sub CloseButton_Click()
On Error GoTo ERR_CloseButton_Click
    Dim MBRes As VbMsgBoxResult

    If Not RunFromWindow Then
        MBRes = MsgBox("? האם ברצונך לצאת מהמסך", vbYesNo + vbDefaultButton2, "Nautilus - Modify SDG")
        If MBRes = vbNo Then Exit Sub
    End If

   ' Call zLang.SetOrigLang

    If RunFromWindow Then
        RaiseEvent CloseClicked
    Else
  If Not NtlsSite Is Nothing Then
        NtlsSite.CloseWindow
      End If
    End If
    
    Exit Sub
ERR_CloseButton_Click:
MsgBox "ERR_CloseButton_Click" & vbCrLf & Err.Description
End Sub

Private Sub CmbFieldType_Change(Index As Integer)
    RefreshText (Index)
End Sub

Private Sub CmbFieldType_Click(Index As Integer)
    RefreshText (Index)
End Sub

Private Sub CmbFieldType_GotFocus(Index As Integer)
    RefreshText (Index)
End Sub

Private Sub CmbFieldType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RefreshText (Index)
End Sub

Private Sub CmbFieldType_KeyPress(Index As Integer, KeyAscii As Integer)
    RefreshText (Index)
End Sub

Private Sub CmbFieldType_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    RefreshText (Index)
End Sub

Private Sub CmbFieldType_LostFocus(Index As Integer)
    RefreshText (Index)
End Sub

Private Sub CmbFieldType_Scroll(Index As Integer)
    RefreshText (Index)
End Sub

Private Sub CmdFind_Click(Index As Integer)
On Error GoTo ERR_CmdFind_Click
    Dim o As Object
    Dim TmpID As String
    Dim Currline As LineRec

    Set o = CreateObject(LblFindObject(Index).Caption)
    Set o.Con = Con
    o.ShowDlg
    If o.Description <> "" Then
        LblFieldType(Index).Caption = o.Description
        Currline = LinesRec(CStr(Index))
        Call LinesRec.Remove(CStr(Index))
        TmpID = Currline.SDGModifyID
        Currline.ID = o.ID
        Call LinesRec.Add(Currline, (CStr(Index)))
'        UpdateSDG (Index)
        Call RefreshParentID(Index, TmpID)
    End If
    
    Exit Sub
ERR_CmdFind_Click:
MsgBox "ERR_CmdFind_Click" & vbCrLf & Err.Description
End Sub

Private Sub Calendar_Click()
On Error GoTo ERR_Calendar_Click
    If Calendar.Value <> "" Then
        LblFieldType(SaveLine).Caption = Calendar.Value
    End If
    Calendar.Visible = False
    
    Exit Sub
ERR_Calendar_Click:
MsgBox "ERR_Calendar_Click" & vbCrLf & Err.Description
End Sub

Private Sub CmdOpenCalander_Click(Index As Integer)
On Error GoTo ERR_CmdOpenCalander_Click
    Calendar.Value = ""
    If Calendar.Visible = False Then
        Calendar.Visible = True
        SaveLine = Index
        If LblFieldType(Index).Caption <> "" Then
            Calendar.Value = LblFieldType(Index).Caption
        Else
            Calendar.Value = Now
        End If
    Else
        Calendar.Visible = False
        If Calendar.Value <> "" Then
            LblFieldType(Index).Caption = Calendar.Value
        End If
    End If
    
    Exit Sub
ERR_CmdOpenCalander_Click:
MsgBox "ERR_CmdOpenCalander_Click" & vbCrLf & Err.Description
End Sub

Public Function IExtensionWindow_CloseQuery() As Boolean
    'Happens when the user close the window
    Set Sdg = Nothing
    IExtensionWindow_CloseQuery = True
End Function

Public Function IExtensionWindow_DataChange() As LSExtensionWindowLib.WindowRefreshType
    IExtensionWindow_DataChange = windowRefreshNone
End Function

Public Function IExtensionWindow_GetButtons() As LSExtensionWindowLib.WindowButtonsType
    IExtensionWindow_GetButtons = windowButtonsNone
End Function

Public Sub IExtensionWindow_Internationalise()
End Sub

Public Sub IExtensionWindow_PreDisplay()
On Error GoTo ERR_IExtensionWindow_PreDisplay
    Set Sdg = New ADODB.Recordset
    Set Con = New ADODB.Connection
    Dim CS As String
        CS = NtlsCon.GetADOConnectionString

          If NtlsCon.GetServerIsProxy Then
            CS = "Provider=OraOLEDB.Oracle;Data Source=" & _
            NtlsCon.GetServerDetails & ";User id=/;Persist Security Info=True;"
          End If
    Con.Open CS
    Con.CursorLocation = adUseClient
    Con.Execute "SET ROLE LIMS_USER"

    VScroll1.LargeChange = 20  ' Cross in 5 clicks.
    VScroll1.SmallChange = 5   ' Cross in 20 clicks.
    Frame2.Container = Frame1

    Call ConnectSameSession(CDbl(NtlsCon.GetSessionId))

    Set sdg_log.Con = Con
    sdg_log.Session = CDbl(NtlsCon.GetSessionId)

    SdgFlag = False
    
    With Calendar
        .GridFont.Name = "Arial"
        .GridFont.Size = 9
        .GridFont.Bold = False
        .DayFont.Name = "Arial"
        .DayFont.Size = 9
        .DayFont.Bold = True
        .TitleFont.Name = "Arial"
        .TitleFont.Size = 12
        .TitleFont.Bold = True
        .FirstDay = 7 'sunday
    End With
    
    Exit Sub
ERR_IExtensionWindow_PreDisplay:
MsgBox "ERR_IExtensionWindow_PreDisplay" & vbCrLf & Err.Description
End Sub

Public Sub IExtensionWindow_refresh()
    'Code for refreshing the window
End Sub

Public Sub IExtensionWindow_RestoreSettings(ByVal hKey As Long)
End Sub

Public Function IExtensionWindow_SaveData() As Boolean
End Function

Public Sub IExtensionWindow_SaveSettings(ByVal hKey As Long)
End Sub

Public Sub IExtensionWindow_SetParameters(ByVal parameters As String)
On Error GoTo ERR_IExtensionWindow_SetParameters
    
    '05.09.2006
    strPrivilagedOperators = parameters
    
    Exit Sub
ERR_IExtensionWindow_SetParameters:
End Sub

Public Sub IExtensionWindow_SetServiceProvider(ByVal serviceProvider As Object)
    Dim sp As LSSERVICEPROVIDERLib.NautilusServiceProvider
    Set sp = serviceProvider
    Set ProcessXML = sp.QueryServiceProvider("ProcessXML")
    Set NtlsCon = sp.QueryServiceProvider("DBConnection")
    Set NtlsUser = sp.QueryServiceProvider("User")
End Sub

Public Sub IExtensionWindow_SetSite(ByVal Site As Object)
    Set NtlsSite = Site
    If RunFromWindow Then Exit Sub
    NtlsSite.SetWindowInternalName ("MacabiModifySDG")
    NtlsSite.SetWindowRegistryName ("MacabiModifySDG")
    Call NtlsSite.SetWindowTitle("Modify SDG")
End Sub

Public Sub IExtensionWindow_Setup()
On Error GoTo ERR_IExtensionWindow_Setup
    sdg_log_desc = ""

  '  Call zLang.English
    RequestText.Alignment = vbLeftJustify
    RequestText.RightToLeft = False

    If Not RunFromWindow Then
        Call RequestText.SetFocus
    End If

    OpenedRequest = False
    

    '05.09.2006
    'call to check if operator is valid:
    If strPrivilagedOperators <> "" Then
        Call CheckOperatorValidation(NtlsUser.GetOperatorId(), strPrivilagedOperators)
    End If
    
    
    Exit Sub
ERR_IExtensionWindow_Setup:
MsgBox "ERR_IExtensionWindow_Setup" & vbCrLf & Err.Description
End Sub

Public Function IExtensionWindow_ViewRefresh() As LSExtensionWindowLib.WindowRefreshType
    IExtensionWindow_ViewRefresh = windowRefreshNone
End Function

Private Sub ConnectSameSession(ByVal aSessionID)
On Error GoTo ERR_ConnectSameSession
    Dim aProc As New ADODB.Command
    Dim aSession As New ADODB.Parameter
    
    aProc.ActiveConnection = Con
    aProc.CommandText = "lims.lims_env.connect_same_session"
    aProc.CommandType = adCmdStoredProc

    aSession.Type = adDouble
    aSession.Direction = adParamInput
    aSession.Value = aSessionID
    aProc.parameters.Append aSession

    aProc.Execute
    Set aSession = Nothing
    Set aProc = Nothing
    
    Exit Sub
ERR_ConnectSameSession:
MsgBox "ERR_ConnectSameSession" & vbCrLf & Err.Description
End Sub

'05.09.2006
'allow only users with the right role
'to enter the application
Private Sub CheckOperatorValidation(strOperatorId As String, strParameters As String)
On Error GoTo ERR_CheckOperatorValidation
    Dim sql As String
    Dim rs As Recordset

    sql = " select r.NAME "
    sql = sql & " from lims_sys.operator o, lims_sys.lims_role r"
    sql = sql & " where r.ROLE_ID=o.ROLE_ID"
    sql = sql & " and   o.OPERATOR_ID = " & strOperatorId

    Set rs = Con.Execute(sql)

    If rs.EOF = True Then Exit Sub
 
    
    If InStr(1, strParameters, nte(rs("NAME")), vbTextCompare) = 0 Then
        MsgBox "Insufficient privileges for this operation"
        If Not NtlsSite Is Nothing Then
        Call NtlsSite.CloseWindow
        End If
    End If

    Exit Sub
ERR_CheckOperatorValidation:
MsgBox "GoTo ERR_CheckOperatorValidation" & vbCrLf & Err.Description
End Sub

Public Sub IExtensionWindow2_Close()
End Sub

Private Function nte(e As Variant) As Variant
    nte = IIf(IsNull(e), "", e)
End Function

Private Sub OkButton_Click()
On Error GoTo ERR_OkButton_Click
    Dim i As Integer
    Dim UpdateFlag As Boolean

    UpdateFlag = False
    sdg_log_desc = ""
    For i = 1 To LblFieldTitle.Count - 1
        If LblFieldTitle(i).Enabled = True Then
            UpdateFlag = True
            UpdateSDG (i)
        End If
    Next i
    If UpdateFlag = True Then
        Call sdg_log.InsertLog(Sdg("SDG_ID"), "UP.UPD", sdg_log_desc)
        MsgBox " הדרישה עודכנה בהצלחה ", , "Nautilus - Modify SDG"

    '    Call zLang.SetOrigLang

        If RunFromWindow Then
            RaiseEvent CloseClicked
        Else
   If Not NtlsSite Is Nothing Then
            NtlsSite.CloseWindow
               End If
        End If
    End If
    
    Exit Sub
ERR_OkButton_Click:
MsgBox "ERR_OkButton_Click" & vbCrLf & Err.Description
End Sub

Private Sub RequestText_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_RequestText_KeyDown
    If Not KeyCode = vbKeyReturn Then Exit Sub
    Dim rst As ADODB.Recordset
    Dim StrTitile As String
    Dim RowNum As Integer
    Dim Hi As Long
    Dim CommaPosition As Integer
    

    If OpenedRequest Then
        If MsgBox("הדרישה לא עודכנה." & vbCrLf & _
                " ? האם אתה בטוח שברצונך להמשיך ", vbYesNo) = vbNo Then
            RequestText.Text = ""
            Exit Sub
        End If
    End If
     
      
    'Text Correction (changed 04/2008)
     RequestText.Text = Trim(UCase(RequestText.Text))
     CommaPosition = InStr(1, RequestText.Text, ".")
     If CommaPosition <> 0 Then
        RequestText.Text = Trim(Left(RequestText.Text, CommaPosition - 1))
     End If
    
        
    
    If Len(RequestText.Text) < 8 Then
       ' RequestText.Text = ""
        MsgBox " ! מינימום 8 תווים למס. הדרישה ", , "Nautilus - Modify SDG"
        Call RequestText.SetFocus
        Exit Sub
    End If
  
    Set Sdg = Con.Execute("select * from lims_sys.sdg, lims_sys.sdg_user where " & _
        "sdg.sdg_id = sdg_user.sdg_id and " & _
       "(sdg.name = '" & UCase(RequestText.Text) & "' OR sdg_user.U_PATHOLAB_NUMBER= '" & UCase(RequestText.Text) & "')")

    If Sdg.EOF Then
        'RequestText.Text = ""
        MsgBox " ! מס הדרישה אינה קיימת במערכת ", , "Nautilus - Modify SDG"
        Call RequestText.SetFocus
        Exit Sub
    End If

    If Sdg("STATUS") = "X" Then
        RequestText.Text = ""
        MsgBox " ! הדרישה בוטלה ", , "Nautilus - Modify SDG"
        Call RequestText.SetFocus
        Exit Sub
    End If

    RowNum = 0
    Call UnloadFileds

    Set rst = Con.Execute("select * from lims_sys.u_modify_sdg, lims_sys.u_modify_sdg_user " & _
        "where u_modify_sdg.u_modify_sdg_id = u_modify_sdg_user.u_modify_sdg_id " & _
        "and u_visible = 'T' " & _
        "order by u_modify_sdg_user.u_order")

    If rst.EOF Then
        RequestText.Text = ""
        MsgBox " לא נמצאו שדות להצגה ", , "Nautilus - מאקרו היסטולוגיה"
        Call RequestText.SetFocus
        Exit Sub
    Else
        rst.MoveFirst
    End If
    While Not rst.EOF
        RowNum = RowNum + 1
        Hi = InitSDGScreen(rst, RowNum)
        rst.MoveNext
    Wend

    Frame2.Height = Hi + 170
    VScroll1.Max = (Frame2.Height - Frame1.Height) / ScaleHeight * 100
    If VScroll1.Max < 0 Then
        VScroll1.Visible = False
    Else
        VScroll1.Visible = True
    End If

    RequestLabel.Caption = UCase(RequestText.Text) & " - " & nte(Sdg("U_PATHOLAB_NUMBER"))
    StrTitile = "Modify SDG - " & UCase(RequestText.Text)
          If Not NtlsSite Is Nothing Then
              Call NtlsSite.SetWindowTitle(StrTitile)
               End If


    SdgFlag = True
    RequestText.Text = ""
 '   Call zLang.SetOrigLang
    
    Exit Sub
ERR_RequestText_KeyDown:
MsgBox "ERR_RequestText_KeyDown" & vbCrLf & Err.Description
End Sub

Private Sub UnloadFileds()
On Error GoTo ERR_UnloadFileds
    Dim i As Integer
    Dim j As Integer

    For i = 1 To LinesRec.Count
'        For j = 1 To LinesRec(i).ComboIDs.Count
'            Call LinesRec(j).ComboIDs.Remove(1)
'        Next j
        Call LinesRec.Remove(1)
    Next i

    For i = 1 To LblFieldTitle.Count - 1
        Unload LblFieldTitle(i)
    Next

    For i = 1 To CmbFieldType.Count - 1
        Unload CmbFieldType(i)
    Next

    For i = 1 To TxtFieldType.Count - 1
        Unload TxtFieldType(i)
    Next

    For i = 1 To LblFieldType.Count - 1
        Unload LblFieldType(i)
    Next

    For i = 1 To CheckFieldType.Count - 1
        Unload CheckFieldType(i)
    Next

    For i = 1 To LblFindObject.Count - 1
        Unload LblFindObject(i)
    Next

    For i = 1 To CmdFind.Count - 1
        Unload CmdFind(i)
    Next

    For i = 1 To CmdOpenCalander.Count - 1
        Unload CmdOpenCalander(i)
    Next
    
    Exit Sub
ERR_UnloadFileds:
MsgBox "ERR_UnloadFileds" & vbCrLf & Err.Description
End Sub

Private Function InitSDGScreen(RSTRec As ADODB.Recordset, RowNum As Integer) As Integer
On Error GoTo ERR_InitSDGScreen
    Dim i As Integer
    Dim rstTemp As ADODB.Recordset
    Dim strSQL As String
    Dim itm As Collection 'Dictionary
    Dim strAuthorizedRoles As String, strUserRoleId As String

    ' for title label
    Load LblFieldTitle(RowNum)
    With LblFieldTitle(RowNum)
        .Top = 500 * (RowNum - 1) + 10
        .Caption = " " & RSTRec("U_LABEL")
        .Visible = True
        InitSDGScreen = .Top + .Height
        If RSTRec("U_READ_ONLY") = "T" Then
            .Enabled = False
        ElseIf RSTRec("U_READ_ONLY") = "F" Then
            .Enabled = True
        Else
            .Enabled = True
        End If
        .FontBold = False
        If RSTRec("U_BOLD") = "T" Then
            .FontBold = True
        End If
        If RSTRec("U_FONT_COLOR") <> "" Then
            .ForeColor = RSTRec("U_FONT_COLOR")
        End If
        .FontSize = FONT_SIZE
        If RSTRec("U_FONT_SIZE") <> "" And RSTRec("U_FONT_SIZE") > 0 Then
            .FontSize = RSTRec("U_FONT_SIZE")
        End If
    End With

    ' for check box
    Load CheckFieldType(RowNum)
    With CheckFieldType(RowNum)
        .Top = 500 * (RowNum - 1) + 10
        If RSTRec("U_FIELD_TYPE") = "B" Then
            If Sdg(RSTRec("U_NAME").Value) = "T" Then
                .Value = 1
            Else
                .Value = 0
            End If
        End If
        .Visible = True
        If RSTRec("U_FONT_COLOR") <> "" Then
            .ForeColor = RSTRec("U_FONT_COLOR")
        End If
        If RSTRec("U_READ_ONLY") = "T" Then
            .Enabled = False
        ElseIf RSTRec("U_READ_ONLY") = "F" Then
            .Enabled = True
        Else
            .Enabled = True
        End If
    End With

    ' for text box
    Load TxtFieldType(RowNum)
    With TxtFieldType(RowNum)
'        .Alignment = nvl(RSTRec("U_ALIGNMENT").Value, 0)
        .Alignment = 1
        .Top = 500 * (RowNum - 1) + 10
        If RSTRec("U_FIELD_TYPE") = "T" Then
            .Text = nvl(Sdg(RSTRec("U_NAME").Value).Value, "")
        End If
        .Visible = True
        If RSTRec("U_READ_ONLY") = "T" Then
            .Enabled = False
        ElseIf RSTRec("U_READ_ONLY") = "F" Then
            .Enabled = True
        Else
            .Enabled = True
        End If
        .FontBold = False
        If RSTRec("U_BOLD") = "T" Then
            .FontBold = True
        End If
        .FontSize = FONT_SIZE
        If RSTRec("U_FONT_SIZE") <> "" And RSTRec("U_FONT_SIZE") > 0 Then
            .FontSize = RSTRec("U_FONT_SIZE")
        End If
        If RSTRec("U_FONT_COLOR") <> "" Then
            .ForeColor = RSTRec("U_FONT_COLOR")
        End If
    End With

    ' for combo box
    Load CmbFieldType(RowNum)
    With CmbFieldType(RowNum)
        .Top = 500 * (RowNum - 1) + 10
        .Visible = True
        If RSTRec("U_READ_ONLY") = "T" Then
            .Enabled = False
        ElseIf RSTRec("U_READ_ONLY") = "F" Then
            .Enabled = True
        Else
            .Enabled = True
        End If
        .FontBold = False
        If RSTRec("U_BOLD") = "T" Then
            .FontBold = True
        End If
        .FontSize = FONT_SIZE
        If RSTRec("U_FONT_SIZE") <> "" And RSTRec("U_FONT_SIZE") > 0 Then
            .FontSize = RSTRec("U_FONT_SIZE")
        End If
    End With

    ' for label
    Load LblFieldType(RowNum)
    With LblFieldType(RowNum)
'        .Alignment = nvl(RSTRec("U_ALIGNMENT").Value, 0)
        .Top = 500 * (RowNum - 1) + 10
        .Visible = True
        If RSTRec("U_READ_ONLY") = "T" Then
            .Enabled = False
        ElseIf RSTRec("U_READ_ONLY") = "F" Then
            .Enabled = True
        Else
            .Enabled = True
        End If
        .FontBold = False
        If RSTRec("U_BOLD") = "T" Then
            .FontBold = True
        End If
        .FontSize = FONT_SIZE
        If RSTRec("U_FONT_SIZE") <> "" And RSTRec("U_FONT_SIZE") > 0 Then
            .FontSize = RSTRec("U_FONT_SIZE")
        End If
        If RSTRec("U_FONT_COLOR") <> "" Then
            .ForeColor = RSTRec("U_FONT_COLOR")
        End If
        If RSTRec("U_FIELD_TYPE") = "D" Then
            .Caption = nvl(Sdg(RSTRec("U_NAME").Value).Value, "")
        End If
    End With

    ' for find object label
    Load LblFindObject(RowNum)
    With LblFindObject(RowNum)
        .Top = 500 * (RowNum - 1) + 10
        If RSTRec("U_FIND_OBJECT") <> "" Then
            .Caption = nvl(RSTRec("U_FIND_OBJECT"), "")
        End If
    End With

    ' for the command button
    Load CmdFind(RowNum)
    strAuthorizedRoles = nte(RSTRec("U_AUTHORIZED_ROLE"))
    With CmdFind(RowNum)
        .Top = 500 * (RowNum - 1) + 10
        .Visible = True
        If RSTRec("U_READ_ONLY") = "T" Then
            .Enabled = False
            
        '11.07.2011
        'only authorized roles can access this CmdFind
        'checking if there are roles for this record, if there are - checking if roleID is among them
        ElseIf strAuthorizedRoles <> "" Then
            strAuthorizedRoles = "," & strAuthorizedRoles & ","
            strUserRoleId = CStr(NtlsUser.GetRoleId())
            strUserRoleId = "," & strUserRoleId & ","
            If InStr(1, strAuthorizedRoles, strUserRoleId, vbTextCompare) = 0 Then
                .Enabled = False
            End If
        Else
            .Enabled = True
        End If
    End With

    ' for the calender button
    Load CmdOpenCalander(RowNum)
    With CmdOpenCalander(RowNum)
        .Top = 500 * (RowNum - 1) + 10
        .Visible = True
        If RSTRec("U_READ_ONLY") = "T" Then
            .Enabled = False
        ElseIf RSTRec("U_READ_ONLY") = "F" Then
            .Enabled = True
        Else
            .Enabled = True
        End If
    End With

    Select Case RSTRec("U_FIELD_TYPE")
        Case "B"
            CheckFieldType(RowNum).Visible = True
            TxtFieldType(RowNum).Visible = False
            CmbFieldType(RowNum).Visible = False
            LblFieldType(RowNum).Visible = False
            CmdFind(RowNum).Visible = False
            CmdOpenCalander(RowNum).Visible = False

        Case "T"
            CheckFieldType(RowNum).Visible = False
            TxtFieldType(RowNum).Visible = True
            CmbFieldType(RowNum).Visible = False
            LblFieldType(RowNum).Visible = False
            CmdFind(RowNum).Visible = False
            CmdOpenCalander(RowNum).Visible = False

        Case "C"
            CheckFieldType(RowNum).Visible = False
            TxtFieldType(RowNum).Visible = False
            CmbFieldType(RowNum).Visible = True
            LblFieldType(RowNum).Visible = False
            CmdFind(RowNum).Visible = False
            CmdOpenCalander(RowNum).Visible = False

        Case "L"
            CheckFieldType(RowNum).Visible = False
            TxtFieldType(RowNum).Visible = False
            CmbFieldType(RowNum).Visible = False
            LblFieldType(RowNum).Visible = True
            CmdFind(RowNum).Visible = True
            CmdOpenCalander(RowNum).Visible = False

        Case "D"
            CheckFieldType(RowNum).Visible = False
            TxtFieldType(RowNum).Visible = False
            CmbFieldType(RowNum).Visible = False
            LblFieldType(RowNum).Visible = True
            CmdFind(RowNum).Visible = False
            CmdOpenCalander(RowNum).Visible = True
    End Select

    Call InitTypeRec(RSTRec, RowNum)
    
    Exit Function
ERR_InitSDGScreen:
MsgBox "ERR_InitSDGScreen" & vbCrLf & Err.Description
End Function

Private Sub VScroll1_Change()
    Frame2.Top = -(VScroll1.Value / 100) * ScaleHeight + 50
End Sub

Private Function nvl(e As Variant, v As Variant) As Variant
  nvl = IIf(IsNull(e), v, e)
End Function

Private Sub UpdateSDG(i As Integer)
On Error GoTo ERR_UpdateSDG
    Dim strSQL As String
    Dim Currline As LineRec

    Currline = LinesRec(CStr(i))

    strSQL = "update lims_sys.sdg"
    If Left(Currline.FieldName, 2) = "U_" Then
            strSQL = strSQL & "_user"
    End If
    strSQL = strSQL & " set " & _
              Currline.FieldName & _
              " = "
    Select Case Currline.FieldType
        Case "L", "C"
            If Currline.ID <> "" Then
                strSQL = strSQL & "'" & Currline.ID & "'"
            Else
                strSQL = strSQL & "NULL"
            End If
        Case "B"
            If CheckFieldType(i).Value = 0 Then
                strSQL = strSQL & "'F'"
            ElseIf CheckFieldType(i).Value = 1 Then
                strSQL = strSQL & "'T'"
            End If
        Case "T"
            strSQL = strSQL & "'" & TxtFieldType(i).Text & "'"
        Case "D"
            strSQL = strSQL & "to_date('" & LblFieldType(i).Caption & "', 'dd/mm/yyyy')"
    End Select

    strSQL = strSQL & " where sdg_id = " & Sdg("SDG_ID")
'    MsgBox "strSQL = " & strSQL
    Call Con.Execute(strSQL)
    
    Exit Sub
ERR_UpdateSDG:
MsgBox "ERR_UpdateSDG" & vbCrLf & Err.Description
End Sub

Private Sub RefreshParentID(Index As Integer, SDGModifyID As String)
On Error GoTo ERR_RefreshParentID
    Dim i As Integer
    Dim j As Integer
    Dim Currline As LineRec
    Dim rstTemp As ADODB.Recordset

    For i = 1 To LblFieldTitle.Count - 1

        Currline = LinesRec(CStr(i))

        If (Currline.ParentID = SDGModifyID) And _
                (LblFieldTitle(i).Enabled = True) And _
                (i <> Index) Then

            Currline.ID = ""

            For j = 1 To Currline.ComboIDs.Count
                  Currline.ComboIDs.Remove (1)
            Next j

            If Currline.ComboQuery <> "" Then

                Set rstTemp = Con.Execute(Currline.ComboQuery)
                CmbFieldType(i).Clear
                CmbFieldType(i).List(0) = "None"
                If Not rstTemp.EOF Then
                    rstTemp.MoveFirst
                End If

                Set Currline.ComboIDs = New Collection
                While Not rstTemp.EOF
                    CmbFieldType(i).List(CmbFieldType(i).ListCount) = rstTemp("description") ' nvl(rstTemp("description"), "None")

                    ' fill the sub collection
                    ' -----------------------
                    Call Currline.ComboIDs.Add(rstTemp("id").Value, rstTemp("description").Value)

                    rstTemp.MoveNext
                Wend
                Call rstTemp.Close
            End If

            If CmbFieldType(i).ListCount > 0 Then CmbFieldType(i).ListIndex = 0
            Currline.ID = ""

            Call LinesRec.Remove(CStr(i))

            Call LinesRec.Add(Currline, (CStr(i)))

            UpdateSDG (i)

        End If

    Next i
    
    Exit Sub
ERR_RefreshParentID:
MsgBox "ERR_RefreshParentID" & vbCrLf & Err.Description
End Sub

Private Sub RefreshText(Index As Integer)
On Error GoTo ERR_RefreshText
    If Index > LinesRec.Count Then Exit Sub
    Dim Currline As LineRec
    Dim i As Integer

    Currline = LinesRec(CStr(Index))
    Call LinesRec.Remove(CStr(Index))

'    MsgBox CmbFieldType(Index).Text
'    For i = 1 To currline.ComboIDs.Count
'        MsgBox currline.ComboIDs(i)
'    Next i

    If CmbFieldType(Index).Text = "None" Then
        Currline.ID = ""
    Else
        Currline.ID = Currline.ComboIDs(CmbFieldType(Index).Text)
    End If
    Call LinesRec.Add(Currline, (CStr(Index)))
    
    Exit Sub
ERR_RefreshText:
MsgBox "ERR_RefreshText" & vbCrLf & Err.Description
End Sub

Private Sub InitTypeRec(RSTRec As ADODB.Recordset, RowNum As Integer)
On Error GoTo ERR_InitTypeRec
    Dim itm As LineRec
    Dim rstTemp As ADODB.Recordset
    Dim strSQL As String

    itm.FieldName = nvl(RSTRec("U_NAME"), "")
    itm.FieldType = nvl(RSTRec("U_FIELD_TYPE"), "")
    itm.SDGModifyID = nvl(RSTRec("U_MODIFY_SDG_ID"), "")
    itm.ParentID = nvl(RSTRec("U_PARENT_ID"), "")
    itm.ComboQuery = Replace(nvl(RSTRec("U_COMBO_QUERY"), ""), "#SDG_ID#", Sdg("SDG_ID"))
    itm.TextQuery = Replace(nvl(RSTRec("U_TEXT_QUERY"), ""), "#SDG_ID#", Sdg("SDG_ID"))

    ' for the combo box
    strSQL = nvl(RSTRec("U_COMBO_QUERY"), "")
    If strSQL <> "" Then
        strSQL = Replace(strSQL, "#SDG_ID#", Sdg("SDG_ID"))
        Set rstTemp = Con.Execute(strSQL)
        CmbFieldType(RowNum).Clear
        CmbFieldType(RowNum).List(0) = "None"
        If Not rstTemp.EOF Then
            rstTemp.MoveFirst
        End If
        Set itm.ComboIDs = New Collection
        While Not rstTemp.EOF
            CmbFieldType(RowNum).List(CmbFieldType(RowNum).ListCount) = rstTemp("description")

            ' fill the sub collection
            ' -----------------------
            Call itm.ComboIDs.Add(rstTemp("id").Value, rstTemp("description").Value)

            rstTemp.MoveNext
        Wend
        Call rstTemp.Close
    End If

    ' for the text
    strSQL = nvl(RSTRec("U_TEXT_QUERY"), "")
    If strSQL <> "" Then
        strSQL = Replace(strSQL, "#SDG_ID#", Sdg("SDG_ID"))
        Set rstTemp = Con.Execute(strSQL)
        If Not rstTemp.EOF Then
            If itm.FieldType = "C" Then
                CmbFieldType(RowNum).Text = nvl(rstTemp("description"), "None")
            End If
            If itm.FieldType = "L" Then
                LblFieldType(RowNum).Caption = nvl(rstTemp("description"), "")
            End If
            '05.09.2006
            'check also for a query if it's a text field:
            If itm.FieldType = "T" Then
                TxtFieldType(RowNum).Text = nvl(rstTemp("description"), "")
            End If
            itm.ID = rstTemp("id")
        Else
            If CmbFieldType(RowNum).ListCount > 0 Then CmbFieldType(RowNum).ListIndex = 0
            itm.ID = ""
        End If
        Call rstTemp.Close
    End If

    ' add a new record to type
    Call LinesRec.Add(itm, CStr(RowNum))
    
    Exit Sub
ERR_InitTypeRec:
MsgBox "ERR_InitTypeRec" & vbCrLf & Err.Description
End Sub

Private Sub UserControl_Initialize()
    RunFromWindow = False
End Sub

Public Sub InitiateSdg(sn As String)
    RequestText.Text = sn
    Call RequestText_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_UserControl_KeyDown
    Dim strVer As String

    If KeyCode = vbKeyF10 And Shift = 1 Then
        strVer = "Name: " & App.EXEName & vbCrLf & vbCrLf & _
                 "Path: " & App.Path & vbCrLf & vbCrLf & _
                 "Version: " & "[" & App.Major & "." & App.Minor & "." & App.Revision & "]" & vbCrLf & vbCrLf & _
                 "Company: One Software Technologies (O.S.T) Ltd."
        MsgBox strVer, vbInformation, "Nautilus - Project Properties"
        Call RequestText.SetFocus
    End If
    
    Exit Sub
ERR_UserControl_KeyDown:
MsgBox "ERR_UserControl_KeyDown" & vbCrLf & Err.Description
End Sub





