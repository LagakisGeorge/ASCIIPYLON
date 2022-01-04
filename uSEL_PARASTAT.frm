VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   ScaleHeight     =   10275
   ScaleWidth      =   13815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCommand2 
      Caption         =   "Command2"
      Height          =   360
      Left            =   11520
      TabIndex        =   4
      Top             =   8520
      Visible         =   0   'False
      Width           =   990
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   372
      Left            =   6120
      Top             =   2640
      Visible         =   0   'False
      Width           =   2892
      _ExtentX        =   5106
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton EKFORTOSI 
      BackColor       =   &H0080FF80&
      Caption         =   "≈ ÷œ—‘Ÿ”« "
      Height          =   720
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   3372
   End
   Begin VB.CommandButton cmdPELATES 
      BackColor       =   &H00FFFF80&
      Caption         =   "–≈À¡‘≈”"
      Height          =   598
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1651
   End
   Begin VB.TextBox txtWITHSUMS 
      Height          =   5278
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "uSEL_PARASTAT.frx":0000
      Top             =   4080
      Visible         =   0   'False
      Width           =   12012
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "≈…ƒ« √…¡ ÷œ—«‘¡"
      Height          =   598
      Left            =   585
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   351
      Width           =   1651
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IS_MERC As Integer  ' 1= PYLON TO MERCURY    0=PYLON TO GIORANIDIS
Dim gdb As New ADODB.Connection

Dim GMERC As New ADODB.Connection





Public Function CNull(C) As String

        '¡Õ ≈…Õ¡… NULL ≈–…”‘—≈÷≈… " "
        '<EhHeader>
       ' On Error GoTo CNull_Err

        '</EhHeader>
        On Error Resume Next

100     If IsNull(C) Then
110         CNull = " "
        Else

120         CNull = C
        End If

        '<EhFooter>
        Exit Function
End Function







Private Sub cmdPELATES_Click()

Dim R As New ADODB.Recordset
Dim sql As String
sql = txtWITHSUMS.Text
sql = Replace$(sql, "2021", Format(Year(Now), "0000"))
R.Open sql, gdb, adOpenDynamic, adLockOptimistic

Open "C:\CL\DATA\CUST" For Output As #1
Dim s As String

Dim fpa As String
 If IS_MERC = 1 Then
      Dim R2 As New ADODB.Recordset
      GMERC.Execute "UPDATE PEL SET AYP=0"
 End If
 
      
Do While Not R.EOF
   s = ""
   s = s + Left(R!HECODE + Space(15), 15) + " "
   s = s + Left(to437(R!HENAME) + Space(35), 31) + " "
   s = s + Left(to437(R!EPAGG) + Space(35), 22) + " "
   
   
   If IS_MERC = 1 Then
      R2.Open "SELECT COUNT(*) FROM PEL   WHERE  EIDOS='e' and KOD='" + R!HECODE + "'", GMERC, adOpenDynamic, adLockOptimistic
      If R2(0) = 0 Then
          GMERC.Execute "INSERT INTO PEL (EIDOS,KOD) VALUES ('e','" + R!HECODE + "')"
      End If
      R2.Close
   
  
 
   GMERC.Execute "update PEL set EPO='" + Left(R!HENAME, 35) + "'  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
   
    GMERC.Execute "update PEL set EPA='" + R!EPAGG + "'  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
   
   GMERC.Execute "update PEL set DIE='" + Left(R!DIE, 35) + "'  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
   
   GMERC.Execute "update PEL set POL='" + R!POL + "'  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
   
   GMERC.Execute "update PEL set AFM='" + R!AFM + "'  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
   GMERC.Execute "update PEL set DOY='" + Replace(R!DOY, "'", "`") + "'  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
   GMERC.Execute "update PEL set AYP=AYP+" + Replace(Format(R!DD, "####0.00"), ",", ".") + "  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
   
   GMERC.Execute "update PEL set THL='" + Left(R!THL1, 10) + "'  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
   
   GMERC.Execute "update PEL set PEK=" + Replace(Format(R!EKPTOSIS, "####0.00"), ",", ".") + "  WHERE  EIDOS='e' and KOD='" + R!HECODE + "'"
 End If
   
   ' Left(to437(R!THL1)
   
   
   
   
   s = s + Left(to437(R!DIE) + Space(30), 22) + " "
    s = s + Left(to437(R!DIE) + Space(30), 22) + " "
    
    s = s + Left(to437(R!POL) + Space(30), 22) + " "
    
    
   
    s = s + Left(to437(R!AFM) + Space(13), 11) + " "  'AFM
    
     s = s + Left(to437(R!DOY) + Space(30), 22) + " " 'DOY
    
    s = s + Right(Space(35) + Format(R!DD, "####0.00"), 13) + " "
    
    s = s + Left("00.00", 5) + " "  ' EKPT
    
      s = s + Left("0000002000.00", 13) + " "  ' PLAFON
    
     s = s + Left(to437(R!THL1) + Space(25), 22) + " " 'THL
     
     s = s + Left(to437(R!THL2) + Space(30), 22) + " " 'THL2
     
     
     
  
     s = s + Space(5) + " " 'KENA
     s = s + Space(1) + " " 'KENA
     
     
     
     
     s = s + "00000" 'ekpt
     
   
   
  
   Print #1, s
   R.MoveNext
Loop


Close #1


MsgBox " –≈À¡‘≈” œ "




























End Sub





Private Sub Command1_Click()
Dim R As New ADODB.Recordset
Dim sql As String
sql = "SELECT v.HECODE as vat, u.hename AS MON,i.hecode AS KOD ,i.hename AS NAME, "
sql = sql + "  ISNULL(HEWHOLESALESPRICE,0) AS PRICE "
sql = sql + " FROM dbo.HEITEMS i WITH (NOLOCK) "
sql = sql + " LEFT OUTER JOIN HEMEASUREMENTUNITS u WITH (NOLOCK) ON i.HEAMSNTID = u.HEID"
sql = sql + " inner join [HEVATCLASSES] v  on (I.[HEVTCLID] = v.[HEID])"


   R.Open sql, gdb, adOpenDynamic, adLockOptimistic

If IS_MERC = 1 Then
     Dim R2 As New ADODB.Recordset
End If

Open "C:\CL\DATA\PROD" For Output As #1
Dim s As String

Dim fpa As String

Do While Not R.EOF
   s = ""
   s = s + Left(R!KOD + Space(15), 15) + " "
   s = s + Left(to437(R!Name) + Space(35), 35) + " "
   s = s + Left(R!KOD + Space(35), 15) + " "
   
   s = s + Left(to437(R!MON) + Space(3), 3) + " "
    s = s + Left(to437(R!MON) + Space(3), 3) + " "
     s = s + Left(Space(5), 5) + " "  'METATR
     
 s = s + Right(Space(35) + Format(R!PRICE, "####0.00"), 13) + " "
 
 
 
 
 If IS_MERC = 1 Then
 
 R2.Open "SELECT COUNT(*) FROM EID   WHERE   KOD='" + R!KOD + "'", GMERC, adOpenDynamic, adLockOptimistic
   If R2(0) = 0 Then
          GMERC.Execute "INSERT INTO EID (KOD) VALUES ('" + R!KOD + "')"
   End If
   R2.Close
   




   GMERC.Execute "update EID set ONO='" + R!Name + "'  WHERE   KOD='" + R!KOD + "'"
   
   GMERC.Execute "update EID set MON='" + R!MON + "'  WHERE   KOD='" + R!KOD + "'"
   GMERC.Execute "update EID set LTI=" + Replace(Format(R!PRICE, "####0.00"), ",", ".") + "  WHERE   KOD='" + R!KOD + "'"
   
End If

 
 
 
 
 
 
     fpa = "    7"
     If Val(Mid(R!vat, 2, 6)) = 1 Then
        fpa = "    7" ' 24%
        If IS_MERC = 1 Then
            GMERC.Execute "update EID set FPA=2  WHERE   KOD='" + R!KOD + "'"
        End If
     ElseIf Val(Mid(R!vat, 2, 6)) = 2 Then
        fpa = "    9" '13%
        If IS_MERC = 1 Then
            GMERC.Execute "update EID set FPA=1  WHERE   KOD='" + R!KOD + "'"
        End If
      ElseIf Val(Mid(R!vat, 2, 6)) = 3 Then
        fpa = "    6"
      ElseIf Val(Mid(R!vat, 2, 6)) = 2 Then
        fpa = "    0"
     End If
     s = s + fpa + " "
     s = s + Space(13) + " " 'ypol
     s = s + Space(5) + " " 'ekpt
     s = s + Space(4)
     
 
'   S = S + Left(R!HENAME + Space(35), 35) + " "
'   S = S + Left(R!HENAME + Space(35), 35) + " "
'   S = S + Left(R!HENAME + Space(35), 35) + " "
'   S = S + Left(R!HENAME + Space(35), 35) + " "
'
   
      
   
   
   
 '  "  " + R("HECODE")
   Print #1, s
   R.MoveNext
Loop


Close #1

MsgBox "≈…ƒ« œ "


End Sub
  
Private Sub EKFORTOSI_Click()
 
 
 
Dim R As New ADODB.Recordset
Dim sql As String
Dim GDB2 As New ADODB.Connection
GDB2.Open "DSN=IMPORTS;"

GDB2.Execute "delete from [Last_ComeBack_PaperProducts] "

GDB2.Execute "delete from [Last_ComeBack_Papers] "
 
 
 
 FILL_Last_ComeBack_Papers
 
 
 
 
 Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim strConnection As String

    Dim strProvider As String
    Dim strSource As String

    strProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
   ' strSource = "Data Source=\ADOPROG2\GIORAN\DB.mdb;"  '& App.Path &
    strSource = "Data Source=\SFV\DB\DB.mdb;"  '& App.Path &

    Set cnn = New ADODB.Connection
    strConnection = strProvider & strSource & "Persist Security Info=False"
    cnn.Open strConnection
    Set rst = New ADODB.Recordset

    With rst
        .ActiveConnection = cnn
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open "[Last_ComeBack_PaperProducts]"
    End With

    'Clear combo box
   Dim K As Integer
   
Dim MMID As Long


    Dim MFN As String
    'Load combo box
    Do While Not rst.EOF
     ' cboData.AddItem rst!fldData
     ' cboData.ItemData(cboData.NewIndex) = rst!fldID
       
      GDB2.Execute "INSERT INTO [Last_ComeBack_PaperProducts] (PaperCode) VALUES (" + Str(rst(0)) + ") "
      
      
     R.Open "SELECT MAX(ID) FROM Last_ComeBack_PaperProducts", GDB2, adOpenDynamic, adLockOptimistic
 
     MMID = R(0)
     R.Close
     
      ' GDB2.Execute "UPDATE Last_ComeBack_PaperProducts SET PaperCode=" + rst(0) + "' WHERE ID=" + Str(MMID)
          
 
       
     
       For K = 1 To rst.Fields.Count - 1
          MFN = rst.Fields(K).Name
          GDB2.Execute "UPDATE Last_ComeBack_PaperProducts SET " + MFN + "='" + rst(K) + "' WHERE ID=" + Str(MMID)
          
       
       Next
     
     
     
     
      rst.MoveNext  '
   Loop

   ' Set initial display to first item in ListIndex
  ' cboData.ListIndex = 0






'
' Do While Not REGGTIM.EOF
'
'            If printCrystal < 0 Then Exit Do
'
'320         R.AddNew
'
'            'On Error GoTo printCrystal_Err
'
'330         For k = 0 To REGGTIM.Fields.Count - 1
'
'                On Error GoTo LATOS
'
'340             If REGGTIM(k).Type = 202 Or REGGTIM(k).Type = 200 Or REGGTIM(k).Type = 129 Then    ' STRING
'                    'R(REGGTIM(k).Name) = Left(REGGTIM(k), R(REGGTIM(k).Name).Size)
'350                 R(k) = Left(REGGTIM(k), R(k).Size)
'                Else
'360                 R(k) = REGGTIM(k)
'                    'R(REGGTIM(k).Name) = REGGTIM(k)
'                End If
'
'            Next
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'


MsgBox "≈‘œ…Ã« « Ã≈‘¡÷œ—¡ ‘ŸÕ  …Õ«”≈ŸÕ." + Chr(13) + "–—œ◊Ÿ—≈…”‘≈ ”‘«Õ Ã≈‘¡÷œ—¡ ‘œ’ PYLON"



End Sub



Private Sub cmdCommand2_Click()


   FILL_Last_ComeBack_Papers
End Sub


Sub FILL_Last_ComeBack_Papers()
'========================================================================

Dim R As New ADODB.Recordset
Dim sql As String
Dim GDB2 As New ADODB.Connection
GDB2.Open "DSN=IMPORTS;"


 
 
 
 
 
 
 
 
 Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim strConnection As String

    Dim strProvider As String
    Dim strSource As String

    strProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
    strSource = "Data Source=\SFV\DB\DB.mdb;" '& App.Path &
   '"Data Source=\SFV\DB\DB.mdb;"
    Set cnn = New ADODB.Connection
    strConnection = strProvider & strSource & "Persist Security Info=False"
    cnn.Open strConnection
    Set rst = New ADODB.Recordset

    With rst
        .ActiveConnection = cnn
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open "[Last_ComeBack_Papers]"
    End With

    'Clear combo box
   Dim K As Integer
   
Dim MMID As Long


    Dim MFN As String
    'Load combo box
    Do While Not rst.EOF
     ' cboData.AddItem rst!fldData
     ' cboData.ItemData(cboData.NewIndex) = rst!fldID
       
      GDB2.Execute "INSERT INTO [Last_ComeBack_Papers] (PaperCode,printed,canceled) VALUES (" + Str(rst(0)) + ",0,0 ) "
      
      
     R.Open "SELECT MAX(ID) FROM Last_ComeBack_Papers", GDB2, adOpenDynamic, adLockOptimistic
 
     MMID = R(0)
     R.Close
     
      ' GDB2.Execute "UPDATE Last_ComeBack_PaperProducts SET PaperCode=" + rst(0) + "' WHERE ID=" + Str(MMID)
          
 
       
     
       For K = 1 To rst.Fields.Count - 1
        If K = 9 Or K = 19 Then ' BIT
        Else
          MFN = rst.Fields(K).Name
          GDB2.Execute "UPDATE Last_ComeBack_Papers SET " + MFN + "='" + rst(K) + "' WHERE ID=" + Str(MMID)
        End If
        
       
       Next
     
     
     
     
      rst.MoveNext  '
   Loop

 









End Sub















Private Sub Form_Load()
  
  
  
  IS_MERC = 0
  
  If Len(Dir("C:\MERCVB", vbDirectory)) > 0 Then
      IS_MERC = 1
  End If
  
  Me.Caption = IS_MERC
  
   
  If IS_MERC = 1 Then
     gdb.Open "DSN=PYLON;uid=sa;pwd=p@ssw0rd"
     
     GMERC.Open "DSN=MERCSQL"
  Else
     gdb.Open "DSN=PYLON;uid=sa;pwd=p@ssw0rd"
  
  End If
  
  
  



End Sub

Function to437(CC) As String

    Dim A As String

    If IsNull(CC) Then
        A = " "
    Else
        A = CC
    End If

    Dim b, C, s As String

    Dim K As Integer

    'metatrepei eggrafo apo 437->928
    s928 = "¡¬√ƒ≈∆«»… ÀÃÕŒœ–—”‘’÷◊ÿŸ-·‚„‰ÂÊÁËÈÍÎÏÌÓÔÒÛÙıˆ˜¯˘-Ú‹›ﬁﬂ¸˝˛"
    s437 = "ÄÅÇÉÑÖÜáàâäãåçéèêëíìîïñó-òôöõúùûü†°¢£§•¶ß®©´¨≠ÆØ‡-™·‚„ÂÊÁÈ" ' saehioyv
    's437 = "ÄÅÇÉÑÖÜáàâäãåçéèêëíìîïñó-òôöõúùûü†°¢£§•¶ß®©´¨≠ÆØ‡™" '·‚„ÂÊÁÈ"
    '                                                        saehioyv

    'Open Text2.Text For Output As #2
    'Open Text1.Text For Input As #1
    'Do While Not EOF(1)
    'Line Input #1, a
    For K = 1 To Len(A)
        s = Mid(A, K, 1)
        t = InStr(s928, s)

        If t > 0 Then
            Mid$(A, K, 1) = Mid$(s437, t, 1)
        End If

    Next

    'Print #2, a$
    to437 = A

    'Loop
    'Close #1
    'Close #2

End Function

'QUERY –œ’ ’–œÀœ√…∆≈… ‘œ’” –≈À¡‘≈” ¡ÀÀ¡ ¬√¡∆≈… ’–œÀœ…–¡ ¡Õ¡ ’–œ ¡‘¡”‘«Ã¡

'WITH SUMS AS (
'
'select CSTM.HEID , CSTM.HECODE, [CSTM].[HENAME] ,SUM(ISNULL(TRCM.HEEBALANCE,0)) AS YPOL
', [CNTC].[HETIN]
'
', [TRBR].[HECITY]
', [TRBR].[HEPOSTALCODE], [TRBR].[HEPHONE1]
'[TRBRPHONE1], [TRBR].[HEPHONE2] [TRBRPHONE2], [TRBR].[HEPHONE3] [TRBRPHONE3], [TRBR].[HEPHONE4] [TRBRPHONE4], [TRBR].[HEPHONE5] [TRBRPHONE5],
'[TRBR].[HEEMAIL] [TRBREMAIL],
'   Trbr.heStreet+' '+coalesce(Trbr.heStreetNumber, '') [ADDRESS],
'  [HETAXOFFICES].[HENAME] [TFFCNAME], [HEPROFESSIONS].[HENAME]
'[PRFSNAME], [HEINDUSTRIES].[HENAME] [IDRSNAME], [HECOUNTRIES].[HENAME] [CNTRNAME],
' [HECOUNTRIES].[HENAME] [DSTTNAME], [HECOUNTRIES].[HENAME] [MNCPNAME], [CATEGORIESCSTM1].[HENAME] [CCSTNAME1],
'[CATEGORIESCSTM2].[HENAME] [CCSTNAME2], [TREESCSTM1].[HENAME] [TCSTNAME1], [TREESCSTM2].[HENAME] [TCSTNAME2],
' [CURR].[HECODE] [CURRCODE], [CURR].[HENAME] [CURRNAME], [TRDRACCCAT].[HECODE]
'[TRDRACCCATCODE], [TRDRACCCAT].[HENAME] [TRDRACCCATNAME], [WHOLESALEPRCLIST].[HECODE] [WHOLESALEPRCLISTCODE],
'[WHOLESALEPRCLIST].[HENAME] [WHOLESALEPRCLISTNAME], [RETAILPRCLIST].[HECODE]
'[RETAILPRCLISTCODE], [RETAILPRCLIST].[HENAME] [RETAILPRCLISTNAME], [PAYMENTMETHOD].[HECODE] [PAYMENTMETHODCODE],
' [PAYMENTMETHOD].[HENAME] [PAYMENTMETHODNAME], [ADDCHRGGRP].[HECODE]
'[ADDCHRGGRPCODE], [ADDCHRGGRP].[HENAME] [ADDCHRGGRPNAME], [SHIPPINGMETHOD].[HECODE] [SHIPPINGMETHODCODE],
'[SHIPPINGMETHOD].[HENAME] [SHIPPINGMETHODNAME], [VATSTS].[HECODE] [VATSTSCODE],
'[VATSTS].[HENAME] [VATSTSNAME], [VATEXM].[HECODE] [VATEXMCODE], [VATEXM].[HENAME] [VATEXMNAME],
' [SALESMAN].[HECODE] [SALESMANCODE], [SALESMAN].[HENAME] [SALESMANNAME], [COLLECTOR].[HECODE]
'[COLLECTORCODE], [COLLECTOR].[HENAME] [COLLECTORNAME]
'
'
'
'from [HECUSTOMERS] [CSTM] WITH(NOLOCK)
'inner join [HETRADERS] [TRDR] WITH(NOLOCK)  on ([CSTM].[HETRDRID] = [TRDR].[HEID])
'Inner Join
'[HECONTACTS] [CNTC] WITH(NOLOCK)  on ([TRDR].[HECNTCID] = [CNTC].[HEID])
'inner join [HETRADERBRANCHES] [TRBR] WITH(NOLOCK)  on ([TRBR].[HETRDRID] = [TRDR].[HEID] )
'left join [HETAXOFFICES] WITH(NOLOCK)  on ([TRDR].[HETFFCID] = [HETAXOFFICES].[HEID])
'left join [HEPROFESSIONS] WITH(NOLOCK)  on ([TRBR].[HEPRFSID] = [HEPROFESSIONS].[HEID])
'Left
'join [HEINDUSTRIES] WITH(NOLOCK)  on ([TRDR].[HEIDRSID] = [HEINDUSTRIES].[HEID])
'left join [HECOUNTRIES] WITH(NOLOCK)  on ([TRBR].[HECNTRID] = [HECOUNTRIES].[HEID])
'left join [HEDISTRICTS]
'WITH(NOLOCK)  on ([TRBR].[HEDSTTID] = [HEDISTRICTS].[HEID])
'left join [HEMUNICIPALITIES] WITH(NOLOCK)  on ([TRBR].[HEMNCPID] = [HEMUNICIPALITIES].[HEID])
'left join [HECATEGORIESCSTM] [CATEGORIESCSTM1] WITH(NOLOCK)  on ([CSTM].[HECAT01ID] = [CATEGORIESCSTM1].[HEID])
'left join [HECATEGORIESCSTM] [CATEGORIESCSTM2] WITH(NOLOCK)  on ([CSTM].[HECAT02ID] = [CATEGORIESCSTM2].[HEID])
'left join [HETREESCSTM] [TREESCSTM1] WITH(NOLOCK)  on ([CSTM].[HETREECAT01ID] = [TREESCSTM1].[HEID])
'left join [HETREESCSTM] [TREESCSTM2] WITH(NOLOCK)  on ([CSTM].[HETREECAT02ID] = [TREESCSTM2].[HEID])
'inner join [HECURRENCIES] [CURR] WITH(NOLOCK)  on ([CSTM].[HECURRID] = [CURR].[HEID])
'left join [HETRDRACCCATEGORIES] [TRDRACCCAT] WITH(NOLOCK)  on ([CSTM].[HETRACID] = [TRDRACCCAT].[HEID])
'inner join [HECUSTOMERBRANCHES] [CSBR] WITH(NOLOCK)  on ([TRBR].[HEID] = [CSBR].[HETRBRID] and [CSTM].[HEID] = [CSBR].[HECSTMID])
'left join [HEPRICELISTS] [WHOLESALEPRCLIST] WITH(NOLOCK)  on ([CSBR].[HEPRLSID] = [WHOLESALEPRCLIST].[HEID])
'left join [HEPRICELISTS] [RETAILPRCLIST] WITH(NOLOCK)  on ([CSBR].[HERETAILPRLSID] = [RETAILPRCLIST].[HEID])
'left join [HEPAYMENTMETHODS] [PAYMENTMETHOD] WITH(NOLOCK)  on ([CSBR].[HEPMMTID] = [PAYMENTMETHOD].[HEID])
'left join [HEADDDOCCHARGESGROUPS] [ADDCHRGGRP] WITH(NOLOCK)  on ([CSBR].[HEADCGID] = [ADDCHRGGRP].[HEID])
'left join [HESHIPPINGMETHODS] [SHIPPINGMETHOD] WITH(NOLOCK)  on ([CSBR].[HESHPMID] = [SHIPPINGMETHOD].[HEID])
'left join [HEVATSTATUSES] [VATSTS] WITH(NOLOCK)  on ([CSBR].[HEVTSTID] = [VATSTS].[HEID])
'left join [HEVATEXEMPTIONS] [VATEXM] WITH(NOLOCK)  on ([CSBR].[HEVTEXID] = [VATEXM].[HEID])
'left join [HEAGENTS] [SALESMAN] WITH(NOLOCK)  on ([CSBR].[HEAGNTID] = [SALESMAN].[HEID])
'left join [HEAGENTS] [COLLECTOR] WITH(NOLOCK)  on ([CSBR].[HECOLLAGNTID] = [COLLECTOR].[HEID])
'LEFT join [HETRADERACCUMULATORS] [TRCM] WITH(NOLOCK)  on ([CSTM].[HEID] = [TRCM].[HECSTMID] and [TRBR].[HEID] = [TRCM].[HETRBRID])
'Where [CSTM].[HEACTIVE] = 1 And Year(TRCM.HEDATE) = 2021
'
'Group By
' CSTM.HEID , [CSTM].[HENAME],
'CSTM.HECODE , [CSTM].[HENAME], HETIN
'
',  [TRBR].[HECITY]
', [TRBR].[HEPOSTALCODE] , [TRBR].[HEPHONE1]
', [TRBR].[HEPHONE2] , [TRBR].[HEPHONE3] , [TRBR].[HEPHONE4] , [TRBR].[HEPHONE5] , [TRBR].[HEEMAIL]  ,
'    Trbr.heStreet+' '+coalesce(Trbr.heStreetNumber, ''),
'  [HETAXOFFICES].[HENAME] , [HEPROFESSIONS].[HENAME]
', [HEINDUSTRIES].[HENAME] , [HECOUNTRIES].[HENAME] , [HECOUNTRIES].[HENAME] ,
' [HECOUNTRIES].[HENAME] , [CATEGORIESCSTM1].[HENAME] ,
'[CATEGORIESCSTM2].[HENAME] , [TREESCSTM1].[HENAME], [TREESCSTM2].[HENAME], [CURR].[HECODE], [CURR].[HENAME], [TRDRACCCAT].[HECODE]
', [TRDRACCCAT].[HENAME] , [WHOLESALEPRCLIST].[HECODE] , [WHOLESALEPRCLIST].[HENAME] , [RETAILPRCLIST].[HECODE]
', [RETAILPRCLIST].[HENAME] , [PAYMENTMETHOD].[HECODE], [PAYMENTMETHOD].[HENAME] , [ADDCHRGGRP].[HECODE]
', [ADDCHRGGRP].[HENAME] , [SHIPPINGMETHOD].[HECODE] , [SHIPPINGMETHOD].[HENAME] , [VATSTS].[HECODE] ,
'[VATSTS].[HENAME] , [VATEXM].[HECODE], [VATEXM].[HENAME], [SALESMAN].[HECODE], [SALESMAN].[HENAME], [COLLECTOR].[HECODE]
', [COLLECTOR].[HENAME]
'
')
'
'SELECT
' HEID,YPOL,
'[HENAME],
'HECODE , HETIN
'
',  [HECITY]
',[HEPOSTALCODE] , [TRBRPHONE1]
', [TRBRPHONE2] ,[TRBRPHONE3] , [TRBRPHONE4] , [TRBRPHONE5] ,[TRBREMAIL] ,
'   ADDRESS ,
' TFFCNAME , [PRFSNAME]
', [IDRSNAME],   [CNTRNAME],[DSTTNAME],
'  [MNCPNAME],   [CCSTNAME1],
'[CCSTNAME2] ,
' [TCSTNAME1],  [TCSTNAME2],
'  [CURRCODE],  [CURRNAME],
'[TRDRACCCATCODE],  [TRDRACCCATNAME],[WHOLESALEPRCLISTCODE],
' [WHOLESALEPRCLISTNAME],
'[RETAILPRCLISTCODE],[RETAILPRCLISTNAME],  [PAYMENTMETHODCODE],
'  [PAYMENTMETHODNAME],
'[ADDCHRGGRPCODE],  [ADDCHRGGRPNAME], [SHIPPINGMETHODCODE],
' [SHIPPINGMETHODNAME],  [VATSTSCODE],
' [VATSTSNAME], [VATEXMCODE],  [VATEXMNAME],
' [SALESMANCODE],  [SALESMANNAME],
'[COLLECTORCODE] , [COLLECTORNAME]
'
'
'
' From SUMS
'
' Where YPOL <> 0
' ORDER BY HEID
'

'

