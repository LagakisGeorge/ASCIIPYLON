VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5590
   ClientLeft      =   52
   ClientTop       =   377
   ClientWidth     =   9074
   LinkTopic       =   "Form1"
   ScaleHeight     =   5590
   ScaleWidth      =   9074
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPELATES 
      Caption         =   "PELATES"
      Height          =   360
      Left            =   819
      TabIndex        =   2
      Top             =   1521
      Width           =   990
   End
   Begin VB.TextBox txtWITHSUMS 
      Height          =   5278
      Left            =   4797
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "uSEL_PARASTAT.frx":0000
      Top             =   0
      Width           =   4342
   End
   Begin VB.CommandButton Command1 
      Caption         =   "¡—◊≈…¡ √…¡ ÷œ—«‘¡"
      Height          =   598
      Left            =   585
      TabIndex        =   0
      Top             =   351
      Width           =   1651
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   299
      Left            =   702
      Top             =   4095
      Visible         =   0   'False
      Width           =   2470
      _ExtentX        =   4361
      _ExtentY        =   527
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
         Size            =   7.47
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim gdb As New ADODB.Connection



Public Function CNull(C) As String

        '¡Õ ≈…Õ¡… NULL ≈–…”‘—≈÷≈… " "
        '<EhHeader>
        On Error GoTo CNull_Err

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

R.Open sql, gdb, adOpenDynamic, adLockOptimistic

Open "C:\CL\DATA\CUST.TXT" For Output As #1
Dim s As String

Dim fpa As String

Do While Not R.EOF
   s = ""
   s = s + Left(R!HECODE + Space(15), 15) + " "
   s = s + Left(to437(R!HEName) + Space(35), 31) + " "
   s = s + Left(to437(R!PRFSNAME) + Space(35), 22) + " "
   
   s = s + Left(to437(R!ADDRESS) + Space(30), 22) + " "
    s = s + Left(to437(R!ADDRESS) + Space(30), 22) + " "
    
    s = s + Left(to437(R!HECITY) + Space(30), 22) + " "
    
    
   
    s = s + Left(to437(R!HETIN) + Space(13), 11) + " "  'AFM
    
     s = s + Left(to437(R!TFFCNAME) + Space(30), 22) + " " 'DOY
    
    s = s + Right(Space(35) + Format(R!YPOL, "####0.00"), 13) + " "
    
    s = s + Left("00.00", 5) + " "  ' EKPT
    
      s = s + Left("0000002000.00", 13) + " "  ' PLAFON
    
     s = s + Left(to437(R!TRBRPHONE1) + Space(25), 22) + " " 'THL
     
     s = s + Left(to437(R!TRBRPHONE2) + Space(30), 22) + " " 'THL2
     
     
     
  
     s = s + Space(5) + " " 'KENA
     s = s + Space(1) + " " 'KENA
     
     
     
     
     s = s + "00000" 'ekpt
     
   
   
  
   Print #1, s
   R.MoveNext
Loop


Close #1































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

Open "C:\CL\DATA\PROD.TXT" For Output As #1
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
 
 
     fpa = "    7"
     If Val(Mid(R!vat, 2, 6)) = 1 Then
        fpa = "    7" ' 24%
     ElseIf Val(Mid(R!vat, 2, 6)) = 2 Then
        fpa = "    9" '13%
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



End Sub
  
Private Sub Form_Load()
   gdb.Open "DSN=PYLON"
   


End Sub

Function to437(CC) As String

    Dim A As String

    If IsNull(CC) Then
        A = " "
    Else
        A = CC
    End If

    Dim b, C, s As String

    Dim k As Integer

    'metatrepei eggrafo apo 437->928
    s928 = "¡¬√ƒ≈∆«»… ÀÃÕŒœ–—”‘’÷◊ÿŸ-·‚„‰ÂÊÁËÈÍÎÏÌÓÔÒÛÙıˆ˜¯˘-Ú‹›ﬁﬂ¸˝˛"
    s437 = "ÄÅÇÉÑÖÜáàâäãåçéèêëíìîïñó-òôöõúùûü†°¢£§•¶ß®©´¨≠ÆØ‡-™·‚„ÂÊÁÈ" ' saehioyv
    's437 = "ÄÅÇÉÑÖÜáàâäãåçéèêëíìîïñó-òôöõúùûü†°¢£§•¶ß®©´¨≠ÆØ‡™" '·‚„ÂÊÁÈ"
    '                                                        saehioyv

    'Open Text2.Text For Output As #2
    'Open Text1.Text For Input As #1
    'Do While Not EOF(1)
    'Line Input #1, a
    For k = 1 To Len(A)
        s = Mid(A, k, 1)
        t = InStr(s928, s)

        If t > 0 Then
            Mid$(A, k, 1) = Mid$(s437, t, 1)
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

