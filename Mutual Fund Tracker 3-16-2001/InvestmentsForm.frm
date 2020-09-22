VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form InvestmentsForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mutal Fund Tracker"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   Icon            =   "InvestmentsForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ComboProfile 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   360
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGridList 
      Height          =   810
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   1429
      _Version        =   393216
      FixedCols       =   0
      ScrollBars      =   2
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Manage Profiles"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save To Excel"
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8040
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Get Results"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2730
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   4815
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7320
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox tickersymbol 
      Enabled         =   0   'False
      Height          =   405
      Left            =   8640
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   5055
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9735
      Begin VB.Label Label4 
         Caption         =   "(Click Headers to Sort)"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Current Profile:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Results:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Mutual Funds in Profile:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   2295
      End
   End
End
Attribute VB_Name = "InvestmentsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' FOR SORTING
Dim column0 As String
Dim column1 As String
Dim column2 As String
Dim column3 As String
Dim column4 As String
Dim column5 As String
Dim column6 As String
Dim column7 As String
Dim column8 As String
Dim column9 As String
Dim column10 As String
Dim column11 As String
Dim column12 As String
Dim column13 As String
Dim SelectedProfile As String

Dim ProfileArray(2, 11) As String

Private Sub cmdAdd_Click()
AddSymbols.Show
End Sub

Private Sub ComboProfile_Click()
Dim iAt As Integer

SelectedProfile = ComboProfile.Text

regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker", "LastSelected", SelectedProfile

For iAt = 1 To 10

    If ProfileArray(0, iAt) = SelectedProfile Then
       
        If regDoes_Key_Exist(HKEY_LOCAL_MACHINE, "Software\MutualFundTracker") Then
        
            tickersymbol.Text = regQuery_A_Key(HKEY_LOCAL_MACHINE, _
                                "Software\MutualFundTracker\Profile" & CStr(iAt), _
                                "List")
                            
            PopulateListNew
                            
        End If
         
    End If

Next

End Sub

Private Sub Command3_Click()

Screen.MousePointer = 11

Dim Symbol As Variant
Dim SymbolUbound As Integer
Dim iSymbolAt As Integer

Symbol = Split(tickersymbol.Text, ";")

SymbolUbound = UBound(Symbol)

MSFlexGrid1.Rows = SymbolUbound + 2

For iSymbolAt = 0 To SymbolUbound

    GetInformationGrid Symbol(iSymbolAt), (iSymbolAt + 1)

Next

Screen.MousePointer = 0

End Sub

Private Sub Form_Load()

Dim name As String
Dim iAt As Integer
Dim ComboIndex As Integer

' SET UP THE SORTING STUFF
column0 = ""
column1 = ""
column2 = ""
column3 = ""
column4 = ""
column5 = ""
column6 = ""
column7 = ""
column8 = ""
column9 = ""
column10 = ""
column11 = ""
column12 = ""
column13 = ""

' GOT TO CREATE ALL THE KEYS THE FIRST TIME
CreateRegistryKeys

' LOAD THE COMBO BOX

If regDoes_Key_Exist(HKEY_LOCAL_MACHINE, "Software\MutualFundTracker") Then
    
    ComboIndex = 0
    
    For iAt = 1 To 10

        name = Trim(regQuery_A_Key(HKEY_LOCAL_MACHINE, _
                               "Software\MutualFundTracker\Profile" & CStr(iAt), _
                                "Name"))
             
             ComboProfile.AddItem name, ComboIndex
             
             ComboIndex = ComboIndex + 1
             
             ProfileArray(0, iAt) = name
             ProfileArray(1, iAt) = iAt
      
    Next
    
End If

' SET UP THE HEADERS ON THE RESULTS GRID
BuildHeaders

End Sub

Private Sub MSFlexGrid1_Click()

Dim iSort As Integer

Select Case MSFlexGrid1.Col

    Case 0
    
     If column0 = "" Then
        column0 = "A"
        iSort = 1
     ElseIf column0 = "A" Then
        column0 = "D"
        iSort = 2
     Else
        column0 = "A"
        iSort = 1
     End If
     
    Case 1
    
     If column1 = "" Then
        column1 = "A"
        iSort = 1
     ElseIf column1 = "A" Then
        column1 = "D"
        iSort = 2
     Else
        column1 = "A"
        iSort = 1
     End If
     
    Case 2
    
     If column2 = "" Then
        column2 = "A"
        iSort = 1
     ElseIf column2 = "A" Then
        column2 = "D"
        iSort = 2
     Else
        column2 = "A"
        iSort = 1
     End If
     
    Case 3
    
     If column3 = "" Then
        column3 = "A"
        iSort = 1
     ElseIf column3 = "A" Then
        column3 = "D"
        iSort = 2
     Else
        column3 = "A"
        iSort = 1
     End If
     
    Case 4
    
     If column4 = "" Then
        column4 = "A"
        iSort = 1
     ElseIf column4 = "A" Then
        column4 = "D"
        iSort = 2
     Else
        column4 = "A"
        iSort = 1
     End If
     
    Case 5
    
     If column5 = "" Then
        column5 = "A"
        iSort = 1
     ElseIf column5 = "A" Then
        column5 = "D"
        iSort = 2
     Else
        column5 = "A"
        iSort = 1
     End If
     

    Case 6
    
     If column6 = "" Then
        column6 = "A"
        iSort = 1
     ElseIf column6 = "A" Then
        column6 = "D"
        iSort = 2
     Else
        column6 = "A"
        iSort = 1
     End If
     
    Case 7
    
     If column7 = "" Then
        column7 = "A"
        iSort = 1
     ElseIf column7 = "A" Then
        column7 = "D"
        iSort = 2
     Else
        column7 = "A"
        iSort = 1
     End If
     
    Case 8
    
     If column8 = "" Then
        column8 = "A"
        iSort = 1
     ElseIf column8 = "A" Then
        column8 = "D"
        iSort = 2
     Else
        column8 = "A"
        iSort = 1
     End If
     
    Case 9
    
     If column9 = "" Then
        column9 = "A"
        iSort = 1
     ElseIf column9 = "A" Then
        column9 = "D"
        iSort = 2
     Else
        column9 = "A"
        iSort = 1
     End If
     
    Case 10
    
     If column10 = "" Then
        column10 = "A"
        iSort = 1
     ElseIf column10 = "A" Then
        column10 = "D"
        iSort = 2
     Else
        column10 = "A"
        iSort = 1
     End If
     
    Case 11
    
     If column11 = "" Then
        column11 = "A"
        iSort = 1
     ElseIf column11 = "A" Then
        column11 = "D"
        iSort = 2
     Else
        column11 = "A"
        iSort = 1
     End If
     
    Case 12
    
     If column12 = "" Then
        column12 = "A"
        iSort = 1
     ElseIf column12 = "A" Then
        column12 = "D"
        iSort = 2
     Else
        column12 = "A"
        iSort = 1
     End If
     
     Case 13
     
     If column13 = "" Then
        column13 = "A"
        iSort = 1
     ElseIf column13 = "A" Then
        column13 = "D"
        iSort = 2
     Else
        column13 = "A"
        iSort = 1
     End If

End Select

MSFlexGrid1.Col = MSFlexGrid1.Col
MSFlexGrid1.Sort = iSort

End Sub

Public Function ClearHTMLTags(strHTML As Variant, intWorkFlow As Variant) As Variant

        'Variables used in the function

        Dim regEx, strTagLess

        '---------------------------------------
        strTagLess = strHTML
        'Move the string into a private variable
        'within the function
        '---------------------------------------

        'regEx initialization

        '---------------------------------------
        Set regEx = New RegExp
        'Creates a regexp object
        regEx.IgnoreCase = True
        'Don't give frat about case sensitivity
        regEx.Global = True
        'Global applicability
        '---------------------------------------


        'Phase I
        '   "bye bye html tags"


        If intWorkFlow <> 1 Then

            '---------------------------------------
            regEx.Pattern = "<[^>]*>"
            'this pattern mathces any html tag
            strTagLess = regEx.Replace(strTagLess, "")
            'all html tags are stripped
            '---------------------------------------

        End If
        
       If intWorkFlow <> 1 Then

            '---------------------------------------
            regEx.Pattern = "&nbsp;"
            'this pattern mathces any html tag
            strTagLess = regEx.Replace(strTagLess, "")
            'all html tags are stripped
            '---------------------------------------

        End If
        
       If intWorkFlow <> 1 Then

            '---------------------------------------
            regEx.Pattern = "  "
            'this pattern mathces any html tag
            strTagLess = regEx.Replace(strTagLess, "")
            'all html tags are stripped
            '---------------------------------------

        End If
               

        'Phase II
        '   "bye bye rouge leftovers"
        '   "or, I want to render the source"
        '   "as html."

        '---------------------------------------
        'We *might* still have rouge < and >
        'let's be positive that those that remain
        'are changed into html characters
        '---------------------------------------


        If intWorkFlow > 0 And intWorkFlow < 3 Then


            regEx.Pattern = "[<]"
            'matches a single <
            strTagLess = regEx.Replace(strTagLess, "&lt;")

            regEx.Pattern = "[>]"
            'matches a single >
            strTagLess = regEx.Replace(strTagLess, "&gt;")
            '---------------------------------------

        End If


        'Clean up

        '---------------------------------------
        Set regEx = Nothing
        'Destroys the regExp object
        '---------------------------------------

        '---------------------------------------
        ClearHTMLTags = strTagLess
        'The results are passed back
        '---------------------------------------

    End Function


Private Sub Command4_Click()
  
Dim sDate
Dim sTime
Dim sResult

MSFlexGrid1.Row = 1
MSFlexGrid1.Col = 0

sResult = Trim(MSFlexGrid1.Text)



If MSFlexGrid1.Rows = 2 And sResult = "" Then

    MsgBox "You must have some results to save to an Excel file."
    
Else

    sDate = Format(Date, "dddd mmm d yyyy")
    sTime = Time
      
    CommonDialog1.FileName = ComboProfile.Text & " " & sDate & ".xls"
    CommonDialog1.Filter = "Excel File(*.xls)|*.xls"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.ShowSave
    ExportGrid MSFlexGrid1, CommonDialog1.FileName

End If

End Sub

Public Function SaveTextToFile(FileFullPath As String, _
 sText As String, Optional Overwrite As Boolean = False) As _
 Boolean
    
'Purpose: Save Text to a file
'Parameters:
        '-- FileFullPath - Directory/FileName to save file to
        '-- sText - Text to write to file
       '-- Overwrite (optional): If true, if the file exists, it
                                 'is overwritten.  If false,
                                 'contents are appended to file
                                  'if the file exists

'Returns:   True if successful, false otherwise

'Example:
'SaveTextToFile "C:\My Documents\MyFile.txt", "Hello There"

On Error GoTo ErrorHandler
Dim iFileNumber As Integer
iFileNumber = FreeFile

If Overwrite Then
    Open FileFullPath For Output As #iFileNumber
Else
    Open FileFullPath For Append As #iFileNumber
End If

Print #iFileNumber, sText
SaveTextToFile = True

ErrorHandler:
Close #iFileNumber
End Function

Private Sub GetInformationGrid(ByVal sTickerSymbol, ByVal iRowAt As Integer)
Dim objhttp As New MSXML.XMLHTTPRequest
Dim strResponse As String
Dim sCleaned As String
Dim sCleaned2 As String
Dim sCleaned3 As String
Dim lStartPosition As Long
Dim iBasicData As Integer
Dim sName As String
Dim i1WeekStart As Integer
Dim sWeek1Return As String
Dim i13WeekStart As Integer
Dim sWeek13Return As String
Dim iYTDStart As Integer
Dim sYTDReturn As String
Dim i1YearStart As Integer
Dim s1YearReturn As String
Dim i3YearStart As Integer
Dim s3YearReturn As String
Dim i5YearStart As Integer
Dim s5YearReturn As String
Dim i10YearStart As Integer
Dim s10YearReturn As String
Dim iSEStart As Integer
Dim sSEReturn As String
Dim iExpensesStart As Integer

Dim iNAVStart As Integer
Dim sNAVReturn As String

Dim iLastStart As Integer
Dim slastReturn As String

Dim ichangeStart As Integer
Dim schangeReturn As String

Dim ichangepStart As Integer
Dim schangepReturn As String

Dim iAssetsStart As Integer

Dim sReturn As String

'objhttp.Open "GET", "http://funds.marketwatch.com/index.phtml?ticker=" & sTickerSymbol, False
'objhttp.Open "GET", "http://zacks.dbc.com/zcom/index.phtml?ticker=avlfx", False
objhttp.Open "GET", "http://zacks.dbc.com/zcom/index.phtml?ticker=" & sTickerSymbol, False
'objhttp.Open "GET", "http://167.8.29.44/lipper/pgPerformance.asp?TICKER=AVLFX", False
objhttp.send
strResponse = objhttp.responseText

'Text1.Text = strResponse
'Exit Sub

Set objhttp = Nothing

If Mid(strResponse, 1, 2) = "<!" Then
    'GetInformation = sTickerSymbol & " cannot be recognized " & vbCrLf & vbCrLf
    Exit Sub
End If

sCleaned = ClearHTMLTags(strResponse, 0)



sCleaned = Replace(sCleaned, vbCr, "")

sCleaned = Replace(sCleaned, vbCrLf, "")

sCleaned = Replace(sCleaned, vbLf, "")

'MsgBox sCleaned

sCleaned = Replace(sCleaned, vbTab, "")

'Text1.Text = sCleaned
'Exit Sub

'MsgBox sCleaned

lStartPosition = InStr(1, sCleaned, "Updated:")

sCleaned2 = Mid(sCleaned, lStartPosition + 8)

lStartPosition = InStr(1, sCleaned2, ", 20")

sCleaned3 = Mid(sCleaned2, lStartPosition + 6)

' BASIC DATA START

iBasicData = InStr(1, sCleaned3, "Basic Data")



sName = Mid(sCleaned3, 1, iBasicData - 1)

With MSFlexGrid1

    .Row = iRowAt
    .Col = 1
    .Text = sName

End With

' PERFORMANCE

i1WeekStart = InStr(1, sCleaned3, "1 Week")
i13WeekStart = InStr(1, sCleaned3, "13 Week")
iYTDStart = InStr(1, sCleaned3, "YTD")
i1YearStart = InStr(1, sCleaned3, "1 Year")
i3YearStart = InStr(1, sCleaned3, "3 Year")
i5YearStart = InStr(1, sCleaned3, "5 Year")
i10YearStart = InStr(1, sCleaned3, "10 Year")
iSEStart = InStr(1, sCleaned3, "Since Inception")
iExpensesStart = InStr(1, sCleaned3, "Expenses")

iNAVStart = InStr(1, sCleaned3, "NAV")
iLastStart = InStr(1, sCleaned3, "Last")
ichangeStart = InStr(1, sCleaned3, "Change")
ichangepStart = InStr(ichangeStart + 7, sCleaned3, "Change")
iAssetsStart = InStr(1, sCleaned3, "Assets")


sWeek1Return = Mid(sCleaned3, i1WeekStart + 6, (i13WeekStart - (i1WeekStart + 6))) & "%"

sWeek13Return = Mid(sCleaned3, i13WeekStart + 7, (iYTDStart - (i13WeekStart + 7))) & "%"

sYTDReturn = Mid(sCleaned3, iYTDStart + 3, (i1YearStart - (iYTDStart + 3))) & "%"

s1YearReturn = Mid(sCleaned3, i1YearStart + 6, (i3YearStart - (i1YearStart + 6))) & "%"

s3YearReturn = Mid(sCleaned3, i3YearStart + 6, (i5YearStart - (i3YearStart + 6))) & "%"

s5YearReturn = Mid(sCleaned3, i5YearStart + 6, (i10YearStart - (i5YearStart + 6))) & "%"

s10YearReturn = Mid(sCleaned3, i10YearStart + 7, (iSEStart - (i10YearStart + 7))) & "%"

sSEReturn = Mid(sCleaned3, iSEStart + 15, (iExpensesStart - (iSEStart + 15))) & "%"

sNAVReturn = Mid(sCleaned3, iNAVStart + 3, (iLastStart - (iNAVStart + 3)))

slastReturn = Mid(sCleaned3, iLastStart + 4, (ichangeStart - (iLastStart + 4)))

schangeReturn = Mid(sCleaned3, ichangeStart + 6, ((ichangepStart - 2) - (ichangeStart + 6)))

schangepReturn = Mid(sCleaned3, ichangepStart + 6, (iAssetsStart - (ichangepStart + 6)))

With MSFlexGrid1
    
    .Row = iRowAt
    .Col = 0
    .Text = sTickerSymbol

    .Row = iRowAt
    .Col = 1
    .Text = sName
    
    .Row = iRowAt
    .Col = 2
    .Text = sNAVReturn
    
    .Row = iRowAt
    .Col = 3
    .Text = slastReturn
    
    .Row = iRowAt
    .Col = 4
    .Text = "$" & schangeReturn
    
    .Row = iRowAt
    .Col = 5
    .Text = schangepReturn
    
    .Row = iRowAt
    .Col = 6
    .Text = sWeek1Return
    
    .Row = iRowAt
    .Col = 7
    .Text = sWeek13Return
    
    .Row = iRowAt
    .Col = 8
    .Text = sYTDReturn
    
    .Row = iRowAt
    .Col = 9
    .Text = s1YearReturn
    
    .Row = iRowAt
    .Col = 10
    .Text = s3YearReturn
    
    .Row = iRowAt
    .Col = 11
    .Text = s5YearReturn
    
    .Row = iRowAt
    .Col = 12
    .Text = s10YearReturn
    
    .Row = iRowAt
    .Col = 13
    .Text = sSEReturn

End With

End Sub

Public Sub ExportGrid(Grid As MSFlexGrid, FileName As String)

    Dim i As Long
    Dim j As Long
    Dim sDate As String
    
    On Error GoTo ErrHandler
    
    'Let's put a HourGlass pointer for the mouse
    Screen.MousePointer = vbHourglass
    
    Dim FileType
    
    FileType = 1
    
    If FileType = 1 Then 'Export to excel
        
        'Gimme the workbook
        Dim wkbNew As Excel.Workbook
        'Gimme the worksheet for the workbook
        Dim wkbSheet As Excel.Worksheet
        'Gimme the range for the worksheet
        Dim Rng As Excel.Range
        
        'Does the file exist?
        If Dir(FileName) <> "" Then
            'Kill it boy!
            Kill FileName
        End If
       
        On Error GoTo CreateNew_Err
        
        'Let's create the workbook kid!
        Set wkbNew = Workbooks.Add
        wkbNew.SaveAs FileName
        
        'Add a WorkPage
        Set wkbSheet = wkbNew.Worksheets(1)
        
        'MsgBox (Grid.TextMatrix(0, 0))
        'MsgBox (Grid.TextMatrix(0, 1))
                
        'Set the values in the range
        Set Rng = wkbSheet.Range("A1:" + Chr(Grid.Cols + 64) + CStr(Grid.Rows))
        'For j = 0 To Grid.Cols - 1
        For j = 1 To Grid.Cols - 1
            For i = 0 To Grid.Rows - 1
                'MsgBox Grid.TextMatrix(i, j) & "  Row = " & i & "  Col = " & j
                If j <> 0 Then
                'Rng.Range(Chr(j + 1 + 64) + CStr(i + 1)) = Grid.TextMatrix(i, j)
                ' START IN A:
                Rng.Range(Chr(j + 64) + CStr(i + 1)) = Trim(Grid.TextMatrix(i, j))
                End If
            Next
        Next
        
        'LEFT ALIGN EVERYTHING
        wkbSheet.Range("A1:" + Chr(Grid.Cols + 64) + CStr(Grid.Rows)).HorizontalAlignment = xlHAlignLeft
        
        ' WANT TO SEE THE NAME
        For j = 1 To Grid.Cols - 1
            wkbSheet.Columns(j).AutoFit
        Next
        
         ' SET UP PRINTING
         sDate = Format(Date, "dddd mmm d yyyy")
        wkbSheet.PageSetup.CenterHeader = ComboProfile.Text & " " & sDate
        wkbSheet.PageSetup.LeftMargin = 0.25
        wkbSheet.PageSetup.RightMargin = 0.25
        wkbSheet.PageSetup.PrintArea = "A1:" + Chr((Grid.Cols - 1) + 64) + CStr(Grid.Rows)
        wkbSheet.PageSetup.Orientation = xlLandscape
       
        'Close and save the file
        wkbNew.Close True
        
        GoTo NoErrors
        
CreateNew_Err:
        'Stop the show!
        wkbNew.Close False
        Set wkbNew = Nothing
        Resume ErrHandler
    
    Else 'Export to text
        
        Dim Fs As Variant
        Dim a As Variant
        
        'I know, the File # sounds smarter, but, I like weird things :) !
        On Error GoTo ErrHandler
        Set Fs = CreateObject("Scripting.FileSystemObject")
        Set a = Fs.CreateTextFile(FileName, True)
        Dim Line As String
        For j = 0 To Grid.Rows - 1
            For i = 0 To Grid.Cols - 1
                Line = Line + Grid.TextMatrix(i, j) + vbTab
            Next
            a.WriteLine (Line)
            Line = ""
        Next
        a.Close
    
    End If
    
NoErrors:
    'Gimme the default mouse pointer dude!
    Screen.MousePointer = vbDefault
    MsgBox "File has been saved", vbOKOnly, "Finished"
    Exit Sub

ErrHandler:
    'Gimme the default mouse pointer dude!
    Screen.MousePointer = vbDefault
    MsgBox "Can't export the file", vbOKOnly, "Error"
    Exit Sub
End Sub

Public Sub PopulateListNew()

Dim Symbol
Dim SymbolUbound
Dim SymbolList As String
Dim i As Integer
Dim ListGridRow As Integer
Dim ListGridColumn As Integer
Dim SymbolCount As Integer

' GET THE SYMBOL IF THERE ARE ANY

SymbolList = tickersymbol.Text
                            
Symbol = Split(SymbolList, ";")

SymbolUbound = UBound(Symbol)

If InStr(1, SymbolList, ";") = 0 Then

    MSFlexGridList.Clear
    MSFlexGridList.Rows = 2
    
    With MSFlexGridList
            .Row = 1
            .Col = 1
            .Text = SymbolList
    End With

Else
                   
    If SymbolUbound > 0 Then
    
        MSFlexGridList.Clear
        MSFlexGridList.Rows = 2
    
        ListGridRow = 1
        ListGridColumn = 1
    
        For i = 0 To SymbolUbound
          
            With MSFlexGridList
                .Row = ListGridRow
                .Col = ListGridColumn
                .Text = Symbol(i)
            End With
            
           ListGridColumn = ListGridColumn + 1
            
            If ListGridColumn = 11 Then
                If i <> 0 Then
                    ListGridRow = ListGridRow + 1
                    MSFlexGridList.Rows = MSFlexGridList.Rows + 1
                End If
             
                ListGridColumn = 1
             
            End If
       
        Next
       
    End If

End If

End Sub

Public Sub RefreshList()

MSFlexGrid1.Clear
MSFlexGridList.Clear

'BuildHeaders

If regDoes_Key_Exist(HKEY_LOCAL_MACHINE, "Software\MutualFundTracker") Then
    Dim name As String
    Dim iAt As Integer
    Dim ComboIndex As Integer
    
    ComboIndex = 0
    
    ComboProfile.Clear
    
    For iAt = 1 To 10

    name = Trim(regQuery_A_Key(HKEY_LOCAL_MACHINE, _
                           "Software\MutualFundTracker\Profile" & CStr(iAt), _
                            "Name"))
                            
         ComboProfile.AddItem name, ComboIndex
         
         ComboIndex = ComboIndex + 1
         
         ProfileArray(0, iAt) = name
         ProfileArray(1, iAt) = iAt
        
    Next
    
    BuildList
    
End If

BuildHeaders

End Sub

Public Sub BuildList()
Dim iAt As Integer
Dim name As String

name = SelectedProfile

For iAt = 1 To 10

    If ProfileArray(0, iAt) = name Then
       
        If regDoes_Key_Exist(HKEY_LOCAL_MACHINE, "Software\MutualFundTracker") Then
        
            tickersymbol.Text = regQuery_A_Key(HKEY_LOCAL_MACHINE, _
                                "Software\MutualFundTracker\Profile" & CStr(iAt), _
                                "List")
                            
            PopulateListNew
                            
        End If
         
    End If

Next
End Sub

Public Sub BuildHeaders()

With MSFlexGrid1

    .Row = 0
        
    .Col = 0
    .ColWidth(0) = 700
    .Text = "Symbol"

    .Cols = 14
    
    .Col = 1
    .ColWidth(1) = 600
    .Text = "Name"

    .Col = 2
    .ColWidth(2) = 700
    .Text = "Current"
    
    .Col = 3
    .ColWidth(3) = 700
    .Text = "Previous"
    
    .Col = 4
    .ColWidth(4) = 700
    .Text = "Change"
    
    .Col = 5
    .ColWidth(5) = 850
    .Text = "% Change"
    
    .Col = 6
    .ColWidth(6) = 700
    .Text = "1 Week"
    
    .Col = 7
    .ColWidth(7) = 700
    .Text = "13 Week"
    
    .Col = 8
    .ColWidth(8) = 500
    .Text = "YTD"
    
    .Col = 9
    .ColWidth(9) = 600
    .Text = "1 Year"
    
    .Col = 10
    .ColWidth(10) = 600
    .Text = "3 Year"
    
    .Col = 11
    .ColWidth(11) = 600
    .Text = "5 Year"
    
    .Col = 12
    .ColWidth(12) = 600
    .Text = "10 Year"
    
    .Col = 13
    .ColWidth(13) = 600
    .Text = "Life"

End With

With MSFlexGridList

    .Row = 0
    .RowHeight(0) = 0

    .Col = 0
    .ColWidth(0) = 0
    .Text = "List"

    .Cols = 11

    .Col = 1
    .ColWidth(1) = 700
    .Text = "List"

    .Col = 2
    .ColWidth(2) = 700
    .Text = "List"

    .Col = 3
    .ColWidth(3) = 700
    .Text = "List"

    .Col = 4
    .ColWidth(4) = 700
    .Text = "List"

    .Col = 5
    .ColWidth(5) = 700
    .Text = "List"

    .Col = 5
    .ColWidth(5) = 700
    .Text = "List"

    .Col = 6
    .ColWidth(6) = 700
    .Text = "List6"

    .Col = 7
    .ColWidth(7) = 700
    .Text = "List7"

    .Col = 8
    .ColWidth(8) = 700
    .Text = "List8"

    .Col = 9
    .ColWidth(9) = 700
    .Text = "List9"
    
    .Col = 10
    .ColWidth(10) = 700
    .Text = "List10"


End With

' GET THE LAST SELECTED
ComboProfile.Text = regQuery_A_Key(HKEY_LOCAL_MACHINE, _
                                    "Software\MutualFundTracker", _
                                    "LastSelected")

End Sub

Public Sub CreateRegistryKeys()

' BUILD EVERYTHING THE FIRST TIME THE APP IS RUN
If Not regDoes_Key_Exist(HKEY_LOCAL_MACHINE, "Software\MutualFundTracker") Then
        
    regCreate_A_Key HKEY_LOCAL_MACHINE, "Software\MutualFundTracker"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker", "LastSelected", "Demo Profile"
    
    regCreate_A_Key HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile1"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile1", "Name", "Demo Profile"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile1", "List", "AVLFX;OENBX"
    
    regCreate_A_Key HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile2"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile2", "Name", "Open Profile 1"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile2", "List", ""
    
    regCreate_A_Key HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile3"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile3", "Name", "Open Profile 2"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile3", "List", ""
    
    regCreate_A_Key HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile4"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile4", "Name", "Open Profile 3"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile4", "List", ""
    
    regCreate_A_Key HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile5"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile5", "Name", "Open Profile 4"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile5", "List", ""
    
    regCreate_A_Key HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile6"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile6", "Name", "Open Profile 5"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile6", "List", ""
    
    regCreate_A_Key HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile7"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile7", "Name", "Open Profile 6"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile7", "List", ""
    
    regCreate_A_Key HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile8"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile8", "Name", "Open Profile 7"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile8", "List", ""
    
    regCreate_A_Key HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile9"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile9", "Name", "Open Profile 8"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile9", "List", ""
    
    regCreate_A_Key HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile10"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile10", "Name", "Open Profile 9"
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile10", "List", ""
    
End If

End Sub

Public Function GetSelectedProfile() As String
GetSelectedProfile = SelectedProfile
End Function

Public Sub SetSelectedProfile(SetProfile As String)
SelectedProfile = SetProfile
End Sub

