VERSION 5.00
Begin VB.Form AddSymbols 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Symbols"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3900
   Icon            =   "AddSymbols.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   3900
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox ProfileName 
      Height          =   285
      Left            =   720
      TabIndex        =   11
      Top             =   720
      Width           =   2775
   End
   Begin VB.ComboBox ComboProfile 
      Height          =   315
      Left            =   720
      TabIndex        =   9
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cndApply 
      Caption         =   "Apply"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtAdd 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   3495
      Begin VB.Label Label1 
         Caption         =   "Current List"
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4455
      Left            =   0
      TabIndex        =   10
      Top             =   480
      Width           =   3855
      Begin VB.Label Label3 
         Caption         =   "Name:"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   260
         Width           =   615
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Profiles:"
      Height          =   255
      Left            =   155
      TabIndex        =   12
      Top             =   150
      Width           =   615
   End
End
Attribute VB_Name = "AddSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ProfileArray(2, 11) As String
Dim bChanged As Boolean
Dim sCurrentProfile As String

Private Sub ComboProfile_Click()

If bChanged = True Then
    SaveList
End If

ProfileName.Text = ComboProfile.Text
sCurrentProfile = ComboProfile.Text
GetList

End Sub

Private Sub cmdAdd_Click()
If Len(Trim(txtAdd.Text)) > 0 Then
List1.AddItem UCase(txtAdd.Text)
txtAdd.Text = ""
bChanged = True
Else
 MsgBox "Enter a Mutual Fund symbol."
End If
End Sub

Private Sub cmdClose_Click()
AddSymbols.Hide
Set AddSymbols = Nothing
End Sub

Private Sub cmdDelete_Click()
Dim ListBoxCount As Integer
Dim iRowAt As Integer
Dim iSelectedRow  As Integer


iSelectedRow = 10000

ListBoxCount = List1.ListCount - 1

For iRowAt = 0 To ListBoxCount
    bChanged = True
    If List1.Selected(iRowAt) Then
        iSelectedRow = iRowAt
        List1.RemoveItem (iRowAt)
        Exit For
    End If
Next

If iSelectedRow = 10000 Then
    MsgBox "Please select a Mutual Fund to delete."
End If
End Sub

Private Sub cndApply_Click()

SaveList
InvestmentsForm.RefreshList

End Sub

Private Sub Form_Load()

Dim name As String
Dim iAt As Integer
Dim ComboIndex As Integer

' LOAD THE PROFILES
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
    
    ' SET THE PROFILE
    sCurrentProfile = InvestmentsForm.GetSelectedProfile()
    ComboProfile.Text = sCurrentProfile
    ProfileName.Text = sCurrentProfile
    
    ' ADDED 1/30/2001
    ComboProfile_Click
    
End If

End Sub

Public Sub GetList()

Dim Symbol
Dim SymbolUbound
Dim SymbolList As String
Dim i As Integer
Dim name As String

name = ComboProfile.Text

List1.Clear

For iAt = 1 To 10

    If ProfileArray(0, iAt) = name Then
       
        If regDoes_Key_Exist(HKEY_LOCAL_MACHINE, "Software\MutualFundTracker") Then
        
           SymbolList = regQuery_A_Key(HKEY_LOCAL_MACHINE, _
                                "Software\MutualFundTracker\Profile" & CStr(iAt), _
                                "List")
                            
             Symbol = Split(SymbolList, ";")
           
            SymbolUbound = UBound(Symbol)
            
            If SymbolUbound <> -1 Then
            
                If InStr(1, SymbolList, ";") = 0 Then
 
                    List1.AddItem SymbolList
                Else
                                   
                    If SymbolUbound > 0 Then
                    
                        For i = 0 To SymbolUbound
                        List1.AddItem Symbol(i)
                        Next
                       
                    End If
                
                End If
            
            End If ' SYMBOLUBOUND <> -1
                             
        End If 'If regDoes_Key_Exist(HKEY_LOCAL_MACHINE,
         
    End If 'If ProfileArray(0, iAt) = name

Next

End Sub

Public Sub SaveList()

Dim ListBoxCount As Integer
Dim iRowAt As Integer
Dim iSelectedRow  As Integer
Dim SaveList As String
Dim iAt As Integer
Dim iDupAt As Integer
Dim name As String
Dim ibail As Integer
Dim sSelectedProfile As String

SaveList = ""

ListBoxCount = List1.ListCount - 1

For iRowAt = 0 To ListBoxCount

    If iRowAt <> 0 Then
        SaveList = SaveList & ";"
    End If
    List1.Selected(iRowAt) = True
    SaveList = SaveList & List1.Text
   
Next

ibail = 0

name = sCurrentProfile

For iAt = 1 To 10

    If ProfileArray(0, iAt) = name Then
    
        For iDupAt = 1 To 10
            If ProfileName.Text = ProfileArray(0, iDupAt) And iAt <> iDupAt Then
                MsgBox "You cannot have two profiles of the same name."
                ibail = 1
            End If
        Next
       
       If ibail = 0 Then
       
            sSelectedProfile = InvestmentsForm.GetSelectedProfile()
       
            ' UPDATE REGISTRY LAST SELECTED IF CHANGING ITS' NAME
            If sCurrentProfile = sSelectedProfile Then
                regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker", "LastSelected", ProfileName.Text
                InvestmentsForm.SetSelectedProfile ProfileName.Text
            End If
       
            If regDoes_Key_Exist(HKEY_LOCAL_MACHINE, "Software\MutualFundTracker") Then
            
                regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile" & ProfileArray(1, iAt), "List", SaveList
                regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\MutualFundTracker\Profile" & ProfileArray(1, iAt), "Name", ProfileName.Text
                            
            End If
        
      End If 'If ibail = 0
         
    End If 'If ProfileArray(0, iAt)

Next

End Sub



Private Sub ProfileName_LostFocus()
bChanged = True
End Sub
