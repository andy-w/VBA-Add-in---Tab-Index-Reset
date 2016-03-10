VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tab Index Reset"
   ClientHeight    =   1080
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Reset Form Tab Index order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   735
      TabIndex        =   2
      Top             =   315
      Width           =   2985
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBE
Public Connect As Connect

Option Explicit

Private Sub CancelButton_Click()
    Connect.Hide
End Sub

Private Sub OKButton_Click()
    
    Dim pVBAProject As VBProject
    Dim vbComp As VBComponent  'VBA module, form, etc...
    Dim oControls As MSForms.Controls
    Dim oUserForm As MSForms.UserForm
    
    Set pVBAProject = VBInstance.ActiveVBProject
  
    'Loop through all the components (modules, forms, etc) in the VBA project
    
    Set vbComp = VBInstance.SelectedVBComponent
    
    'For Each vbComp In pVBAProject.VBComponents
    
        If vbComp.Type = vbext_ct_MSForm Then
            
            Set oUserForm = vbComp.Designer
            
            If Not oUserForm Is Nothing Then
                
                On Error Resume Next
                
                Set oControls = oUserForm.Controls
                
                Call sortControls(oControls, False)
                                
                On Error GoTo 0
                    
            End If
            
        End If
          
    'Next
           
    Connect.Hide
    
End Sub

Private Sub sortControls( _
    ByRef oControls As MSForms.Controls, _
    ByVal boolRightToLeft As Boolean)
    
    Dim i As Integer
    Dim oControl1 As MSForms.Control
    Dim oControl2 As MSForms.Control
    Dim j As Integer
    Dim colSortPos As Collection
    Dim boolDone As Boolean
    Dim oContainControls As MSForms.Controls
    
    Set colSortPos = New Collection
                                     
    For i = 1 To oControls.Count
                
        boolDone = False
                    
        Set oControl1 = oControls(i - 1)
                    
        For j = 1 To colSortPos.Count
                                  
            Set oControl2 = colSortPos(j)
                        
            If compareControlForTabIndex(oControl1, oControl2, boolRightToLeft) < 0 Then
                            
                Call colSortPos.Add(oControl1, , j)
                                    
                boolDone = True
                                    
                Exit For
                            
            End If
                                
        Next
                            
        If Not boolDone Then
            Call colSortPos.Add(oControl1)
        End If
                            
        Set oContainControls = Nothing
        
        On Error Resume Next
        
        Set oContainControls = oControl1.Controls
        
        If Not oContainControls Is Nothing Then
        
            Call sortControls(oContainControls, boolRightToLeft)
            
        End If
        
        On Error GoTo 0
        
    Next
                
    j = 0
    For Each oControl1 In colSortPos
        oControl1.TabIndex = j
        j = j + 1
    Next
    
    
End Sub
    
Private Function compareControlForTabIndex( _
    ByRef objControl1 As MSForms.Control, _
    ByRef objControl2 As MSForms.Control, _
    ByVal bRightToLeft As Boolean) As Long
   
    ' Returns:
    ' -1 if Control1 has lower TabIndex
    ' +1 if Control1 has higher TabIndex
    ' 0 if Control1 and Control2 are the same

    Dim lResult  As Long
   
    Dim lTopControl1 As Long
    Dim lHeightControl1 As Long
    Dim lLeftControl1 As Long
    Dim lWidthControl1 As Long
    Dim sNameControl1 As String
    Dim lIndexControl1 As Long
   
    Dim lTopControl2 As Long
    Dim lHeightControl2 As Long
    Dim lLeftControl2 As Long
    Dim lWidthControl2 As Long
    Dim sNameControl2 As String
    Dim lIndexControl2 As Long
   
    ' By default
    lResult = -1
      
    ' There are 6 relative positions in vertical:
    ' 1) Top2 + Height2 < Top1
    ' 2) Top2 > Top1 + Height1
    ' 3) Top2 < Top1, Top1 < Top2 + Height2 < Top1 + Height1
    ' 4) Top2 > Top1, Top1 < Top2 + Height2 < Top1 + Height1
    ' 5) Top2 > Top1, Top2 + Height2 > Top1 + Height1
    ' 6) Top2 < Top1, Top2 + Height2 > Top1 + Height1
   
    With objControl1
        sNameControl1 = .Name
        lTopControl1 = .top
        lHeightControl1 = .height
        lLeftControl1 = .left
        lWidthControl1 = .width
    End With
    
    With objControl2
        sNameControl2 = .Name
        lTopControl2 = .top
        lHeightControl2 = .height
        lLeftControl2 = .left
        lWidthControl2 = .width
    End With
         
    If lTopControl2 + lHeightControl2 <= lTopControl1 Then

         ' Control 2 is completely above Control 1. It must have lower TabIndex
         lResult = 1

    ElseIf lTopControl1 + lHeightControl1 <= lTopControl2 Then

         ' Control 2 is completely below Control 1.

    Else
     
        If (lTopControl2 < lTopControl1) And (lTopControl2 + lHeightControl2) <= (lTopControl1 + lHeightControl1) Then
        
            ' Control 2 starts above control 1 but doesn't ends below Control 1. It must have lower TabIndex
            lResult = 1
        
        ElseIf (lTopControl1 < lTopControl2) And (lTopControl1 + lHeightControl1) <= (lTopControl2 + lHeightControl2) Then
        
            ' Control 1 starts above control 2 but doesn't end below control 2.

        Else

            ' There are 2 remaining positions with controls 1 and 2 overlapping vertically
            If lLeftControl2 + lWidthControl2 <= lLeftControl1 Then

               ' Control 2 is completely to the left of control 1. It must have lower TabIndex
               If Not bRightToLeft Then
                  lResult = 1
               Else
                  lResult = -1
               End If
                  
            ElseIf lLeftControl1 + lWidthControl1 <= lLeftControl2 Then

               ' Control 2 is completely to the right of control 1.
               If Not bRightToLeft Then
                  lResult = -1
               Else
                  lResult = 1

               End If
            Else

               ' Controls 1 y 2 overlap also horizontally.
               If lTopControl2 < lTopControl1 Then
                  lResult = 1
               End If

            End If

        End If

    End If
   
    compareControlForTabIndex = lResult
   
End Function
