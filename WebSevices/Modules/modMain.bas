Attribute VB_Name = "modMain"
Option Explicit

Global Const LISTVIEW_MODE0 = "View Large Icons"
Global Const LISTVIEW_MODE1 = "View Small Icons"
Global Const LISTVIEW_MODE2 = "View List"
Global Const LISTVIEW_MODE3 = "View Details"
Public fMainForm As frmMain

Private aJoin() As String

Sub Main()
    frmSplash.Show
    frmSplash.Refresh
    Set fMainForm = New frmMain
    Load fMainForm
    Get_All_Countries fMainForm, False, , True
    DoEvents
    
    Unload frmSplash
    fMainForm.Show
End Sub

Public Function AddNode(ByRef InForm As Form) As Boolean
Dim nodX As Node
    
    With InForm.tvTreeView
        .Nodes.Clear
        Set nodX = .Nodes.Add(, , "B", "Business And Commerce", 1, 1)
        Set nodX = .Nodes.Add("B", tvwChild, "B1", "Validate Credit Card", 3, 3)
        Set nodX = .Nodes.Add("B", tvwChild, "B2", "Mortgage Web Service", 6, 6)
        Set nodX = .Nodes.Add("B", tvwChild, "B3", "Currency Converter", 2, 2)
        nodX.EnsureVisible
        Set nodX = .Nodes.Add(, , "S", "Standards And Lookup Data", 4, 4)
        Set nodX = .Nodes.Add("S", tvwChild, "S1", "Get Country By Country Code", 9, 9)
        Set nodX = .Nodes.Add("S", tvwChild, "S2", "Get International Dialing Code", 9, 9)
        Set nodX = .Nodes.Add("S", tvwChild, "S3", "Get All Countries With ISO", 9, 9)
        Set nodX = .Nodes.Add("S", tvwChild, "S4", "Get Currency Code", 9, 9)
        nodX.EnsureVisible
        Set nodX = .Nodes.Add(, , "V", "Value Manipulation/Unit Convertor", 7, 7)
        Set nodX = .Nodes.Add("V", tvwChild, "V1", "Speed Unit Convertor", 7, 7)
        Set nodX = .Nodes.Add("V", tvwChild, "V2", "Weight Unit Convertor", 7, 7)
        Set nodX = .Nodes.Add("V", tvwChild, "V3", "Length/Distance Convertor", 7, 7)
        Set nodX = .Nodes.Add("V", tvwChild, "V4", "Area Unit Convertor", 7, 7)
        Set nodX = .Nodes.Add("V", tvwChild, "V5", "Metric Weight Unit Convertor", 7, 7)
        Set nodX = .Nodes.Add("V", tvwChild, "V6", "Torque Unit Convertor", 7, 7)
        Set nodX = .Nodes.Add("V", tvwChild, "V7", "Power Unit Convertor", 7, 7)
        Set nodX = .Nodes.Add("V", tvwChild, "V8", "Acceleration Unit Convertor", 7, 7)
        nodX.EnsureVisible
        Set nodX = .Nodes.Add(, , "C", "Communications", 15, 15)
        Set nodX = .Nodes.Add("C", tvwChild, "C1", "Send SMS World", 10, 10)
        Set nodX = .Nodes.Add("C", tvwChild, "C2", "Validate Email Address", 14, 14)
        Set nodX = .Nodes.Add("C", tvwChild, "C3", "Send Fax", 18, 18)
        nodX.EnsureVisible
        Set nodX = .Nodes.Add(, , "G", "Create Bar Codes", 5, 5)
        Set nodX = .Nodes.Add("G", tvwChild, "G1", "Bar Code Generator", 5, 5)
        Set nodX = .Nodes.Add("G", tvwChild, "G2", "Code 39 Bar Code", 5, 5)
        nodX.EnsureVisible
        Set nodX = .Nodes.Add(, , "U", "Utilities", 17, 17)
        Set nodX = .Nodes.Add("U", tvwChild, "U1", "Global Weather", 16, 16)
        nodX.EnsureVisible
        '.Style = tvwPlusMinusText
        .BorderStyle = vbFixedSingle
    End With
    
    Set nodX = Nothing
End Function

Public Function Setup_Form(ByRef InForm As Form) As Boolean
Dim StrNode As String
Dim IsOk As Boolean
    
    InForm.fraMain.Visible = False
    
    IsOk = Unload_Controls(InForm)
    If IsOk Then
        With InForm
            'Me.img(0).Picture = Me.ImageList1.ListImages(1).Picture
            StrNode = .tvTreeView.SelectedItem.Key
            Select Case StrNode
                Case "B", "S", "V", "C", "G", "U"
                    IsOk = Setup_Parents(InForm)
                Case "B1"
                    IsOk = Validate_Credit_Card(InForm)
                Case "B2"
                    IsOk = Mortgage_Payment(InForm)
                Case "B3"
                    IsOk = Currency_Converter(InForm)
                Case "S1"
                    IsOk = GetCountryByCountryCode(InForm)
                Case "S2"
                    IsOk = GetInternationDialingCodes(InForm)
                Case "S3"
                    IsOk = GetCountries(InForm)
                Case "S4"
                    IsOk = GetInternationDialingCodes(InForm)
                Case "V1", "V2", "V3", "V4", "V5", "V6", "V7", "V8"
                    IsOk = Unit_Converter(InForm, StrNode)
                Case "C1"
                    IsOk = SendSMSWorld(InForm)
                Case "C2"
                    IsOk = ValidateEmail(InForm)
                Case "C3"
                    IsOk = SendFax(InForm)
                Case "G1"
                Case "G2"
                Case "U1"
            End Select
        End With
    End If
    InForm.fraMain.Visible = True
    
End Function

Private Function Unload_Controls(ByRef InForm As Form) As Boolean
Dim i As Long
    
    'Remove the text boxes
    If InForm.txt.UBound >= 1 Then
        For i = InForm.txt.UBound To 1 Step -1
            Unload InForm.txt(i)
        Next i
    End If
    InForm.txt(0).Text = vbNullString
    InForm.txt(0).Visible = False
    
    'remove the combo boxes
    If InForm.cbo.UBound >= 1 Then
        For i = InForm.cbo.UBound To 1 Step -1
            Unload InForm.cbo(i)
        Next i
    End If
    InForm.cbo(0).Clear
    InForm.cbo(0).Visible = False
    
    'remove the labels
    If InForm.lblName.UBound >= 1 Then
        For i = InForm.lblName.UBound To 1 Step -1
            Unload InForm.lblName(i)
        Next i
    End If
    InForm.lblName(0).Caption = vbNullString
    InForm.lblName(0).Visible = False
    
    If InForm.txtResults.UBound >= 1 Then
        For i = InForm.txtResults.UBound To 1 Step -1
            Unload InForm.txtResults(i)
        Next i
    End If
    InForm.txtResults(0).Text = vbNullString
    InForm.txtResults(0).Visible = False
    
    If InForm.img.UBound >= 1 Then
        For i = InForm.img.UBound To 1 Step -1
            Unload InForm.img(i)
        Next i
    End If
    InForm.img(0).Visible = False
    
    InForm.cmdRun.Visible = False
    InForm.cmdClear.Visible = False
    
ExitFunction:
    Unload_Controls = True
    Exit Function
    
End Function

Private Function Setup_Parents(ByRef InForm As Form) As Boolean
Dim i As Long
Dim strC As String
Dim N As Integer
Dim iImgIndex As Integer
        
    With InForm.tvTreeView
        If .SelectedItem.Children > 0 Then 'has children
            strC = .SelectedItem.Child.Text
            N = .SelectedItem.Child.Index
            iImgIndex = .SelectedItem.Child.Image
            
            InForm.img(0).Picture = InForm.ImageList1.ListImages(iImgIndex).Picture
            InForm.img(0).Left = 120
            InForm.img(0).Top = 300
            InForm.img(0).Visible = True
                        
            InForm.lblName(0).Caption = strC
            InForm.lblName(0).Top = InForm.img(0).Top
            InForm.lblName(0).Left = InForm.img(0).Left + InForm.img(0).Width + 100
            InForm.lblName(0).Visible = True
            
            i = 1
            While N <> .SelectedItem.Child.LastSibling.Index
                strC = .Nodes(N).Next.Text
                iImgIndex = .Nodes(N).Next.Image
                N = .Nodes(N).Next.Index
                                
                Load InForm.img(i)
                InForm.img(i).Picture = InForm.ImageList1.ListImages(iImgIndex).Picture
                InForm.img(i).Left = InForm.img(0).Left
                InForm.img(i).Top = InForm.img(i - 1).Top + InForm.img(i - 1).Height + 100
                InForm.img(i).Visible = True
                
                Load InForm.lblName(i)
                InForm.lblName(i).Caption = strC
                InForm.lblName(i).Top = InForm.img(i).Top
                InForm.lblName(i).Left = InForm.lblName(0).Left
                InForm.lblName(i).Visible = True
                i = i + 1
            Wend
            
        End If
    End With
ExitFunction:
    i = 0
    strC = vbNullString
    N = 0
    iImgIndex = 0
    
    Setup_Parents = True
    Exit Function
    
End Function

Private Function Validate_Credit_Card(ByRef InForm As Form) As Boolean
    With InForm
        .lblName(0).Visible = True
        .lblName(0).Caption = "Card Type:"
        .lblName(0).Top = 300
        .lblName(0).Left = 120
        .cbo(0).Clear
        .cbo(0).AddItem "VISA"
        .cbo(0).AddItem "MASTERCARD"
        .cbo(0).AddItem "DINERS"
        .cbo(0).AddItem "AMEX"
        .cbo(0).Text = "VISA"
        .cbo(0).Top = 300
        .cbo(0).Left = 2000
        .cbo(0).Width = 1500
        .cbo(0).Visible = True
        .txt(0).Text = vbNullString
        .txt(0).Left = .cbo(0).Left
        .txt(0).Top = .cbo(0).Top + .cbo(0).Height + 200
        .txt(0).MaxLength = 50
        .txt(0).Width = 4000
        .txt(0).Visible = True
        Load .lblName(1)
        .lblName(1).Visible = True
        .lblName(1).Caption = "Credit Card #:"
        .lblName(1).Top = .txt(0).Top
        .lblName(1).Left = .lblName(0).Left
                            
        .cmdClear.Visible = True
        .cmdClear.Left = .lblName(0).Left
        .cmdClear.Top = .txt(0).Top + .txt(0).Height + 200
        .cmdRun.Visible = True
        .cmdRun.Left = .txt(0).Left
        .cmdRun.Top = .cmdClear.Top
        
        .cbo(0).TabIndex = 0
        .txt(0).TabIndex = 1
        .cmdRun.TabIndex = 2
        
        .txtResults(0).Text = vbNullString
        .txtResults(0).Left = .txt(0).Left
        .txtResults(0).Top = .cmdRun.Top + .cmdRun.Height + 200
        .txtResults(0).TabStop = False
        .txtResults(0).Width = .fraMain.Width - 3000
        .txtResults(0).Height = 2000
        .txtResults(0).Visible = True
        
        Load .lblName(2)
        .lblName(2).Visible = True
        .lblName(2).Caption = "Results:"
        .lblName(2).Top = .txtResults(0).Top
        .lblName(2).Left = .lblName(0).Left
    End With
ExitFunction:
    Validate_Credit_Card = True
    Exit Function
End Function

Public Function Mortgage_Payment(ByRef InForm As Form) As Boolean
    With InForm
        .lblName(0).Visible = True
        .lblName(0).Caption = "Years:"
        .lblName(0).Top = 300
        .lblName(0).Left = 120
        .txt(0).Text = vbNullString
        .txt(0).Left = 2000
        .txt(0).Top = 300
        .txt(0).Width = 500
        .txt(0).MaxLength = 4
        .txt(0).Visible = True
        Load .lblName(1)
        .lblName(1).Visible = True
        .lblName(1).Caption = "Interest:"
        .lblName(1).Top = .txt(0).Top + .txt(0).Height + 200
        .lblName(1).Left = .lblName(0).Left
        Load .txt(1)
        .txt(1).Text = vbNullString
        .txt(1).Left = .txt(0).Left
        .txt(1).Top = .lblName(1).Top
        .txt(1).Width = 1000
        .txt(1).MaxLength = 6
        .txt(1).Visible = True
        Load .lblName(2)
        .lblName(2).Visible = True
        .lblName(2).Caption = "Loan Amount:"
        .lblName(2).Top = .txt(1).Top + .txt(1).Height + 200
        .lblName(2).Left = .lblName(0).Left
        Load .txt(2)
        .txt(2).Text = vbNullString
        .txt(2).Left = .txt(0).Left
        .txt(2).Top = .lblName(2).Top
        .txt(2).Width = 2000
        .txt(2).MaxLength = 20
        .txt(2).Visible = True
        Load .lblName(3)
        .lblName(3).Visible = True
        .lblName(3).Caption = "Annual Tax:"
        .lblName(3).Top = .txt(2).Top + .txt(2).Height + 200
        .lblName(3).Left = .lblName(0).Left
        Load .txt(3)
        .txt(3).Text = vbNullString
        .txt(3).Left = .txt(0).Left
        .txt(3).Top = .lblName(3).Top
        .txt(3).Width = 2000
        .txt(3).MaxLength = 20
        .txt(3).Visible = True
        Load .lblName(4)
        .lblName(4).Visible = True
        .lblName(4).Caption = "Annual Insurance:"
        .lblName(4).Top = .txt(3).Top + .txt(3).Height + 200
        .lblName(4).Left = .lblName(0).Left
        Load .txt(4)
        .txt(4).Text = vbNullString
        .txt(4).Left = .txt(0).Left
        .txt(4).Top = .lblName(4).Top
        .txt(4).Width = 2000
        .txt(4).MaxLength = 20
        .txt(4).Visible = True
                            
        .cmdClear.Visible = True
        .cmdClear.Left = .lblName(0).Left
        .cmdClear.Top = .txt(4).Top + .txt(4).Height + 200
        .cmdRun.Visible = True
        .cmdRun.Left = .txt(0).Left
        .cmdRun.Top = .cmdClear.Top
        
        .txt(0).TabIndex = 0
        .txt(1).TabIndex = 1
        .txt(2).TabIndex = 2
        .txt(3).TabIndex = 3
        .txt(4).TabIndex = 4
        .cmdRun.TabIndex = 5
        
        .txtResults(0).Text = vbNullString
        .txtResults(0).Left = .txt(0).Left
        .txtResults(0).Top = .cmdRun.Top + .cmdRun.Height + 200
        .txtResults(0).TabStop = False
        .txtResults(0).Width = .fraMain.Width - 3000
        .txtResults(0).Height = 2000
        .txtResults(0).Visible = True
        
        Load .lblName(5)
        .lblName(5).Visible = True
        .lblName(5).Caption = "Results:"
        .lblName(5).Top = .txtResults(0).Top
        .lblName(5).Left = .lblName(0).Left
    End With
ExitFunction:
    Mortgage_Payment = True
    Exit Function
End Function

Public Function Currency_Converter(ByRef InForm As Form) As Boolean
    With InForm
        .lblName(0).Visible = True
        .lblName(0).Caption = "From Currency:"
        .lblName(0).Top = 300
        .lblName(0).Left = 120
        
        Load_Combos .cbo(0), "B3"
        With .cbo(0)
            .Left = 2000
            .Top = 300
            .Width = 3000
            .Visible = True
        End With
        Load .lblName(1)
        .lblName(1).Visible = True
        .lblName(1).Caption = "To Currency:"
        .lblName(1).Top = .cbo(0).Top + .cbo(0).Height + 200
        .lblName(1).Left = .lblName(0).Left
                            
        Load .cbo(1)
        Load_Combos .cbo(1), "B3"
        With .cbo(1)
            .Left = InForm.cbo(0).Left
            .Top = InForm.lblName(1).Top
            .Width = 3000
            .Visible = True
        End With
        
        .cmdClear.Visible = True
        .cmdClear.Left = .lblName(0).Left
        .cmdClear.Top = .cbo(1).Top + .cbo(1).Height + 200
        .cmdRun.Visible = True
        .cmdRun.Left = .cbo(0).Left
        .cmdRun.Top = .cmdClear.Top
        
        .cbo(0).TabIndex = 0
        .cbo(1).TabIndex = 1
        .cmdRun.TabIndex = 3
        
        .txtResults(0).Text = vbNullString
        .txtResults(0).Left = .cbo(0).Left
        .txtResults(0).Top = .cmdRun.Top + .cmdRun.Height + 200
        .txtResults(0).TabStop = False
        .txtResults(0).Width = .fraMain.Width - 3000
        .txtResults(0).Height = 2000
        .txtResults(0).Visible = True
        
        Load .lblName(2)
        .lblName(2).Visible = True
        .lblName(2).Caption = "Results:"
        .lblName(2).Top = .txtResults(0).Top
        .lblName(2).Left = .lblName(0).Left
    End With
ExitFunction:
    Currency_Converter = True
    Exit Function
End Function

Private Function GetCountryByCountryCode(ByRef InForm As Form) As Boolean
    With InForm
        .lblName(0).Visible = True
        .lblName(0).Caption = "Country Code:"
        .lblName(0).Top = 300
        .lblName(0).Left = 120
        .txt(0).Text = vbNullString
        .txt(0).Left = 2000
        .txt(0).Top = 300
        .txt(0).MaxLength = 2
        .txt(0).Width = 350
        .txt(0).Visible = True
                            
        .cmdClear.Visible = True
        .cmdClear.Left = .lblName(0).Left
        .cmdClear.Top = .txt(0).Top + .txt(0).Height + 200
        .cmdRun.Visible = True
        .cmdRun.Left = .txt(0).Left
        .cmdRun.Top = .cmdClear.Top
        
        .txt(0).TabIndex = 0
        .cmdRun.TabIndex = 1
        
        .txtResults(0).Text = vbNullString
        .txtResults(0).Left = .txt(0).Left
        .txtResults(0).Top = .cmdRun.Top + .cmdRun.Height + 200
        .txtResults(0).TabStop = False
        .txtResults(0).Width = .fraMain.Width - 3000
        .txtResults(0).Height = 2000
        .txtResults(0).Visible = True
        
        Load .lblName(1)
        .lblName(1).Visible = True
        .lblName(1).Caption = "Results:"
        .lblName(1).Top = .txtResults(0).Top
        .lblName(1).Left = .lblName(0).Left
    End With
ExitFunction:
    GetCountryByCountryCode = True
    Exit Function
End Function

Private Function GetInternationDialingCodes(ByRef InForm As Form) As Boolean
    With InForm
        .lblName(0).Visible = True
        .lblName(0).Caption = "Country:"
        .lblName(0).Top = 300
        .lblName(0).Left = 120
        .cbo(0).Left = 2000
        .cbo(0).Top = 300
        .cbo(0).Width = 3000
        .cbo(0).Visible = True
        
        Get_All_Countries InForm, True, .cbo(0), False
                            
        .cmdClear.Visible = True
        .cmdClear.Left = .lblName(0).Left
        .cmdClear.Top = .cbo(0).Top + .cbo(0).Height + 200
        .cmdRun.Visible = True
        .cmdRun.Left = .cbo(0).Left
        .cmdRun.Top = .cmdClear.Top
        
        .cbo(0).TabIndex = 0
        .cmdRun.TabIndex = 1
        
        .txtResults(0).Text = vbNullString
        .txtResults(0).Left = .cbo(0).Left
        .txtResults(0).Top = .cmdRun.Top + .cmdRun.Height + 200
        .txtResults(0).TabStop = False
        .txtResults(0).Width = .fraMain.Width - 3000
        .txtResults(0).Height = 2000
        .txtResults(0).Visible = True
        On Error Resume Next
        Load .lblName(1)
        .lblName(1).Visible = True
        .lblName(1).Caption = "Results:"
        .lblName(1).Top = .txtResults(0).Top
        .lblName(1).Left = .lblName(0).Left
    End With
ExitFunction:
    GetInternationDialingCodes = True
    Exit Function
End Function

Private Function GetCountries(ByRef InForm As Form) As Boolean
    With InForm
        .cmdRun.Visible = True
        .cmdRun.Left = 120
        .cmdRun.Top = 300
        
        .cmdRun.TabIndex = 0
        
        .txtResults(0).Text = vbNullString
        .txtResults(0).Left = 2000
        .txtResults(0).Top = .cmdRun.Top + .cmdRun.Height + 200
        .txtResults(0).TabStop = False
        .txtResults(0).Width = .fraMain.Width - 3000
        .txtResults(0).Height = .fraMain.Height - 1000
        .txtResults(0).Visible = True
        
        .lblName(0).Visible = True
        .lblName(0).Caption = "Results:"
        .lblName(0).Top = .txtResults(0).Top
        .lblName(0).Left = 120
    End With
ExitFunction:
    GetCountries = True
    Exit Function
End Function

Public Function Unit_Converter(ByRef InForm As Form, ByVal InType As String) As Boolean
    With InForm
        .lblName(0).Visible = True
        .lblName(0).Caption = "Value:"
        .lblName(0).Top = 300
        .lblName(0).Left = 120
        .txt(0).Text = vbNullString
        .txt(0).Left = 2000
        .txt(0).Top = 300
        .txt(0).MaxLength = 10
        .txt(0).Width = 2000
        .txt(0).Visible = True
        
        Load_Combos .cbo(0), InType
        With .cbo(0)
            .Left = 2000
            .Top = InForm.txt(0).Top + InForm.txt(0).Height + 200
            .Width = 4000
            .Visible = True
        End With
        Load .lblName(1)
        .lblName(1).Visible = True
        .lblName(1).Caption = "From Unit:"
        .lblName(1).Top = .cbo(0).Top
        .lblName(1).Left = .lblName(0).Left
                            
        Load .cbo(1)
        Load_Combos .cbo(1), InType
        With .cbo(1)
            .Left = InForm.cbo(0).Left
            .Top = InForm.cbo(0).Top + InForm.cbo(0).Height + 200
            .Width = 4000
            .Visible = True
        End With
        Load .lblName(2)
        .lblName(2).Visible = True
        .lblName(2).Caption = "To Unit:"
        .lblName(2).Top = .cbo(1).Top
        .lblName(2).Left = .lblName(0).Left
        
        
        .cmdClear.Visible = True
        .cmdClear.Left = .lblName(0).Left
        .cmdClear.Top = .cbo(1).Top + .cbo(1).Height + 200
        .cmdRun.Visible = True
        .cmdRun.Left = .cbo(0).Left
        .cmdRun.Top = .cmdClear.Top
        
        .txt(0).TabIndex = 0
        .cbo(0).TabIndex = 1
        .cbo(1).TabIndex = 2
        .cmdRun.TabIndex = 3
        
        .txtResults(0).Text = vbNullString
        .txtResults(0).Left = .cbo(0).Left
        .txtResults(0).Top = .cmdRun.Top + .cmdRun.Height + 200
        .txtResults(0).TabStop = False
        .txtResults(0).Width = .fraMain.Width - 3000
        .txtResults(0).Height = 2000
        .txtResults(0).Visible = True
        
        Load .lblName(3)
        .lblName(3).Visible = True
        .lblName(3).Caption = "Results:"
        .lblName(3).Top = .txtResults(0).Top
        .lblName(3).Left = .lblName(0).Left
    End With
ExitFunction:
    Unit_Converter = True
    Exit Function
End Function

Public Function SendSMSWorld(ByRef InForm As Form) As Boolean
    With InForm
        .lblName(0).Visible = True
        .lblName(0).Caption = "From Email Address:"
        .lblName(0).Top = 300
        .lblName(0).Left = 120
        .txt(0).Text = vbNullString
        .txt(0).Left = 2000
        .txt(0).Top = 300
        .txt(0).Width = 4000
        .txt(0).MaxLength = 100
        .txt(0).Visible = True
        Load .lblName(1)
        .lblName(1).Visible = True
        .lblName(1).Caption = "Country Code:"
        .lblName(1).Top = .txt(0).Top + .txt(0).Height + 200
        .lblName(1).Left = .lblName(0).Left
        Load .txt(1)
        .txt(1).Text = vbNullString
        .txt(1).Left = .txt(0).Left
        .txt(1).Top = .lblName(1).Top
        .txt(1).Width = 300
        .txt(1).MaxLength = 2
        .txt(1).Visible = True
        Load .lblName(2)
        .lblName(2).Visible = True
        .lblName(2).Caption = "Mobile Number:"
        .lblName(2).Top = .txt(1).Top + .txt(1).Height + 200
        .lblName(2).Left = .lblName(0).Left
        Load .txt(2)
        .txt(2).Text = vbNullString
        .txt(2).Left = .txt(0).Left
        .txt(2).Top = .lblName(2).Top
        .txt(2).Width = 2000
        .txt(2).MaxLength = 30
        .txt(2).Visible = True
        Load .lblName(3)
        .lblName(3).Visible = True
        .lblName(3).Caption = "Message:"
        .lblName(3).Top = .txt(2).Top + .txt(2).Height + 200
        .lblName(3).Left = .lblName(0).Left
        Load .txt(3)
        .txt(3).Text = vbNullString
        .txt(3).Left = .txt(0).Left
        .txt(3).Top = .lblName(3).Top
        .txt(3).Width = 7000
        .txt(3).MaxLength = 150
        .txt(3).Visible = True
                                    
        .cmdClear.Visible = True
        .cmdClear.Left = .lblName(0).Left
        .cmdClear.Top = .txt(3).Top + .txt(3).Height + 200
        .cmdRun.Visible = True
        .cmdRun.Left = .txt(0).Left
        .cmdRun.Top = .cmdClear.Top
        
        .txt(0).TabIndex = 0
        .txt(1).TabIndex = 1
        .txt(2).TabIndex = 2
        .txt(3).TabIndex = 3
        .cmdRun.TabIndex = 4
        
        .txtResults(0).Text = vbNullString
        .txtResults(0).Left = .txt(0).Left
        .txtResults(0).Top = .cmdRun.Top + .cmdRun.Height + 200
        .txtResults(0).TabStop = False
        .txtResults(0).Width = .fraMain.Width - 3000
        .txtResults(0).Height = 2000
        .txtResults(0).Visible = True
        
        Load .lblName(4)
        .lblName(4).Visible = True
        .lblName(4).Caption = "Results:"
        .lblName(4).Top = .txtResults(0).Top
        .lblName(4).Left = .lblName(0).Left
    End With
ExitFunction:
    SendSMSWorld = True
    Exit Function
End Function

Private Function ValidateEmail(ByRef InForm As Form) As Boolean
    With InForm
        .lblName(0).Visible = True
        .lblName(0).Caption = "Email:"
        .lblName(0).Top = 300
        .lblName(0).Left = 120
        .txt(0).Left = 2000
        .txt(0).Top = 300
        .txt(0).Width = 4000
        .txt(0).MaxLength = 100
        .txt(0).Text = vbNullString
        .txt(0).Visible = True
                                    
        .cmdClear.Visible = True
        .cmdClear.Left = .lblName(0).Left
        .cmdClear.Top = .txt(0).Top + .txt(0).Height + 200
        .cmdRun.Visible = True
        .cmdRun.Left = .txt(0).Left
        .cmdRun.Top = .cmdClear.Top
        
        .txt(0).TabIndex = 0
        .cmdRun.TabIndex = 1
        
        .txtResults(0).Text = vbNullString
        .txtResults(0).Left = .txt(0).Left
        .txtResults(0).Top = .cmdRun.Top + .cmdRun.Height + 200
        .txtResults(0).TabStop = False
        .txtResults(0).Width = .fraMain.Width - 3000
        .txtResults(0).Height = 2000
        .txtResults(0).Visible = True
        On Error Resume Next
        Load .lblName(1)
        .lblName(1).Visible = True
        .lblName(1).Caption = "Results:"
        .lblName(1).Top = .txtResults(0).Top
        .lblName(1).Left = .lblName(0).Left
    End With
ExitFunction:
    ValidateEmail = True
    Exit Function
End Function

Public Function SendFax(ByRef InForm As Form) As Boolean
    With InForm
        .lblName(0).Visible = True
        .lblName(0).Caption = "From Email Address:"
        .lblName(0).Top = 300
        .lblName(0).Left = 120
        .txt(0).Text = vbNullString
        .txt(0).Left = 2000
        .txt(0).Top = 300
        .txt(0).Width = 4000
        .txt(0).MaxLength = 100
        .txt(0).Visible = True
        Load .lblName(1)
        .lblName(1).Visible = True
        .lblName(1).Caption = "Subject:"
        .lblName(1).Top = .txt(0).Top + .txt(0).Height + 200
        .lblName(1).Left = .lblName(0).Left
        Load .txt(1)
        .txt(1).Text = vbNullString
        .txt(1).Left = .txt(0).Left
        .txt(1).Top = .lblName(1).Top
        .txt(1).Width = 4000
        .txt(1).MaxLength = 100
        .txt(1).Visible = True
        Load .lblName(2)
        .lblName(2).Visible = True
        .lblName(2).Caption = "Fax Number:"
        .lblName(2).Top = .txt(1).Top + .txt(1).Height + 200
        .lblName(2).Left = .lblName(0).Left
        Load .txt(2)
        .txt(2).Text = vbNullString
        .txt(2).Left = .txt(0).Left
        .txt(2).Top = .lblName(2).Top
        .txt(2).Width = 2000
        .txt(2).MaxLength = 30
        .txt(2).Visible = True
        Load .lblName(3)
        .lblName(3).Visible = True
        .lblName(3).Caption = "Body:"
        .lblName(3).Top = .txt(2).Top + .txt(2).Height + 200
        .lblName(3).Left = .lblName(0).Left
        Load .txt(3)
        .txt(3).Text = vbNullString
        .txt(3).Left = .txt(0).Left
        .txt(3).Top = .lblName(3).Top
        .txt(3).Width = 7000
        .txt(3).MaxLength = 0
        .txt(3).Visible = True
        Load .lblName(4)
        .lblName(4).Visible = True
        .lblName(4).Caption = "To Name:"
        .lblName(4).Top = .txt(3).Top + .txt(3).Height + 200
        .lblName(4).Left = .lblName(0).Left
        Load .txt(4)
        .txt(4).Text = vbNullString
        .txt(4).Left = .txt(0).Left
        .txt(4).Top = .lblName(4).Top
        .txt(4).Width = 3000
        .txt(4).MaxLength = 100
        .txt(4).Visible = True
                                                                        
        .cmdClear.Visible = True
        .cmdClear.Left = .lblName(0).Left
        .cmdClear.Top = .txt(4).Top + .txt(4).Height + 200
        .cmdRun.Visible = True
        .cmdRun.Left = .txt(0).Left
        .cmdRun.Top = .cmdClear.Top
        
        .txt(0).TabIndex = 0
        .txt(1).TabIndex = 1
        .txt(2).TabIndex = 2
        .txt(3).TabIndex = 3
        .txt(4).TabIndex = 4
        .cmdRun.TabIndex = 5
        
        .txtResults(0).Text = vbNullString
        .txtResults(0).Left = .txt(0).Left
        .txtResults(0).Top = .cmdRun.Top + .cmdRun.Height + 200
        .txtResults(0).TabStop = False
        .txtResults(0).Width = .fraMain.Width - 3000
        .txtResults(0).Height = 2000
        .txtResults(0).Visible = True
        
        Load .lblName(5)
        .lblName(5).Visible = True
        .lblName(5).Caption = "Results:"
        .lblName(5).Top = .txtResults(0).Top
        .lblName(5).Left = .lblName(0).Left
    End With
ExitFunction:
    SendFax = True
    Exit Function
End Function

Public Function Run_WebService(ByRef InForm As Form) As Boolean
Dim StrNode As String
Dim IsOk As Boolean
Dim StrResults As String
Dim ObjXMLHTTP As Object
Dim ObjXMLDOM As Object
Const WebSite As String = "http://www.webservicex.net/"
Dim iPos As Long, jPos As Long
Dim StrCallingWeb As String

    IsOk = False
    Set ObjXMLHTTP = CreateObject("Microsoft.XMLHTTP")
    Set ObjXMLDOM = CreateObject("Microsoft.XMLDOM")
    
    With InForm
        .txtResults(0).Text = vbNullString
        StrNode = .tvTreeView.SelectedItem.Key
        Select Case StrNode
            Case "B1"
                 'make sure values have been entered
                If Trim(.txt(0).Text) = vbNullString Then
                    MsgBox "Credit Card number is required.", vbInformation, "Required Field"
                    IsOk = False
                Else
                    With ObjXMLHTTP
                        .Open "GET", WebSite & "CreditCard.asmx/ValidateCardNumber?cardType=" & InForm.cbo(0).Text & "&cardNumber=" & InForm.txt(0).Text, False
                        .Send
                        DoEvents
                        
                        If .Status = 200 Then 'is ok
                            'Load the XML document from the webservice
                            ObjXMLDOM.LoadXml .ResponseText
                            
                            'check if there are any errors
                            If ObjXMLDOM.parseError.ErrorCode <> 0 Then
                                InForm.txtResults(0).Text = .ResponseText
                                IsOk = False
                            Else
                                'Get the results
                                InForm.txtResults(0).Text = ObjXMLDOM.GetElementsByTagName("string").Item(0).Text
                                IsOk = True
                            End If
                        Else
                            MsgBox "ERROR - " & .Status, vbCritical, "ERROR"
                            IsOk = False
                        End If
                    End With
                                        
                End If
            Case "B2"
                If Trim(.txt(0).Text) = vbNullString Then
                    MsgBox "Year is required.", vbInformation, "Required Field"
                    IsOk = False
                ElseIf Trim(.txt(1).Text) = vbNullString Then
                    MsgBox "Interest is required.", vbInformation, "Required Field"
                    IsOk = False
                ElseIf Trim(.txt(2).Text) = vbNullString Then
                    MsgBox "Loan Amount is required.", vbInformation, "Required Field"
                    IsOk = False
                ElseIf Trim(.txt(3).Text) = vbNullString Then
                    MsgBox "Annual Tax is required.", vbInformation, "Required Field"
                    IsOk = False
                ElseIf Trim(.txt(4).Text) = vbNullString Then
                    MsgBox "Annual Insurance is required.", vbInformation, "Required Field"
                    IsOk = False
                Else
                    With ObjXMLHTTP
                        .Open "GET", WebSite & "mortgage.asmx/GetMortgagePayment?Years=" & InForm.txt(0).Text _
                                                                          & "&Interest=" & InForm.txt(1).Text _
                                                                          & "&LoanAmount=" & InForm.txt(2).Text _
                                                                          & "&AnnualTax=" & InForm.txt(3).Text _
                                                                          & "&AnnualInsurance=" & InForm.txt(4).Text, False
                        .Send
                        DoEvents
                        
                        If .Status = 200 Then 'is ok
                            'Load the XML document from the webservice
                            ObjXMLDOM.LoadXml .ResponseText
                            
                            'check if there are any errors
                            If ObjXMLDOM.parseError.ErrorCode <> 0 Then
                                InForm.txtResults(0).Text = .ResponseText
                                IsOk = False
                            Else
                                'Get the results
                                StrResults = "Monthly Principal And Interest = " & Round(ObjXMLDOM.GetElementsByTagName("MonthlyPrincipalAndInterest").Item(0).Text, 2)
                                StrResults = StrResults & vbCrLf & "Monthly Tax = " & Round(ObjXMLDOM.GetElementsByTagName("MonthlyTax").Item(0).Text, 2)
                                StrResults = StrResults & vbCrLf & "Monthly Insurance = " & Round(ObjXMLDOM.GetElementsByTagName("MonthlyInsurance").Item(0).Text, 2)
                                StrResults = StrResults & vbCrLf & "Total Payment = " & Round(ObjXMLDOM.GetElementsByTagName("TotalPayment").Item(0).Text, 2)
                                InForm.txtResults(0).Text = StrResults
                                IsOk = True
                            End If
                        Else
                            MsgBox "ERROR - " & .Status, vbCritical, "ERROR"
                            IsOk = False
                        End If
                    End With
                                        
                End If
            Case "B3"
                 'make sure values have been entered
                If Trim(.cbo(0).Text) = vbNullString Then
                    MsgBox "From Country is required.", vbInformation, "Required Field"
                    IsOk = False
                ElseIf Trim(.cbo(1).Text) = vbNullString Then
                    MsgBox "To Country is required.", vbInformation, "Required Field"
                    IsOk = False
                Else
                    With ObjXMLHTTP
                        .Open "GET", WebSite & "CurrencyConvertor.asmx/ConversionRate?FromCurrency=" & Mid$(InForm.cbo(0).Text, 1, 3) & "&ToCurrency=" & Mid$(InForm.cbo(1).Text, 1, 3), False
                        .Send
                        DoEvents
                        
                        If .Status = 200 Then 'is ok
                            'Load the XML document from the webservice
                            ObjXMLDOM.LoadXml .ResponseText
                            
                            'check if there are any errors
                            If ObjXMLDOM.parseError.ErrorCode <> 0 Then
                                InForm.txtResults(0).Text = .ResponseText
                                IsOk = False
                            Else
                                'Get the results
                                InForm.txtResults(0).Text = ObjXMLDOM.GetElementsByTagName("double").Item(0).Text
                                IsOk = True
                            End If
                        Else
                            MsgBox "ERROR - " & .Status, vbCritical, "ERROR"
                            IsOk = False
                        End If
                    End With
                End If
            Case "S1"
                 'make sure values have been entered
                If Trim(.txt(0).Text) = vbNullString Then
                    MsgBox "Country Code is required.", vbInformation, "Required Field"
                    IsOk = False
                Else
                    With ObjXMLHTTP
                        .Open "GET", WebSite & "country.asmx/GetCountryByCountryCode?CountryCode=" & InForm.txt(0).Text, False
                        .Send
                        DoEvents
                        
                        If .Status = 200 Then 'is ok
                            'Load the XML document from the webservice
                            ObjXMLDOM.LoadXml .ResponseText
                            
                            'check if there are any errors
                            If ObjXMLDOM.parseError.ErrorCode <> 0 Then
                                InForm.txtResults(0).Text = .ResponseText
                                IsOk = False
                            Else
                                'Get the results
                                StrResults = ObjXMLDOM.GetElementsByTagName("string").Item(0).Text
                                iPos = InStr(1, StrResults, "<name>", vbTextCompare)
                                iPos = iPos + Len("<name>")
                                jPos = InStr(1, StrResults, "</name>", vbTextCompare)
                                If Val(jPos) = 0 Then
                                    StrResults = "Invalid Country Code."
                                Else
                                    StrResults = Mid$(StrResults, iPos, jPos - iPos)
                                End If
                                InForm.txtResults(0).Text = StrResults
                                IsOk = True
                            End If
                        Else
                            MsgBox "ERROR - " & .Status, vbCritical, "ERROR"
                            IsOk = False
                        End If
                    End With
                End If
            Case "S2"
                'make sure values have been entered
                If Trim(.cbo(0).Text) = vbNullString Then
                    MsgBox "Country is required.", vbInformation, "Required Field"
                    IsOk = False
                Else
                    With ObjXMLHTTP
                        .Open "GET", WebSite & "country.asmx/GetISD?CountryName=" & InForm.cbo(0).Text, False
                        .Send
                        DoEvents
                        
                        If .Status = 200 Then 'is ok
                            'Load the XML document from the webservice
                            ObjXMLDOM.LoadXml .ResponseText
                            
                            'check if there are any errors
                            If ObjXMLDOM.parseError.ErrorCode <> 0 Then
                                InForm.txtResults(0).Text = .ResponseText
                                IsOk = False
                            Else
                                'Get the results
                                StrResults = ObjXMLDOM.GetElementsByTagName("string").Item(0).Text
                                iPos = InStr(1, StrResults, "<code>", vbTextCompare)
                                iPos = iPos + Len("<code>")
                                jPos = InStr(1, StrResults, "</code>", vbTextCompare)
                                If Val(jPos) = 0 Then
                                    StrResults = "Invalid Country."
                                Else
                                    StrResults = Mid$(StrResults, iPos, jPos - iPos)
                                End If
                                InForm.txtResults(0).Text = StrResults
                                IsOk = True
                            End If
                        Else
                            MsgBox "ERROR - " & .Status, vbCritical, "ERROR"
                            IsOk = False
                        End If
                    End With
                End If
            Case "S3"
                IsOk = Get_All_Countries(InForm, False, , False)
            Case "S4"
                'make sure values have been entered
                If Trim(.cbo(0).Text) = vbNullString Then
                    MsgBox "Country is required.", vbInformation, "Required Field"
                    IsOk = False
                Else
                    With ObjXMLHTTP
                        .Open "GET", WebSite & "country.asmx/GetCurrencyByCountry?CountryName=" & InForm.cbo(0).Text, False
                        .Send
                        DoEvents
                        
                        If .Status = 200 Then 'is ok
                            'Load the XML document from the webservice
                            ObjXMLDOM.LoadXml .ResponseText
                            
                            'check if there are any errors
                            If ObjXMLDOM.parseError.ErrorCode <> 0 Then
                                InForm.txtResults(0).Text = .ResponseText
                                IsOk = False
                            Else
                                'Get the results
                                StrResults = ObjXMLDOM.GetElementsByTagName("string").Item(0).Text
                                iPos = InStr(1, StrResults, "<CurrencyCode>", vbTextCompare)
                                iPos = iPos + Len("<CurrencyCode>")
                                jPos = InStr(1, StrResults, "</CurrencyCode>", vbTextCompare)
                                If Val(jPos) = 0 Then
                                    StrResults = "Invalid Country."
                                Else
                                    StrResults = Mid$(StrResults, iPos, jPos - iPos)
                                End If
                                InForm.txtResults(0).Text = StrResults
                                IsOk = True
                            End If
                        Else
                            MsgBox "ERROR - " & .Status, vbCritical, "ERROR"
                            IsOk = False
                        End If
                    End With
                End If
            Case "V1", "V2", "V3", "V4", "V5", "V6", "V7", "V8"
                'make sure values have been entered
                If Trim(.txt(0).Text) = vbNullString Then
                    MsgBox "Values is required.", vbInformation, "Required Field"
                    IsOk = False
                ElseIf Trim(.cbo(0).Text) = vbNullString Then
                    MsgBox "From Unit is required.", vbInformation, "Required Field"
                    IsOk = False
                ElseIf Trim(.cbo(1).Text) = vbNullString Then
                    MsgBox "To Unit is required.", vbInformation, "Required Field"
                    IsOk = False
                Else
                    Select Case StrNode
                        Case "V1"
                            StrCallingWeb = "ConvertSpeed.asmx/ConvertSpeed?speed=" & InForm.txt(0).Text _
                                             & "&FromUnit=" & InForm.cbo(0).Text _
                                             & "&ToUnit=" & InForm.cbo(1).Text
                        Case "V2"
                            StrCallingWeb = "ConvertWeight.asmx/ConvertWeight?Weight=" & InForm.txt(0).Text _
                                             & "&FromUnit=" & InForm.cbo(0).Text _
                                             & "&ToUnit=" & InForm.cbo(1).Text
                        Case "V3"
                            StrCallingWeb = "length.asmx/ChangeLengthUnit?LengthValue=" & InForm.txt(0).Text _
                                             & "&fromLengthUnit=" & InForm.cbo(0).Text _
                                             & "&toLengthUnit=" & InForm.cbo(1).Text
                        Case "V4"
                            StrCallingWeb = "length.asmx/ChangeLengthUnit?LengthValue=" & InForm.txt(0).Text _
                                             & "&fromLengthUnit=" & InForm.cbo(0).Text _
                                             & "&toLengthUnit=" & InForm.cbo(1).Text
                        Case "V5"
                            StrCallingWeb = "convertMetricWeight.asmx/ChangeMetricWeightUnit?MetricWeightValue=" & InForm.txt(0).Text _
                                             & "&fromMetricWeightUnit=" & InForm.cbo(0).Text _
                                             & "&toMetricWeightUnit=" & InForm.cbo(1).Text
                        Case "V6"
                            StrCallingWeb = "ConvertTorque.asmx/ChangeTorqueUnit?TorqueValue=" & InForm.txt(0).Text _
                                             & "&fromTorqueUnit=" & InForm.cbo(0).Text _
                                             & "&toTorqueUnit=" & InForm.cbo(1).Text
                        Case "V7"
                            StrCallingWeb = "ConverPower.asmx/ChangePowerUnit?PowerValue=" & InForm.txt(0).Text _
                                             & "&fromPowerUnit=" & InForm.cbo(0).Text _
                                             & "&toPowerUnit=" & InForm.cbo(1).Text
                        Case "V8"
                            StrCallingWeb = "ConvertAcceleration.asmx/ChangeAccelerationUnit?AccelerationValue=" & InForm.txt(0).Text _
                                             & "&fromAccelerationUnit=" & InForm.cbo(0).Text _
                                             & "&toAccelerationUnit=" & InForm.cbo(1).Text
                    End Select
                
                    With ObjXMLHTTP
                        .Open "GET", WebSite & StrCallingWeb, False
                        .Send
                        DoEvents
                        
                        If .Status = 200 Then 'is ok
                            'Load the XML document from the webservice
                            ObjXMLDOM.LoadXml .ResponseText
                            
                            'check if there are any errors
                            If ObjXMLDOM.parseError.ErrorCode <> 0 Then
                                InForm.txtResults(0).Text = .ResponseText
                                IsOk = False
                            Else
                                'Get the results
                                StrResults = Round(ObjXMLDOM.GetElementsByTagName("double").Item(0).Text, 4)
                                InForm.txtResults(0).Text = StrResults
                                IsOk = True
                            End If
                        Else
                            MsgBox "ERROR - " & .Status, vbCritical, "ERROR"
                            IsOk = False
                        End If
                    End With
                End If
            Case "C1"
                If Trim(.txt(0).Text) = vbNullString Then
                    MsgBox "Email Address is required.", vbInformation, "Required Field"
                    IsOk = False
                ElseIf Trim(.txt(1).Text) = vbNullString Then
                    MsgBox "Country Code is required.", vbInformation, "Required Field"
                    IsOk = False
                ElseIf Trim(.txt(2).Text) = vbNullString Then
                    MsgBox "Phone number is required.", vbInformation, "Required Field"
                    IsOk = False
                ElseIf Trim(.txt(3).Text) = vbNullString Then
                    MsgBox "Message is required.", vbInformation, "Required Field"
                    IsOk = False
                Else
                    StrCallingWeb = "sendsmsworld.asmx/sendSMS?FromEmailAddress=" & InForm.txt(0).Text _
                                             & "&CountryCode=" & InForm.txt(1).Text _
                                             & "&MobileNumber=" & InForm.txt(2).Text _
                                             & "&Message=" & InForm.txt(3).Text
                                             
                    With ObjXMLHTTP
                        .Open "GET", WebSite & StrCallingWeb, False
                        .Send
                        DoEvents
                        
                        If .Status = 200 Then 'is ok
                            'Load the XML document from the webservice
                            ObjXMLDOM.LoadXml .ResponseText
                            
                            'check if there are any errors
                            If ObjXMLDOM.parseError.ErrorCode <> 0 Then
                                InForm.txtResults(0).Text = .ResponseText
                                IsOk = False
                            Else
                                'Get the results
                                StrResults = ObjXMLDOM.GetElementsByTagName("string").Item(0).Text
                                InForm.txtResults(0).Text = StrResults
                                IsOk = True
                            End If
                        Else
                            MsgBox "ERROR - " & .Status, vbCritical, "ERROR"
                            IsOk = False
                        End If
                    End With
                End If
            Case "C2"
                If Trim(.txt(0).Text) = vbNullString Then
                    MsgBox "Email Address is required.", vbInformation, "Required Field"
                    IsOk = False
                Else
                    StrCallingWeb = "ValidateEmail.asmx/IsValidEmail?Email=" & InForm.txt(0).Text
                                             
                    With ObjXMLHTTP
                        .Open "GET", WebSite & StrCallingWeb, False
                        .Send
                        DoEvents
                        
                        If .Status = 200 Then 'is ok
                            'Load the XML document from the webservice
                            ObjXMLDOM.LoadXml .ResponseText
                            
                            'check if there are any errors
                            If ObjXMLDOM.parseError.ErrorCode <> 0 Then
                                InForm.txtResults(0).Text = .ResponseText
                                IsOk = False
                            Else
                                'Get the results
                                StrResults = ObjXMLDOM.GetElementsByTagName("boolean").Item(0).Text
                                InForm.txtResults(0).Text = StrResults
                                IsOk = True
                            End If
                        Else
                            MsgBox "ERROR - " & .Status, vbCritical, "ERROR"
                            IsOk = False
                        End If
                    End With
                End If
            Case "C3"
                If Trim(.txt(0).Text) = vbNullString Then
                    MsgBox "From Email Address is required.", vbInformation, "Required Field"
                    IsOk = False
                ElseIf Trim(.txt(1).Text) = vbNullString Then
                    MsgBox "Subject is required.", vbInformation, "Required Field"
                    IsOk = False
                ElseIf Trim(.txt(2).Text) = vbNullString Then
                    MsgBox "Fax number is required.", vbInformation, "Required Field"
                    IsOk = False
                ElseIf Trim(.txt(3).Text) = vbNullString Then
                    MsgBox "Body is required.", vbInformation, "Required Field"
                    IsOk = False
                ElseIf Trim(.txt(4).Text) = vbNullString Then
                    MsgBox "To Name is required.", vbInformation, "Required Field"
                    IsOk = False
                Else
                    StrCallingWeb = "fax.asmx/SendTextToFax?FromEmail=" & InForm.txt(0).Text _
                                             & "&Subject=" & InForm.txt(1).Text _
                                             & "&FaxNumber=" & InForm.txt(2).Text _
                                             & "&BodyText=" & InForm.txt(3).Text _
                                             & "&ToName=" & InForm.txt(4).Text
                                             
                    With ObjXMLHTTP
                        .Open "GET", WebSite & StrCallingWeb, False
                        .Send
                        DoEvents
                        
                        If .Status = 200 Then 'is ok
                            'Load the XML document from the webservice
                            ObjXMLDOM.LoadXml .ResponseText
                            
                            'check if there are any errors
                            If ObjXMLDOM.parseError.ErrorCode <> 0 Then
                                InForm.txtResults(0).Text = .ResponseText
                                IsOk = False
                            Else
                                'Get the results
                                StrResults = ObjXMLDOM.GetElementsByTagName("string").Item(0).Text
                                InForm.txtResults(0).Text = StrResults
                                IsOk = True
                            End If
                        Else
                            MsgBox "ERROR - " & .Status, vbCritical, "ERROR"
                            IsOk = False
                        End If
                    End With
                End If
            Case "G1"
            Case "G2"
            Case "U1"
        End Select
    End With
ExitFunction:
    Set ObjXMLHTTP = Nothing
    Set ObjXMLDOM = Nothing
    StrResults = vbNullString
    
    Run_WebService = IsOk
    Exit Function
End Function

Private Function Get_All_Countries(ByRef InForm As Form, ByVal LoadCombo As Boolean, Optional ByRef InCbo As ComboBox, Optional ByVal ReLoad As Boolean = False) As Boolean
Dim IsOk As Boolean
Dim StrResults As String
Dim StrParse As String
Dim ObjXMLHTTP As Object
Dim ObjXMLDOM As Object
Const WebSite As String = "http://www.webservicex.net/"
Dim iPos As Long, jPos As Long, i As Long

    IsOk = False
    
    If ReLoad Then
        Set ObjXMLHTTP = CreateObject("Microsoft.XMLHTTP")
        Set ObjXMLDOM = CreateObject("Microsoft.XMLDOM")
        
        With ObjXMLHTTP
            .Open "GET", WebSite & "country.asmx/GetCountries?", False
            .Send
            DoEvents
            
            If .Status = 200 Then 'is ok
                'Load the XML document from the webservice
                ObjXMLDOM.LoadXml .ResponseText
                
                'check if there are any errors
                If ObjXMLDOM.parseError.ErrorCode <> 0 Then
                    InForm.txtResults(0).Text = .ResponseText
                    IsOk = False
                Else
                    'Get the results
                    StrResults = ObjXMLDOM.GetElementsByTagName("string").Item(0).Text
                    
                    StrParse = Replace(StrResults, "<NewDataSet>", vbNullString, , , vbTextCompare)
                    StrParse = Replace(StrParse, "</NewDataSet>", vbNullString, , , vbTextCompare)
                    StrParse = Replace(StrParse, "<Table>", vbNullString, , , vbTextCompare)
                    StrParse = Replace(StrParse, "</Table>", vbNullString, , , vbTextCompare)
                    StrParse = Replace(StrParse, "<Name>", vbNullString, , , vbTextCompare)
                    StrParse = Replace(StrParse, "</Name>", vbTab, , , vbTextCompare)
                    StrParse = Replace(StrParse, "&amp;", "&", , , vbTextCompare)
                    
                    iPos = 1
                    jPos = 1
                    StrResults = StrParse
                    StrParse = vbNullString
                    For i = 0 To Len(StrResults)
                        If Val(iPos) = 0 Then
                            Exit For
                        Else
                            iPos = InStr(iPos, StrResults, " ")
                            jPos = InStr(iPos, StrResults, vbTab)
                            If Val(jPos) = 0 Then
                                'Exit For
                            Else
                                If iPos = 2 Then
                                    StrParse = StrParse & Mid$(StrResults, iPos + 7, jPos - (iPos + 7))
                                Else
                                    If Trim(StrParse) <> vbNullString Then
                                        StrParse = StrParse & vbCrLf & Mid$(StrResults, iPos + 10, jPos - (iPos + 10))
                                    Else
                                        StrParse = StrParse & Mid$(StrResults, iPos + 10, jPos - (iPos + 10))
                                    End If
                                End If
                            End If
                        End If
                        iPos = jPos
                    Next i
                    StrResults = StrParse
                    
                    aJoin = Split(StrResults, vbCrLf)
                    
                    IsOk = True
                End If
            Else
                MsgBox "ERROR - " & .Status, vbCritical, "ERROR"
                IsOk = False
            End If
        End With
    Else
        If LoadCombo Then
            InCbo.Clear
            For i = 0 To UBound(aJoin)
                InCbo.AddItem aJoin(i)
            Next i
        Else
            InForm.txtResults(0).Text = Join(aJoin, vbCrLf)
        End If
        IsOk = True
    End If
    
ExitFunction:
    Set ObjXMLHTTP = Nothing
    Set ObjXMLDOM = Nothing
    StrResults = vbNullString
    StrParse = vbNullString
    Get_All_Countries = IsOk
    Exit Function
    
End Function

Private Function Load_Combos(ByRef InCbo As ComboBox, ByVal InType As String)
'This function will populate the InCBO
    Select Case InType
        Case "V1"
            InCbo.Clear
            InCbo.AddItem "centimetersPersecond"
            InCbo.AddItem "metersPersecond"
            InCbo.AddItem "feetPersecond"
            InCbo.AddItem "feetPerminute"
            InCbo.AddItem "milesPerhour"
            InCbo.AddItem "kilometersPerhour"
            InCbo.AddItem "furlongsPermin"
            InCbo.AddItem "knots"
            InCbo.AddItem "leaguesPerday"
            InCbo.AddItem "Mach"
        Case "V2"
            InCbo.Clear
            InCbo.AddItem "Grains"
            InCbo.AddItem "Scruples"
            InCbo.AddItem "Carats"
            InCbo.AddItem "Grams"
            InCbo.AddItem "Pennyweight"
            InCbo.AddItem "DramAvoir"
            InCbo.AddItem "DramApoth"
            InCbo.AddItem "OuncesAvoir"
            InCbo.AddItem "OuncesTroyApoth"
            InCbo.AddItem "Poundals"
            InCbo.AddItem "PoundsTroy"
            InCbo.AddItem "PoundsAvoir"
            InCbo.AddItem "Kilograms"
            InCbo.AddItem "Stones"
            InCbo.AddItem "QuarterUS"
            InCbo.AddItem "Slugs"
            InCbo.AddItem "weight100UScwt"
            InCbo.AddItem "ShortTons"
            InCbo.AddItem "MetricTonsTonne"
            InCbo.AddItem "LongTons"
        Case "V3"
            InCbo.Clear
            InCbo.AddItem "Angstroms "
            InCbo.AddItem "Nanometers"
            InCbo.AddItem "Microinch"
            InCbo.AddItem "Microns"
            InCbo.AddItem "Mils"
            InCbo.AddItem "Millimeters"
            InCbo.AddItem "Centimeters"
            InCbo.AddItem "Inches"
            InCbo.AddItem "Links"
            InCbo.AddItem "Spans"
            InCbo.AddItem "Feet"
            InCbo.AddItem "Cubits"
            InCbo.AddItem "Varas"
            InCbo.AddItem "Yards"
            InCbo.AddItem "Meters"
            InCbo.AddItem "Fathoms"
            InCbo.AddItem "Rods"
            InCbo.AddItem "Chains"
            InCbo.AddItem "Furlongs"
            InCbo.AddItem "Cablelengths"
            InCbo.AddItem "Kilometers"
            InCbo.AddItem "Miles"
            InCbo.AddItem "Nauticalmile"
            InCbo.AddItem "League"
            InCbo.AddItem "Nauticalleague"
        Case "V4"
            InCbo.Clear
            InCbo.AddItem "acre"
            InCbo.AddItem "acrecommercial"
            InCbo.AddItem "acreIreland"
            InCbo.AddItem "acresurvey"
            InCbo.AddItem "are"
            InCbo.AddItem "baseboxtinplatedsteel"
            InCbo.AddItem "binTaiwan"
            InCbo.AddItem "buJapan"
            InCbo.AddItem "canteroEcuador"
            InCbo.AddItem "caoVietnam"
            InCbo.AddItem "centaire"
            InCbo.AddItem "circularfootinternational"
            InCbo.AddItem "circularfootUSsurvey"
            InCbo.AddItem "circularinchinternational"
            InCbo.AddItem "circularinchUSsurvey"
            InCbo.AddItem "hectare"
            InCbo.AddItem "laborCanada"
            InCbo.AddItem "laborUSsurvey"
            InCbo.AddItem "manzanaCostaRican"
            InCbo.AddItem "manzanaArgentina"
            InCbo.AddItem "rood"
            InCbo.AddItem "saoVietnam"
            InCbo.AddItem "scrupleancientRome"
            InCbo.AddItem "sectionUSsurvey"
            InCbo.AddItem "square"
            InCbo.AddItem "squareSriLanka"
            InCbo.AddItem "squareangstrom"
            InCbo.AddItem "squareastronomicalunit"
            InCbo.AddItem "squarecablelengthUSsurvey"
            InCbo.AddItem "squarecaliber"
            InCbo.AddItem "squarecentimeter"
            InCbo.AddItem "squarechainGunterUSsurvey"
            InCbo.AddItem "squarechainRamdenEngineer"
            InCbo.AddItem "squarecubit"
            InCbo.AddItem "squarecubitancientEgypt"
            InCbo.AddItem "squaredigit"
            InCbo.AddItem "squarefathom"
            InCbo.AddItem "squarefootinternational"
            InCbo.AddItem "squarefootUSsurvey"
            InCbo.AddItem "squarefurlongUSsurvey"
            InCbo.AddItem "squaregigameter"
            InCbo.AddItem "squarehectometer"
            InCbo.AddItem "squareinchinternational"
            InCbo.AddItem "squareinchUSsurvey"
            InCbo.AddItem "squarekilometer"
            InCbo.AddItem "squareleaguenautical"
            InCbo.AddItem "squareleagueUSstatute"
            InCbo.AddItem "squarelightyear"
            InCbo.AddItem "squarelinkGunterUSsurvey"
            InCbo.AddItem "squarelinkRamdenEngineer"
            InCbo.AddItem "squaremegameter"
            InCbo.AddItem "squaremeter"
            InCbo.AddItem "squaremicroinch"
            InCbo.AddItem "squaremicrometer"
            InCbo.AddItem "squaremicromicron"
            InCbo.AddItem "squaremicron"
            InCbo.AddItem "squaremil"
            InCbo.AddItem "squaremileinternational"
            InCbo.AddItem "squaremileintnautical"
            InCbo.AddItem "squaremileUSnautical"
            InCbo.AddItem "squaremileUSstatute"
            InCbo.AddItem "squaremileUSsurvey"
            InCbo.AddItem "squaremillimeter"
            InCbo.AddItem "squaremillimicron"
            InCbo.AddItem "squarenanometer"
            InCbo.AddItem "squareParisfootCanada"
            InCbo.AddItem "squareparsec"
            InCbo.AddItem "squareperchIreland"
            InCbo.AddItem "squareperchUSsurvey"
            InCbo.AddItem "squarepetameter"
            InCbo.AddItem "squarepicometer"
            InCbo.AddItem "squarerodNetherlands"
            InCbo.AddItem "squaretenthmeter"
            InCbo.AddItem "squareyardUSsurvey"
            InCbo.AddItem "squareyardinternational"
            InCbo.AddItem "squareyoctometer"
            InCbo.AddItem "squareyottameter"
            InCbo.AddItem "squarezeptometer"
            InCbo.AddItem "squarezettameter"
            InCbo.AddItem "townshipUSsurvey"
        Case "V5"
            InCbo.Clear
            InCbo.AddItem "microgram "
            InCbo.AddItem "milligram"
            InCbo.AddItem "centigram"
            InCbo.AddItem "decigram"
            InCbo.AddItem "gram"
            InCbo.AddItem "dekagram"
            InCbo.AddItem "hectogram"
            InCbo.AddItem "kilogram"
            InCbo.AddItem "metricton"
        Case "V6"
            InCbo.Clear
            InCbo.AddItem "DyneCentimeters "
            InCbo.AddItem "FootPounds"
            InCbo.AddItem "InchPounds"
            InCbo.AddItem "KilogramMeter"
            InCbo.AddItem "MeterNewtons"
        Case "V7"
            InCbo.Clear
            InCbo.AddItem "ergsPersec "
            InCbo.AddItem "milliwatts"
            InCbo.AddItem "watts"
            InCbo.AddItem "kiloCaloriesPermin"
            InCbo.AddItem "kiloCaloriesPersec"
            InCbo.AddItem "BTUPerhour"
            InCbo.AddItem "footlbsPersec"
            InCbo.AddItem "horsepower"
            InCbo.AddItem "kilowatts"
            InCbo.AddItem "megawatts"
            InCbo.AddItem "gigawatts"
        Case "V8"
            InCbo.Clear
            InCbo.AddItem "celo"
            InCbo.AddItem "centigal"
            InCbo.AddItem "centimeterPersquaresecond "
            InCbo.AddItem "decigal"
            InCbo.AddItem "decimeterPersquaresecond"
            InCbo.AddItem "dekameterPersquaresecond"
            InCbo.AddItem "footPersquaresecond"
            InCbo.AddItem "gunit"
            InCbo.AddItem "gal"
            InCbo.AddItem "galileo"
            InCbo.AddItem "gn"
            InCbo.AddItem "grav"
            InCbo.AddItem "hectometerPersquaresecond"
            InCbo.AddItem "inchPersquaresecond"
            InCbo.AddItem "kilometerPerhoursecond"
            InCbo.AddItem "kilometerPersquaresecond"
            InCbo.AddItem "leo"
            InCbo.AddItem "meterPersquaresecond"
            InCbo.AddItem "milePerhourminute"
            InCbo.AddItem "milePerhoursecond"
            InCbo.AddItem "milePersquaresecond"
            InCbo.AddItem "milligal"
            InCbo.AddItem "millimeterPersquaresecond"
        Case "B3"
            With InCbo
                .Clear
                .AddItem "AFA-Afghanistan Afghani"
                .AddItem "ALL-Albanian Lek"
                .AddItem "DZD-Algerian Dinar"
                .AddItem "ARS-Argentine Peso"
                .AddItem "AWG-Aruba Florin"
                .AddItem "AUD-Australian Dollar"
                .AddItem "BSD-Bahamian Dollar"
                .AddItem "BHD-Bahraini Dinar"
                .AddItem "BDT-Bangladesh Taka"
                .AddItem "BBD-Barbados Dollar"
                .AddItem "BZD-Belize Dollar"
                .AddItem "BMD-Bermuda Dollar"
                .AddItem "BTN-Bhutan Ngultrum"
                .AddItem "BOB-Bolivian Boliviano"
                .AddItem "BWP-Botswana Pula"
                .AddItem "BRL-Brazilian Real"
                .AddItem "GBP-British Pound"
                .AddItem "BND-Brunei Dollar"
                .AddItem "BIF-Burundi Franc"
                .AddItem "XOF-CFA Franc (BCEAO)"
                .AddItem "XAF-CFA Franc (BEAC)"
                .AddItem "KHR-Cambodia Riel"
                .AddItem "CAD-Canadian Dollar"
                .AddItem "CVE-Cape Verde Escudo"
                .AddItem "KYD-Cayman Islands Dollar"
                .AddItem "CLP-Chilean Peso"
                .AddItem "CNY-Chinese Yuan"
                .AddItem "COP-Colombian Peso"
                .AddItem "KMF-Comoros Franc"
                .AddItem "CRC-Costa Rica Colon"
                .AddItem "HRK-Croatian Kuna"
                .AddItem "CUP-Cuban Peso"
                .AddItem "CYP-Cyprus Pound"
                .AddItem "CZK-Czech Koruna"
                .AddItem "DKK-Danish Krone"
                .AddItem "DJF-Dijibouti Franc"
                .AddItem "DOP-Dominican Peso"
                .AddItem "XCD-East Caribbean Dollar"
                .AddItem "EGP-Egyptian Pound"
                .AddItem "SVC-El Salvador Colon"
                .AddItem "EEK-Estonian Kroon"
                .AddItem "ETB-Ethiopian Birr"
                .AddItem "EUR -Euro"
                .AddItem "FKP-Falkland Islands Pound"
                .AddItem "GMD-Gambian Dalasi"
                .AddItem "GHC-Ghanian Cedi"
                .AddItem "GIP-Gibraltar Pound"
                .AddItem "XAU-Gold Ounces"
                .AddItem "GTQ-Guatemala Quetzal"
                .AddItem "GNF-Guinea Franc"
                .AddItem "GYD-Guyana Dollar"
                .AddItem "HTG-Haiti Gourde"
                .AddItem "HNL-Honduras Lempira"
                .AddItem "HKD-Hong Kong Dollar"
                .AddItem "HUF-Hungarian Forint"
                .AddItem "ISK-Iceland Krona"
                .AddItem "INR-Indian Rupee"
                .AddItem "IDR-Indonesian Rupiah"
                .AddItem "IQD-Iraqi Dinar"
                .AddItem "ILS-Israeli Shekel"
                .AddItem "JMD-Jamaican Dollar"
                .AddItem "JPY-Japanese Yen"
                .AddItem "JOD-Jordanian Dinar"
                .AddItem "KZT-Kazakhstan Tenge"
                .AddItem "KES-Kenyan Shilling"
                .AddItem "KRW-Korean Won"
                .AddItem "KWD-Kuwaiti Dinar"
                .AddItem "LAK-Lao Kip"
                .AddItem "LVL-Latvian Lat"
                .AddItem "LBP-Lebanese Pound"
                .AddItem "LSL-Lesotho Loti"
                .AddItem "LRD-Liberian Dollar"
                .AddItem "LYD-Libyan Dinar"
                .AddItem "LTL-Lithuanian Lita"
                .AddItem "MOP-Macau Pataca"
                .AddItem "MKD-Macedonian Denar"
                .AddItem "MGF-Malagasy Franc"
                .AddItem "MWK-Malawi Kwacha"
                .AddItem "MYR-Malaysian Ringgit"
                .AddItem "MVR-Maldives Rufiyaa"
                .AddItem "MTL-Maltese Lira"
                .AddItem "MRO-Mauritania Ougulya"
                .AddItem "MUR-Mauritius Rupee"
                .AddItem "MXN-Mexican Peso"
                .AddItem "MDL-Moldovan Leu"
                .AddItem "MNT-Mongolian Tugrik"
                .AddItem "MAD-Moroccan Dirham"
                .AddItem "MZM-Mozambique Metical"
                .AddItem "MMK-Myanmar Kyat"
                .AddItem "NAD-Namibian Dollar"
                .AddItem "NPR-Nepalese Rupee"
                .AddItem "ANG-Neth Antilles Guilder"
                .AddItem "NZD-New Zealand Dollar"
                .AddItem "NIO-Nicaragua Cordoba"
                .AddItem "NGN-Nigerian Naira"
                .AddItem "KPW-North Korean Won"
                .AddItem "NOK-Norwegian Krone"
                .AddItem "OMR-Omani Rial"
                .AddItem "XPF-Pacific Franc"
                .AddItem "PKR-Pakistani Rupee"
                .AddItem "XPD-Palladium Ounces"
                .AddItem "PAB-Panama Balboa"
                .AddItem "PGK-Papua New Guinea Kina"
                .AddItem "PYG-Paraguayan Guarani"
                .AddItem "PEN-Peruvian Nuevo Sol"
                .AddItem "PHP-Philippine Peso"
                .AddItem "XPT-Platinum Ounces"
                .AddItem "PLN-Polish Zloty"
                .AddItem "QAR-Qatar Rial"
                .AddItem "ROL-Romanian Leu"
                .AddItem "RUB-Russian Rouble"
                .AddItem "WST-Samoa Tala"
                .AddItem "STD-Sao Tome Dobra"
                .AddItem "SAR-Saudi Arabian Riyal"
                .AddItem "SCR-Seychelles Rupee"
                .AddItem "SLL-Sierra Leone Leone"
                .AddItem "XAG-Silver Ounces"
                .AddItem "SGD-Singapore Dollar"
                .AddItem "SKK-Slovak Koruna"
                .AddItem "SIT-Slovenian Tolar"
                .AddItem "SBD-Solomon Islands Dollar"
                .AddItem "SOS-Somali Shilling"
                .AddItem "ZAR-South African Rand"
                .AddItem "LKR-Sri Lanka Rupee"
                .AddItem "SHP-St Helena Pound"
                .AddItem "SDD-Sudanese Dinar"
                .AddItem "SRG-Surinam Guilder"
                .AddItem "SZL-Swaziland Lilageni"
                .AddItem "SEK-Swedish Krona"
                .AddItem "TRY-Turkey Lira"
                .AddItem "CHF-Swiss Franc"
                .AddItem "SYP-Syrian Pound"
                .AddItem "TWD-Taiwan Dollar"
                .AddItem "TZS-Tanzanian Shilling"
                .AddItem "THB-Thai Baht"
                .AddItem "TOP-Tonga Pa'anga"
                .AddItem "TTD-Trinidad and Tobago Dollar"
                .AddItem "TND-Tunisian Dinar"
                .AddItem "TRL-Turkish Lira"
                .AddItem "USD -U.S.Dollar"
                .AddItem "AED-UAE Dirham"
                .AddItem "UGX-Ugandan Shilling"
                .AddItem "UAH-Ukraine Hryvnia"
                .AddItem "UYU-Uruguayan New Peso"
                .AddItem "VUV-Vanuatu Vatu"
                .AddItem "VEB-Venezuelan Bolivar"
                .AddItem "VND-Vietnam Dong"
                .AddItem "YER-Yemen Riyal"
                .AddItem "YUM-Yugoslav Dinar"
                .AddItem "ZMK-Zambian Kwacha"
                .AddItem "ZWD-Zimbabwe Dollar"
            End With
    End Select
    
End Function

