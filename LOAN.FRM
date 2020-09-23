VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form LoanSheet 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LoanSheet"
   ClientHeight    =   6975
   ClientLeft      =   1230
   ClientTop       =   1845
   ClientWidth     =   9915
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6975
   ScaleWidth      =   9915
   Begin VB.CheckBox Check1 
      Caption         =   "Pay Extra Amount"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   1260
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Height          =   1875
      Left            =   120
      TabIndex        =   19
      Top             =   1260
      Width           =   2775
      Begin VB.TextBox txtExtraPay 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1920
         TabIndex        =   23
         Text            =   "1500"
         Top             =   1440
         Width           =   675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pay fixed amount over P + I"
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Value           =   -1  'True
         Width           =   2235
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Extra Payment each month"
         Height          =   435
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Pay Fixed Amount:"
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   1500
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info"
      Height          =   1215
      Left            =   3120
      TabIndex        =   18
      Top             =   5700
      Width           =   6735
      Begin VB.Label lblinfo 
         Caption         =   "lblinfo"
         Height          =   855
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   6375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdPayments 
      Height          =   5115
      Left            =   3060
      TabIndex        =   14
      Top             =   420
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9022
      _Version        =   393216
      Rows            =   50
      Cols            =   6
      WordWrap        =   -1  'True
      AllowUserResizing=   1
      FormatString    =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCalcAmort 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Show Amortization"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton cmdCalcPmts 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Show Payments"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   5985
      Width           =   2295
   End
   Begin VB.Frame frmLoanLen 
      Caption         =   "Years in Loan"
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   3180
      Width           =   2775
      Begin VB.TextBox txtLen1 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtLen2 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblLen 
         Caption         =   "Length 1:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   390
         Width           =   855
      End
      Begin VB.Label lblLen 
         Caption         =   "Length 2:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.TextBox txtDwnPay 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Text            =   "10"
      Top             =   720
      Width           =   1395
   End
   Begin VB.Frame frmIntr 
      Caption         =   "Interest Rates (%)"
      Height          =   1275
      Left            =   120
      TabIndex        =   6
      Top             =   4620
      Width           =   2775
      Begin VB.TextBox txtRate2 
         Height          =   285
         Left            =   1380
         TabIndex        =   17
         Text            =   "8.0"
         Top             =   840
         Width           =   675
      End
      Begin VB.TextBox txtRate1 
         Height          =   315
         Left            =   1380
         TabIndex        =   16
         Text            =   "6.5"
         Top             =   360
         Width           =   675
      End
      Begin VB.Label lblIntr 
         Caption         =   "Rate 2:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   885
         Width           =   1095
      End
      Begin VB.Label lblIntr 
         Caption         =   "Rate 1:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   405
         Width           =   1095
      End
   End
   Begin VB.TextBox txtPurchTotl 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Text            =   "180000"
      Top             =   120
      Width           =   1395
   End
   Begin VB.Label Label2 
      Height          =   315
      Left            =   3060
      TabIndex        =   15
      Top             =   60
      Width           =   6615
   End
   Begin VB.Label lblDwnPay 
      Caption         =   "Percent Down"
      Height          =   435
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   900
   End
   Begin VB.Label lblPurchTotl 
      Caption         =   "Purchase Amount"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   900
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptInstruct 
         Caption         =   "&Instructions..."
      End
      Begin VB.Menu mnuOptSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuDwnPay 
      Caption         =   "&Down Payment"
      Begin VB.Menu mnuOptDwn 
         Caption         =   "&Percent Down"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuOptDwn 
         Caption         =   "&Amount Down"
         Index           =   1
      End
   End
   Begin VB.Menu mnuLoanLen 
      Caption         =   "&Loan Length"
      Begin VB.Menu mnuOptLen 
         Caption         =   "&Years"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuOptLen 
         Caption         =   "&Months"
         Index           =   1
      End
   End
End
Attribute VB_Name = "LoanSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Form variables for purchase/down payment
Dim PurchAmt As Double
Dim DPAmt As Double
' Form variable for loan amount
Dim LoanAmt As Double

'Array for payments
Dim Payments(1 To 2, 1 To 2) As Double

' Array for amortization information
Dim AmortData() As Double
' Form variable to hold payment for amortization calculation
Dim AmortPmnt As Double
' Form variable flag telling if the grid shows amortization
Dim AmortFlag

'New variables added
Dim Rates(1 To 2) As Double
Dim Length(1 To 2) As Double
Dim ExtraPay As Double
Private Sub CalcAmort()
    ' Declare local variables
    Dim i As Integer
    Dim IntrTotl As Double
    Dim Months As Integer
    Dim Interest As Double
    Dim TotalInt As Double, TotalIntExtraPay As Double
    
    ' Save monthly payment from selected cell
    AmortPmnt = Payments(grdPayments.Row, grdPayments.Col)
    ExtraPay = Val(txtExtraPay)

    If Check1.Value <> 1 Then
      ExtraPay = 0
    Else
      If Option1(1).Value = True Then
        If AmortPmnt > ExtraPay Then
          MsgBox "The amount to pay must be greater than the initial principal and Interest"
          txtExtraPay.SetFocus
          Exit Sub
        End If
      End If
    End If
        
    ' Save interest rate from selected row
    Interest = Rates(grdPayments.Row)
    
    ' If the periods = years
    If mnuOptLen(0).Checked = True Then
        ' Make loan length loan match selected cell
        Months = Length(grdPayments.Col) * 12
    Else
        Months = Length(grdPayments.Col)
    End If
    
    ' Size array holding amortization data
    ReDim AmortData(1 To Months, 1 To 5)
    ' Save original loan amount to calculate remaining principal
    AmortData(1, 1) = LoanAmt
    ' For each month in loan period
    
    
    'Go through calculation once to get the total interest paid, if an extra payment is
    'being made
    TotalInt = 0
    
    For i = 2 To Months
      ' Calculate interest paid for current month
      AmortData(i - 1, 2) = AmortData(i - 1, 1) * Interest / 1200
      TotalInt = TotalInt + AmortData(i - 1, 2)
      ' Calculate remaining balance of principal and subtract the extra payment
      AmortData(i, 1) = AmortData(i - 1, 1) - AmortPmnt + AmortData(i - 1, 2)
    Next i
    
    If ExtraPay <> 0 Then
      For i = 2 To Months
        ' Calculate interest paid for current month
        AmortData(i - 1, 2) = AmortData(i - 1, 1) * Interest / 1200
        TotalIntExtraPay = TotalIntExtraPay + AmortData(i - 1, 2)
        
        ' Calculate remaining balance of principal and subtract the extra payment
        If AmortData(i - 1, 1) > AmortPmnt Then
          If Option1(1).Value And ExtraPay <> 0 Then
            AmortData(i, 1) = AmortData(i - 1, 1) - AmortPmnt + AmortData(i - 1, 2) - (ExtraPay - AmortPmnt)
          Else
            AmortData(i, 1) = AmortData(i - 1, 1) - AmortPmnt + AmortData(i - 1, 2) - ExtraPay
          End If
        Else
          AmortData(i, 1) = 0   'The loan has been paid off
        End If
        
      Next i
    End If
    
    'Show the results
    grdPayments.Clear
    
    ' Set # of non-fixed cols (4)
    grdPayments.Cols = 6
    grdPayments.TextMatrix(0, 0) = "P + I " & Format$(AmortPmnt, "Currency")
    grdPayments.ColWidth(0) = 1.1 * TextWidth("000000")
            
    ' Set remaining col headings
    grdPayments.TextMatrix(0, 1) = "Principal Paid"
    grdPayments.TextMatrix(0, 2) = "Interest Paid"
    grdPayments.TextMatrix(0, 3) = "Principal Balance"
    grdPayments.TextMatrix(0, 4) = "Total Interest to Date"
    If Option1(0).Value Then
      grdPayments.TextMatrix(0, 5) = "Extra Payment"
    Else
      grdPayments.TextMatrix(0, 5) = "Fixed Payment above P+I"
    End If
    grdPayments.RowHeight(0) = 500
        
    ' Display vert scroll bar
    grdPayments.ScrollBars = 3

    ' Set # of non-fixed rows
    If Months > 16 Then
        grdPayments.Rows = Months + 1
    Else
        grdPayments.Rows = 16
    End If
    
    
    For i = 1 To Months
        grdPayments.TextMatrix(i, 0) = i
    Next i
    
    For i = 1 To 5
       grdPayments.ColWidth(i) = 1.1 * TextWidth("0000000.00")
    Next i
    
    ' Set initial interest total
    IntrTotl = 0
    ' Put values of non-fixed rows in grid
    For i = 1 To Months
      If AmortData(i, 1) > 0 Then
        'Display part of payment applied to principal
        grdPayments.TextMatrix(i, 1) = Format$(AmortPmnt - AmortData(i, 2), "Currency")
        ' Display interest paid this payment in 2nd non-fixed col
        grdPayments.TextMatrix(i, 2) = Format$(AmortData(i, 2), "Currency")
        ' Display balance remaining on principal in 3rd non-fixed col
        grdPayments.TextMatrix(i, 3) = Format$(AmortData(i, 1), "Currency")
        ' Display total interest paid to date in last non-fixed col
        IntrTotl = IntrTotl + AmortData(i, 2)
        grdPayments.TextMatrix(i, 4) = Format$(IntrTotl, "Currency")
        grdPayments.TextMatrix(i, 5) = Format$(ExtraPay, "Currency")
      Else
        grdPayments.TextMatrix(i, 1) = "Paid Off"
      End If
    Next i
    ' Disable Calculate Amortization button
    cmdCalcAmort.Enabled = False
    ' Enable amortization flag
    AmortFlag = True
    
    lblinfo = "Loan Amount = " & Format(LoanAmt, "Currency") & Chr(13) & Chr(10)
    lblinfo = lblinfo & "Total Interest Paid = " & Format(TotalInt, "Currency") & " (no extra payment)" & Chr(13) & Chr(10)
    If ExtraPay <> 0 Then
      lblinfo = lblinfo & "Total Interest Paid = " & Format(TotalIntExtraPay, "Currency") & " (with extra payment)" & Chr(13) & Chr(10)
      lblinfo = lblinfo & "Amount saved with Extra Payment = " & Format(TotalInt - TotalIntExtraPay, "Currency") & Chr(13) & Chr(10)
    End If
    
        
    
End Sub

Private Sub CalcPmnts()
    ' Declare local variables
    Dim i As Integer
    Dim j As Integer
    Dim LengthUnit As String
       
    'calculate payments
    If mnuOptLen(0).Checked = True Then
      LengthUnit = "Years"
    Else
      LengthUnit = "Months"
    End If
        
    For i = 1 To 2
      For j = 1 To 2
        If Rates(i) <> 0 And Length(j) <> 0 Then
          If mnuOptLen(0).Checked = True Then 'Years
            Payments(i, j) = MonthPay(Rates(i) / 1200, Length(j) * 12)
          Else
            Payments(i, j) = MonthPay(Rates(i) / 1200, Length(j))
          End If
        Else
          Payments(i, j) = 0
        End If
      Next j
    Next i
          
   'Now show the payments on the grid
      
     grdPayments.Clear
      
    ' Initialize line break variable
    Label2 = "Monthly payments, not including extra payment"
       
    ' Put col headings in 1st row
    For i = 1 To 2
       grdPayments.TextMatrix(0, i) = Length(i) & " " & LengthUnit
    Next i
    ' Center text in 1st col
    grdPayments.FixedAlignment(0) = 2
    ' Put row headings in 1st col
    For i = 1 To 2
        grdPayments.TextMatrix(i, 0) = Format(Rates(i), "0.000") & "%"
    Next i
    
    ' Fill grid with payments from array
    For i = 1 To 2
       For j = 1 To 2
          grdPayments.ColAlignment(j) = 2
          If Payments(i, j) <> 0 Then
            grdPayments.TextMatrix(i, j) = Format$(Payments(i, j), "Currency")
          Else
            grdPayments.TextMatrix(i, j) = "N/A"
          End If
        Next j
    Next i
      
     ' Set amortization flag to false
    AmortFlag = False
    grdPayments.Row = 1
   

End Sub





Private Sub Check1_Click()


 Frame2.Enabled = (Check1.Value = 1)
 Option1(0).Enabled = (Check1.Value = 1)
 Option1(1).Enabled = (Check1.Value = 1)
 txtExtraPay.Enabled = (Check1.Value = 1)
 Label1.Enabled = (Check1.Value = 1)
 

End Sub

Private Sub cmdCalcAmort_Click()
    ' Change mouse pointer to hourglass
    MousePointer = 11
    ' Calculate amortization
    CalcAmort
    ' Return mouse pointer to default
    MousePointer = 0
End Sub

Private Sub cmdCalcPmts_Click()
    ' Disable amortization button
    cmdCalcAmort.Enabled = False
    
    Rates(1) = Val(txtRate1)
    Rates(2) = Val(txtRate2)
    Length(1) = Val(Me.txtLen1)
    Length(2) = Val(Me.txtLen2)
    ExtraPay = Val(txtExtraPay)
    PurchAmt = Val(txtPurchTotl)
    DPAmt = Val(txtDwnPay)
    
    ' Does Purchase Amount have a value?
    If PurchAmt <= 0 Then
        MsgBox "You must enter a valid purchase amount.", 48, "LoanSheet Error"
        txtPurchTotl.Text = ""
        txtPurchTotl.SetFocus
        Exit Sub
    End If
     
          
    ' If maximum interest rate is less than minimum interest rate...
    If Length(1) = 0 And Length(2) = 0 Then
        MsgBox "You must enter at least one valid length for the loan.", 48, "LoanSheet Error"
        txtRate1.SetFocus
        Exit Sub
    End If
     
     
    ' If maximum interest rate is less than minimum interest rate...
    If Rates(1) = 0 And Rates(2) = 0 Then
        MsgBox "You must enter at least one valid interest rate.", 48, "LoanSheet Error"
        txtRate1.SetFocus
        Exit Sub
    End If

    ' Calculate loan amount
    If mnuOptDwn(0).Checked = True Then
        LoanAmt = PurchAmt * (1 - (DPAmt / 100))
    ElseIf mnuOptDwn(1).Checked = True Then
        LoanAmt = PurchAmt - DPAmt
    End If
    ' Change mouse pointer to hourglass
    MousePointer = 11
    ' Calculate payments for all loan lengths/interest rates
    CalcPmnts
    ' Change mouse pointer to default
    lblinfo = "Loan Amount = " & Format(LoanAmt, "Currency")
    
    MousePointer = 0
End Sub



Private Sub Form_Load()
    ' Declare local variables
    Dim iRate As Double
    
    ' Set down payment to 0
    DPAmt = 0
    ' Put app in center of screen
    LoanSheet.Left = (Screen.Width - LoanSheet.Width) / 2
    LoanSheet.Top = (Screen.Height - LoanSheet.Height) / 2
    ' Put allowed interest rates in list boxes
      
    txtLen1 = 30
    txtLen2 = 15
    txtPurchTotl = 180000
    
    lblinfo = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub grdPayments_Click()
    ' Only act if cell contains a payment
    If grdPayments.Text = "" Then
      cmdCalcAmort.Enabled = False
      Exit Sub
    End If
    
    If grdPayments.Text <> "N/A" And AmortFlag = False Then
        ' Enable amortization button
        cmdCalcAmort.Enabled = True
    Else
        ' Disable amortization button
        cmdCalcAmort.Enabled = False
    End If
End Sub

Private Sub grdPayments_DblClick()
    ' Only act if cell contains a payment
    If grdPayments.Text <> "" And AmortFlag = False Then
        ' Enable amortization button
        cmdCalcAmort.Enabled = True
        ' Change mouse pointer to hourglass
        MousePointer = 11
        ' Calculate amortization
        CalcAmort
         ' Change mouse pointer to default
        MousePointer = 0
    Else
        ' Disable amortization button
        cmdCalcAmort.Enabled = False
    End If
End Sub



Private Sub mnuOptDwn_Click(Index As Integer)
    Select Case Index
        Case 0
            mnuOptDwn(0).Checked = True
            mnuOptDwn(1).Checked = False
            lblDwnPay.Caption = "Percent Down"
        Case 1
            mnuOptDwn(0).Checked = False
            mnuOptDwn(1).Checked = True
            lblDwnPay.Caption = "Amount Down"
    End Select
End Sub

Private Sub mnuOptExit_Click()
    End
End Sub

Private Sub mnuOptInstruct_Click()
    ' Declare local variables
    Dim MsgText
    Dim PB
    ' Initialize paragraph break variable
    PB = Chr$(10) & Chr$(13) & Chr$(10) & Chr$(13)
    ' Display message
    MsgText = "Enter purchase amount, down payment, length of loan, and interest rates. Click the Calculate Payments button to display monthly payments in the grid on the right."
    MsgText = MsgText & PB & "Select a monthly payment and click the Calculate Amortization button to display an amortization schedule for the selected interest rate and length of loan in the grid."
    MsgText = MsgText & PB & "Highlight cells in the grid and select Copy from the Options menu to copy data to the Clipboard."
    MsgBox MsgText, 64, "LoanSheet Instructions"
End Sub

Private Sub mnuOptLen_Click(Index As Integer)
    Select Case Index
        ' Loan length = years
        Case 0
            mnuOptLen(0).Checked = True
            mnuOptLen(1).Checked = False
            frmLoanLen.Caption = "Years in Loan"
        ' Loan length = months
        Case 1
            mnuOptLen(0).Checked = False
            mnuOptLen(1).Checked = True
            frmLoanLen.Caption = "Months in Loan"
    End Select
End Sub

Private Function MonthPay(Interest, Months)
    MonthPay = LoanAmt * (Interest / (1 - (1 / ((1 + Interest) ^ Months))))
End Function







Private Sub txtDwnPay_Change()
    ' Store number into form variable
    DPAmt = Val(txtDwnPay.Text)
    ' Validate that DPAmt doesn't exceed 100% or Purchase Amount
    If mnuOptDwn(0).Checked = True And DPAmt > 99 Then
        MsgBox "Down payment percentage cannot exceed 99 percent.", 48, "LoanSheet Error"
        txtDwnPay.Text = ""
        txtDwnPay.SetFocus
    ElseIf mnuOptDwn(1).Checked = True And DPAmt > PurchAmt Then
        MsgBox "Down payment amount cannot exceed purchase amount.", 48, "LoanSheet Error"
        txtDwnPay.Text = ""
        txtDwnPay.SetFocus
    ElseIf DPAmt < 0 Then
        MsgBox "Down payment amount must be zero or greater.", 48, "LoanSheet Error"
        txtDwnPay.Text = ""
        txtDwnPay.SetFocus
    End If
End Sub





Private Sub txtPurchTotl_Change()
    PurchAmt = Val(txtPurchTotl.Text)
End Sub

Private Sub txtRate1_Change()
Rates(1) = Val(txtRate1)
End Sub


Private Sub txtRate2_Change()
Rates(2) = Val(txtRate2)
End Sub


