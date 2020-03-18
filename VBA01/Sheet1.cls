VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButton1_Click()
Range("P3").Value = Range("P3").Value + 1
Range("s_no").ClearContents
Range("product_description").ClearContents
Range("hsn").ClearContents
Range("qty").ClearContents
Range("rate").ClearContents
Range("discount").ClearContents
Range("cusd").ClearContents
End Sub

Private Sub CommandButton2_Click()
ActiveSheet.PrintOut
Application.SendKeys ("^s")
End Sub

Private Sub CommandButton3_Click()
 
Dim fName As String
With ActiveSheet
    fName = .Range("invoice_number").Value & " - " & .Range("customer_name").Value & " ( Rs." & .Range("total_amount").Value & ") "
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            "E:\invoice\bills\INVNO -" & fName, Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End With



 
Application.SendKeys ("^s")
 
End Sub

Private Sub CommandButton4_Click()
CommandButton2_Click
CommandButton3_Click
Application.SendKeys ("^s")
End Sub

Private Sub CommandButton5_Click()
Application.SendKeys ("^p")
End Sub

Private Sub CommandButton6_Click()
Range("s_no").ClearContents
Range("product_description").ClearContents
Range("hsn").ClearContents
Range("qty").ClearContents
Range("rate").ClearContents
Range("discount").ClearContents
Range("cusd").ClearContents
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub

