Attribute VB_Name = "Module1"
Option Explicit

Public state As String
Public statecode As String
Public FZNo As String
Public FZClass As String
Public FZ3X5PHPN As String
Public FZ3X5PHPrice As String
Public FZ3X5PHFPN As String
Public FZ3X5PHFPrice As String
Public FZ4X6PHPN As String
Public FZ4X6PHPrice As String
Public FZ4X6PHFPN As String
Public FZ4X6PHFPrice As String
Public FZ7FtHWPN As String
Public FZ7FtHWPrice As String
Public FZ8FtHWPN As String
Public FZ8FtHWPrice As String
Public FZ9FtHWPN As String
Public FZ9FtHWPrice As String
Public BaseSKU As String

Public PRSETPH357SKU As String
Public PRSETPH357MPN As String
Public PRSETPH357Price As String
Public PRSETPH357Cost As String

Public PRSETPHF357SKU As String
Public PRSETPHF357MPN As String
Public PRSETPHF357Price As String
Public PRSETPHF357Cost As String

Public PRSETPH358SKU As String
Public PRSETPH358MPN As String
Public PRSETPH358Price As String
Public PRSETPH358Cost As String

Public PRSETPHF358SKU As String
Public PRSETPHF358MPN As String
Public PRSETPHF358Price As String
Public PRSETPHF358Cost As String

Public PRSETPH469SKU As String
Public PRSETPH469MPN As String
Public PRSETPH469Price As String
Public PRSETPH469Cost As String

Public PRSETPHF469SKU As String
Public PRSETPHF469MPN As String
Public PRSETPHF469Price As String
Public PRSETPHF469Cost As String

Public PRSETPHImage As String
Public PRSETPHFImage As String
Public PHFFlagImage As String
Public HWImage As String
Public ImageDesc As String
Public PageTitle As String
Public Category As String

Sub RunThis()
Dim sourcebook As Workbook
Set sourcebook = Workbooks.Open("S:\Web Site Files\Flags\State Flags\Presentation Sets\State_Flag_Presentation_Set_Worksheet.xlsx")
Dim sourcesheet As Worksheet
Set sourcesheet = sourcebook.Worksheets("Sheet1")
Dim newwb As Workbook
Set newwb = Workbooks.Add
newwb.SaveAs Filename:="S:\Web Site Files\Flags\State Flags\Presentation Sets\VBA\Test1.xlsx", FileFormat:=xlWorkbookDefault
Dim newsheet As Worksheet
Set newsheet = newwb.ActiveSheet

Call WriteHeaderRow(newsheet)
' process input sheet, rows 2 through 57
Dim i As Long  ' index into state list 1 = first (on row 2)
For i = 1 To 57
  DoEvents ' give Windows a moment for housekeeping
  Call ReadStateRow(sourcesheet, i + 1)  ' first state on row 2 of input sheet
  ' for the following, pass first row of each set of 13 rows per product
  Call WriteProductRow(newsheet, ((i - 1) * 13) + 2) ' first state to row 2, next to row 15, etc.
  Call FlagSKUs(newsheet, ((i - 1) * 13) + 2)
  Call FlagSKURules(newsheet, ((i - 1) * 13) + 2)
Next
newwb.Save
' clean up
Set newsheet = Nothing
newwb.Close
Set newwb = Nothing
sourcebook.Close
Set sourcebook = Nothing
MsgBox "Done"
End Sub

Sub WriteHeaderRow(ns As Worksheet)
With ns
  .Cells(1, 1) = "Item Type"
  .Cells(1, 2) = "Product ID"
  .Cells(1, 3) = "Sort Order"
  .Cells(1, 4) = "Product Name"
  .Cells(1, 5) = "Product Type"
  .Cells(1, 6) = "Product Code/SKU"
  .Cells(1, 7) = "Bin Picking Number"
  .Cells(1, 8) = "Origin Locations"
  .Cells(1, 9) = "Shipping Groups"
  .Cells(1, 10) = "Dimensional Rules"
  .Cells(1, 11) = "Brand Name"
  .Cells(1, 12) = "Option Set"
  .Cells(1, 13) = "Option Set Align"
  .Cells(1, 14) = "Product Description"
  .Cells(1, 15) = "Price"
  .Cells(1, 16) = "Cost Price"
  .Cells(1, 17) = "Retail Price"
  .Cells(1, 18) = "Sale Price"
  .Cells(1, 19) = "Fixed Shipping Cost"
  .Cells(1, 20) = "Free Shipping"
  .Cells(1, 21) = "Product Warranty"
  .Cells(1, 22) = "Product Weight"
  .Cells(1, 23) = "Product Width"
  .Cells(1, 24) = "Product Height"
  .Cells(1, 25) = "Product Depth"
  .Cells(1, 26) = "Allow Purchases?"
  .Cells(1, 27) = "Product Visible?"
  .Cells(1, 28) = "Product Availability"
  .Cells(1, 29) = "Track Inventory"
  .Cells(1, 30) = "Current Stock Level"
  .Cells(1, 31) = "Low Stock Level"
  .Cells(1, 32) = "Category"
  .Cells(1, 33) = "Product Image File - 1"
  .Cells(1, 34) = "Product Image URL - 1"
  .Cells(1, 35) = "Product Image ID - 1"
  .Cells(1, 36) = "Product Image File - 1"
  .Cells(1, 37) = "Product Image Description - 1"
  .Cells(1, 38) = "Product Image Is Thumbnail - 1"
  .Cells(1, 39) = "Product Image Sort - 1"
  .Cells(1, 40) = "Product Image File - 2"
  .Cells(1, 41) = "Product Image URL - 2"
  .Cells(1, 42) = "Product Image ID - 2"
  .Cells(1, 43) = "Product Image File - 2"
  .Cells(1, 44) = "Product Image Description - 2"
  .Cells(1, 45) = "Product Image Is Thumbnail - 2"
  .Cells(1, 46) = "Product Image Sort - 2"
  .Cells(1, 47) = "Product Image File - 3"
  .Cells(1, 48) = "Product Image URL - 3"
  .Cells(1, 49) = "Product Image ID - 3"
  .Cells(1, 50) = "Product Image File - 3"
  .Cells(1, 51) = "Product Image Description - 3"
  .Cells(1, 52) = "Product Image Is Thumbnail - 3"
  .Cells(1, 53) = "Product Image Sort - 3"
  .Cells(1, 54) = "Search Keywords"
  .Cells(1, 55) = "Page Title"
  .Cells(1, 56) = "Meta Keywords"
  .Cells(1, 57) = "Meta Description"
  .Cells(1, 58) = "Product Condition"
  .Cells(1, 59) = "Show Product Condition?"
  .Cells(1, 60) = "Product Tax Class"
  .Cells(1, 61) = "Manufacturer Part Number"
  .Cells(1, 62) = "Product UPC/EAN"
  .Cells(1, 63) = "Product URL"
  .Cells(1, 64) = "Redirect Old URL?"
  .Cells(1, 65) = "GPS Global Trade Item Number"
  .Cells(1, 66) = "GPS Color"
  .Cells(1, 67) = "GPS Item Group ID"
  .Cells(1, 68) = "GPS Category"
  .Cells(1, 69) = "Product Custom Fields"
End With
End Sub

Sub ReadStateRow(ss As Worksheet, rownumber As Long)
With ss
  state = .Cells(rownumber, 1)
  statecode = .Cells(rownumber, 2)
  FZNo = .Cells(rownumber, 3)
  FZClass = .Cells(rownumber, 4)
  FZ3X5PHPN = .Cells(rownumber, 5)
  FZ3X5PHPrice = .Cells(rownumber, 6)
  FZ3X5PHFPN = .Cells(rownumber, 7)
  FZ3X5PHFPrice = .Cells(rownumber, 8)
  FZ4X6PHPN = .Cells(rownumber, 9)
  FZ4X6PHPrice = .Cells(rownumber, 10)
  FZ4X6PHFPN = .Cells(rownumber, 11)
  FZ4X6PHFPrice = .Cells(rownumber, 12)
  FZ7FtHWPN = .Cells(rownumber, 13)
  FZ7FtHWPrice = .Cells(rownumber, 14)
  FZ8FtHWPN = .Cells(rownumber, 15)
  FZ8FtHWPrice = .Cells(rownumber, 16)
  FZ9FtHWPN = .Cells(rownumber, 17)
  FZ9FtHWPrice = .Cells(rownumber, 18)
  BaseSKU = .Cells(rownumber, 19)
  PRSETPH357SKU = .Cells(rownumber, 20)
  PRSETPH357Cost = .Cells(rownumber, 21)
  PRSETPH357Price = .Cells(rownumber, 22)
  PRSETPH357MPN = .Cells(rownumber, 23)
  PRSETPHF357SKU = .Cells(rownumber, 24)
  PRSETPHF357Cost = .Cells(rownumber, 25)
  PRSETPHF357Price = .Cells(rownumber, 26)
  PRSETPHF357MPN = .Cells(rownumber, 27)
  PRSETPH358SKU = .Cells(rownumber, 28)
  PRSETPH358Cost = .Cells(rownumber, 29)
  PRSETPH358Price = .Cells(rownumber, 30)
  PRSETPH358MPN = .Cells(rownumber, 31)
  PRSETPHF358SKU = .Cells(rownumber, 32)
  PRSETPHF358Cost = .Cells(rownumber, 33)
  PRSETPHF358Price = .Cells(rownumber, 34)
  PRSETPHF358MPN = .Cells(rownumber, 35)
  PRSETPH469SKU = .Cells(rownumber, 36)
  PRSETPH469Cost = .Cells(rownumber, 37)
  PRSETPH469Price = .Cells(rownumber, 38)
  PRSETPH469MPN = .Cells(rownumber, 39)
  PRSETPHF469SKU = .Cells(rownumber, 40)
  PRSETPHF469Cost = .Cells(rownumber, 41)
  PRSETPHF469Price = .Cells(rownumber, 42)
  PRSETPHF469MPN = .Cells(rownumber, 43)
  PRSETPHImage = .Cells(rownumber, 44)
  PRSETPHFImage = .Cells(rownumber, 45)
  PHFFlagImage = .Cells(rownumber, 46)
  HWImage = .Cells(rownumber, 47)
  ImageDesc = .Cells(rownumber, 48)
  PageTitle = .Cells(rownumber, 49)
  Category = .Cells(rownumber, 50)
End With
End Sub

Sub WriteProductRow(ns As Worksheet, ra As Long)
' write empty strings to all cells in data area
Dim i As Long
With ns
  For i = 1 To 69
    .Cells(ra, i) = ""
  Next
  DoEvents ' give Windows a moment for housekeeping
' write necessary data
  .Cells(ra, 1) = "Product"
  .Cells(ra, 3) = "0"
  .Cells(ra, 4) = state + " Deluxe Indoor Presentation Set with Oak Pole, Gold Base and Hardware (Open Market)"
  .Cells(ra, 5) = "P"
  .Cells(ra, 6) = BaseSKU
  .Cells(ra, 12) = "Oak Presentation Set Options"
  .Cells(ra, 13) = "Right"
  .Cells(ra, 15) = "0.00"
  .Cells(ra, 16) = "0.00"
  .Cells(ra, 17) = "0.00"
  .Cells(ra, 18) = "0.00"
  .Cells(ra, 19) = "0"
  .Cells(ra, 20) = "N"
  .Cells(ra, 22) = "25"
  .Cells(ra, 23) = "53"
  .Cells(ra, 24) = "13"
  .Cells(ra, 25) = "6"
  .Cells(ra, 26) = "Y"
  .Cells(ra, 27) = "Y"
  .Cells(ra, 29) = "none"
  .Cells(ra, 30) = "0"
  .Cells(ra, 31) = "0"
  .Cells(ra, 32) = "Flags/State & U.S. Territory Flags/" + state + " Flags/Indoor" + state + " Flags;Flagpoles/Indoor Flagpoles/Presentation Flagpole Sets/State Indoor & Parade Set"
  .Cells(ra, 33) = "https://emflag.com/content/Flags/State/Indoor/Presentation%20Sets/PRSET-" + state + "-PHF.png"
  .Cells(ra, 34) = "https://emflag.com/content/Flags/State/Indoor/Presentation%20Sets/PRSET-" + state + "-PHF.png"
  .Cells(ra, 37) = "Deluxe Indoor and Parade " + state + "Flag Presentation Set"
  .Cells(ra, 38) = "Y"
  .Cells(ra, 39) = "1"
  .Cells(ra, 54) = "Deluxe Indoor and Parade " + state + "Flag Presentation Set"
  .Cells(ra, 55) = "Shop " + state + " Indoor and Parade Presentation Set - Made in USA"
  .Cells(ra, 56) = "Meta Keywords"
  .Cells(ra, 57) = "Meta Description"
  .Cells(ra, 58) = "New"
  .Cells(ra, 59) = "N"
  .Cells(ra, 60) = "Default Tax Class"
End With
End Sub

Sub FlagSKUs(ns As Worksheet, ra As Long)  ' ra is first row of each set of product data
' write empty strings to all cells in data area
Dim i As Long, j As Long
With ns
  For j = 1 To 69
    For i = ra + 1 To ra + 6
      .Cells(i, j) = ""
    Next
  Next
  DoEvents ' give Windows a moment for housekeeping

'12inX18inHG SKU - product row 2
  ra = ra + 1
  .Cells(ra, 1) = "SKU"
  .Cells(ra, 4) = "[RT]Finishing Options=Pole Hem Only,[RT]Flag and Pole Size=3' X 5' Flag with 7' Oak Pole"
  .Cells(ra, 6) = PRSETPH357SKU
  .Cells(ra, 7) = "FlagZone"
  .Cells(ra, 8) = "FlagZone"
  .Cells(ra, 11) = "FlagZone"
  .Cells(ra, 16) = PRSETPH357Cost
  .Cells(ra, 20) = "N"
  .Cells(ra, 22) = "18"
  .Cells(ra, 23) = "53"
  .Cells(ra, 24) = "13"
  .Cells(ra, 25) = "6"
  .Cells(ra, 61) = PRSETPH357MPN

'2X3 HG SKU  - product row 3
  ra = ra + 1
  .Cells(ra, 1) = "SKU"
  .Cells(ra, 4) = "[RT]Finishing Options=Pole Hem & Fringe,[RT]Flag and Pole Size=3' X 5' Flag with 7' Oak Pole"
  .Cells(ra, 6) = PRSETPHF357SKU
  .Cells(ra, 7) = "FlagZone"
  .Cells(ra, 8) = "FlagZone"
  .Cells(ra, 11) = "FlagZone"
  .Cells(ra, 16) = PRSETPHF357Cost
  .Cells(ra, 20) = "N"
  .Cells(ra, 22) = "18"
  .Cells(ra, 23) = "53"
  .Cells(ra, 24) = "13"
  .Cells(ra, 25) = "6"
  .Cells(ra, 61) = PRSETPHF357MPN

'3X5 HG SKU  - product row 4
  ra = ra + 1
  .Cells(ra, 1) = "SKU"
  .Cells(ra, 4) = "[RT]Finishing Options=Pole Hem Only,[RT]Flag and Pole Size=3' X 5' Flag with 8' Oak Pole"
  .Cells(ra, 6) = PRSETPH358SKU
  .Cells(ra, 7) = "FlagZone"
  .Cells(ra, 8) = "FlagZone"
  .Cells(ra, 11) = "FlagZone"
  .Cells(ra, 16) = PRSETPH358Cost
  .Cells(ra, 20) = "N"
  .Cells(ra, 22) = "18"
  .Cells(ra, 23) = "53"
  .Cells(ra, 24) = "13"
  .Cells(ra, 25) = "6"
  .Cells(ra, 61) = PRSETPH358MPN

'4X6 HG SKU  - product row 5
  ra = ra + 1
  .Cells(ra, 1) = "SKU"
  .Cells(ra, 4) = "[RT]Finishing Options=Pole Hem & Fringe,[RT]Flag and Pole Size=3' X 5' Flag with 8' Oak Pole"
  .Cells(ra, 6) = PRSETPHF358SKU
  .Cells(ra, 7) = "FlagZone"
  .Cells(ra, 8) = "FlagZone"
  .Cells(ra, 11) = "FlagZone"
  .Cells(ra, 16) = PRSETPHF358Cost
  .Cells(ra, 20) = "N"
  .Cells(ra, 22) = "18"
  .Cells(ra, 23) = "53"
  .Cells(ra, 24) = "13"
  .Cells(ra, 25) = "6"
  .Cells(ra, 61) = PRSETPHF358MPN

'5X8 HG SKU  - product row 6
  ra = ra + 1
  .Cells(ra, 1) = "SKU"
  .Cells(ra, 4) = "[RT]Finishing Options=Pole Hem Only,[RT]Flag and Pole Size=4' X 6' Flag with 9' Oak Pole"
  .Cells(ra, 6) = PRSETPH469SKU
  .Cells(ra, 7) = "FlagZone"
  .Cells(ra, 8) = "FlagZone"
  .Cells(ra, 11) = "FlagZone"
  .Cells(ra, 16) = PRSETPH469Cost
  .Cells(ra, 20) = "N"
  .Cells(ra, 22) = "25"
  .Cells(ra, 23) = "53"
  .Cells(ra, 24) = "13"
  .Cells(ra, 25) = "6"
  .Cells(ra, 61) = PRSETPH469MPN

'6X10 HG SKU  - product row 7
  ra = ra + 1
  .Cells(ra, 1) = "SKU"
  .Cells(ra, 4) = "[RT]Finishing Options=Pole Hem & Fringe,[RT]Flag and Pole Size=4' X 6' Flag with 9' Oak Pole"
  .Cells(ra, 6) = PRSETPHF469SKU
  .Cells(ra, 7) = "FlagZone"
  .Cells(ra, 8) = "FlagZone"
  .Cells(ra, 11) = "FlagZone"
  .Cells(ra, 16) = PRSETPHF469Cost
  .Cells(ra, 20) = "N"
  .Cells(ra, 22) = "25"
  .Cells(ra, 23) = "53"
  .Cells(ra, 24) = "13"
  .Cells(ra, 25) = "6"
  .Cells(ra, 61) = PRSETPHF469MPN
End With
End Sub

Sub FlagSKURules(ns As Worksheet, ra As Long)  ' ra is first row of each set of product data
' write empty strings to all cells in data area
Dim i As Long, j As Long
With ns
  For j = 1 To 69
    For i = ra + 7 To ra + 12
        .Cells(i, j) = ""
    Next
  Next
  DoEvents ' give Windows a moment for housekeeping

'12inX18inHG Rule - product row 8
  ra = ra + 7
  .Cells(ra, 1) = "RULE"
  .Cells(ra, 6) = PRSETPH357SKU
  .Cells(ra, 15) = PRSETPH357Price
  .Cells(ra, 26) = "Y"
  .Cells(ra, 27) = "Y"
  .Cells(ra, 33) = "https://emflag.com/content/Flags/State/Indoor/Presentation%20Sets/PRSET-" + state + "-PH.png"
  .Cells(ra, 34) = "https://emflag.com/content/Flags/State/Indoor/Presentation%20Sets/PRSET-" + state + "-PH.png"
  .Cells(ra, 38) = "N"

'2X3 HG Rule - product row 9
  ra = ra + 1
  .Cells(ra, 1) = "RULE"
  .Cells(ra, 6) = PRSETPHF357SKU
  .Cells(ra, 15) = PRSETPHF357Price
  .Cells(ra, 26) = "Y"
  .Cells(ra, 27) = "Y"
  .Cells(ra, 33) = "https://emflag.com/content/Flags/State/Indoor/Presentation%20Sets/PRSET-" + state + "-PHF.png"
  .Cells(ra, 34) = "https://emflag.com/content/Flags/State/Indoor/Presentation%20Sets/PRSET-" + state + "-PHF.png"
  .Cells(ra, 38) = "N"

'3X5 HG Rule - product row 10
  ra = ra + 1
  .Cells(ra, 1) = "RULE"
  .Cells(ra, 6) = PRSETPH358SKU
  .Cells(ra, 15) = PRSETPH358Price
  .Cells(ra, 26) = "Y"
  .Cells(ra, 27) = "Y"
  .Cells(ra, 33) = "https://emflag.com/content/Flags/State/Indoor/Presentation%20Sets/PRSET-" + state + "-PH.png"
  .Cells(ra, 34) = "https://emflag.com/content/Flags/State/Indoor/Presentation%20Sets/PRSET-" + state + "-PH.png"
  .Cells(ra, 38) = "N"

'4X6 HG Rule  - product row 11
  ra = ra + 1
  .Cells(ra, 1) = "RULE"
  .Cells(ra, 6) = PRSETPHF358SKU
  .Cells(ra, 15) = PRSETPHF358Price
  .Cells(ra, 26) = "Y"
  .Cells(ra, 27) = "Y"
  .Cells(ra, 33) = "https://emflag.com/content/Flags/State/Indoor/Presentation%20Sets/PRSET-" + state + "-PHF.png"
  .Cells(ra, 34) = "https://emflag.com/content/Flags/State/Indoor/Presentation%20Sets/PRSET-" + state + "-PHF.png"
  .Cells(ra, 38) = "N"

'5X8 HG Rule  - product row 12
ra = ra + 1
  .Cells(ra, 1) = "RULE"
  .Cells(ra, 6) = PRSETPH469SKU
  .Cells(ra, 15) = PRSETPH469Price
  .Cells(ra, 26) = "Y"
  .Cells(ra, 27) = "Y"
  .Cells(ra, 33) = "https://emflag.com/content/Flags/State/Indoor/Presentation%20Sets/PRSET-" + state + "-PH.png"
  .Cells(ra, 34) = "https://emflag.com/content/Flags/State/Indoor/Presentation%20Sets/PRSET-" + state + "-PH.png"
  .Cells(ra, 38) = "N"

'6X10 HG Rule  - product row 13
  ra = ra + 1
  .Cells(ra, 1) = "RULE"
  .Cells(ra, 6) = PRSETPHF469SKU
  .Cells(ra, 15) = PRSETPHF469Price
  .Cells(ra, 26) = "Y"
  .Cells(ra, 27) = "Y"
  .Cells(ra, 33) = "https://emflag.com/content/Flags/State/Indoor/Presentation%20Sets/PRSET-" + state + "-PHF.png"
  .Cells(ra, 34) = "https://emflag.com/content/Flags/State/Indoor/Presentation%20Sets/PRSET-" + state + "-PHF.png"
  .Cells(ra, 38) = "N"
End With
End Sub

