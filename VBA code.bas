Attribute VB_Name = "Module3"
Sub FindPhoneLink()

    Dim phoneTypeInput As String
    Dim priceInput As Double
    Dim ratingsInput As Double
    Dim phoneType As String
    Dim price As Double
    Dim ratings As Double
    Dim link As String
    Dim row As Integer
    Dim foundMatch As Boolean

    foundMatch = False ' Initialize to check if any match is found

    ' Ask the user for input
    phoneTypeInput = InputBox("Enter Phone Type (e.g., Samsung, Apple)")
    priceInput = CDbl(InputBox("Enter your budget price"))
    ratingsInput = CDbl(InputBox("Enter the minimum rating (e.g., 4.0)"))

    ' Loop through the data in Excel (assuming data starts from row 2)
    For row = 2 To 10 ' Adjust the row range as per your data
        phoneType = Cells(row, 1).value ' PhoneType in Column A
        link = Cells(row, 2).value ' Link in Column B
        ratings = Cells(row, 3).value ' Ratings in Column C
        price = Cells(row, 4).value ' Price in Column D

        ' Check if Phone Type, Price, and Ratings match the user's input
        If InStr(1, phoneType, phoneTypeInput, vbTextCompare) > 0 And price <= priceInput And ratings >= ratingsInput Then
            ' If it matches, display the link
            MsgBox "Check out the phone details here: " & link
            foundMatch = True ' Set to True if a match is found
        End If
    Next row

    ' If no match is found, inform the user
    If Not foundMatch Then
        MsgBox "No phone matches your criteria."
    End If

End Sub

