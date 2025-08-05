VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    Dim mail As Outlook.MailItem
    Dim ns As Outlook.NameSpace
    Dim item As Object
    Dim entryIDs() As String
    Dim htmlBody As String
    Dim orderID As String
    Dim reply As Outlook.MailItem
    Dim productTable As String

    Set ns = Application.GetNamespace("MAPI")
    entryIDs = Split(EntryIDCollection, ",")

    For i = 0 To UBound(entryIDs)
        Set item = ns.GetItemFromID(entryIDs(i))
        If item.Class = olMail Then
            Set mail = item

            ' Only process if sender & subject match

            If InStr(mail.Subject, "[OGAWA]: New order #") = 0 Then Exit Sub

            htmlBody = mail.htmlBody
            orderID = ExtractOrderID(mail.Subject)

            productTable = GetProductTableRows(htmlBody)

            Set reply = mail.ReplyAll
            reply.Subject = "Order Confirmation – #" & orderID
            reply.htmlBody = "<p>Dear Customer,</p>" & _
    "<p>Thank you for your purchase!</p>" & _
    "<p>Your order <strong>#" & orderID & "</strong> is pending courier collection.</p>" & _
    productTable & _
    "<p>You may track the order via <a href=""https://www.jtexpress.sg/trackmyparcel"">https://www.jtexpress.sg/trackmyparcel</a> with tracking code: <strong>[INSERT_TRACKING_CODE]</strong></p>" & _
    "<p>You shall receive the order in 1–3 days after successful pickup.</p>" & _
    "<p>Warm regards,<br>OGAWA Team</p><br><br>" & _
    reply.htmlBody


            reply.Save
        End If
    Next
End Sub

Function ExtractOrderID(subjectLine As String) As String
    Dim startPos As Long, endPos As Long
    startPos = InStr(subjectLine, "#")
    If startPos > 0 Then
        endPos = InStr(startPos, subjectLine, "]")
        If endPos > startPos Then
            ExtractOrderID = Mid(subjectLine, startPos + 1, endPos - startPos - 1)
        Else
            ExtractOrderID = Mid(subjectLine, startPos + 1)
        End If
    Else
        ExtractOrderID = "N/A"
    End If
End Function

Function CleanHTML(ByVal html As String) As String
    html = Replace(html, "<br>", vbCrLf, , , vbTextCompare)
    html = Replace(html, "<br/>", vbCrLf, , , vbTextCompare)
    html = Replace(html, "<br />", vbCrLf, , , vbTextCompare)
    html = Replace(html, "&nbsp;", " ", , , vbTextCompare)
    Do While InStr(html, "<") > 0 And InStr(html, ">") > InStr(html, "<")
        html = Replace(html, Mid(html, InStr(html, "<"), InStr(html, ">") - InStr(html, "<") + 1), "")
    Loop
    CleanHTML = Trim(html)
End Function

Function DecodeHTMLEntities(text As String) As String
    text = Replace(text, "&#8217;", "'")
    text = Replace(text, "&amp;", "&")
    text = Replace(text, "&quot;", "\"")
    text = Replace(text, "&lt;", "<")
    text = Replace(text, "&gt;", ">")
    DecodeHTMLEntities = text
End Function

Function ExtractCell(rowHtml As String, cellIndex As Integer) As String
    Dim tdStart As Long, tdEnd As Long, i As Integer
    Dim searchPos As Long
    searchPos = 1
    For i = 1 To cellIndex
        tdStart = InStr(searchPos, rowHtml, "<td", vbTextCompare)
        If tdStart = 0 Then
            ExtractCell = ""
            Exit Function
        End If
        tdStart = InStr(tdStart, rowHtml, ">", vbTextCompare)
        If tdStart = 0 Then
            ExtractCell = ""
            Exit Function
        End If
        tdStart = tdStart + 1
        tdEnd = InStr(tdStart, rowHtml, "</td>", vbTextCompare)
        If tdEnd = 0 Then
            ExtractCell = ""
            Exit Function
        End If
        searchPos = tdEnd + 1
    Next i
    ExtractCell = CleanHTML(Mid(rowHtml, tdStart, tdEnd - tdStart))
End Function

Function GetProductTableRows(ByVal htmlBody As String) As String
    Dim output As String
    Dim rowStart As Long, rowEnd As Long
    Dim rowBlock As String
    Dim productCell As String, quantityCell As String
    Dim itemText As String, colourText As String
    Dim posColour As Long

    output = ""
    rowEnd = 0
    Do
        rowStart = InStr(rowEnd + 1, htmlBody, "<tr", vbTextCompare)
        If rowStart = 0 Then Exit Do
        rowEnd = InStr(rowStart, htmlBody, "</tr>", vbTextCompare)
        If rowEnd = 0 Then Exit Do

        rowBlock = Mid(htmlBody, rowStart, rowEnd - rowStart + 5)

        If InStr(LCase(rowBlock), "<th") > 0 Or _
           InStr(LCase(rowBlock), "subtotal") > 0 Or _
           InStr(LCase(rowBlock), "discount") > 0 Or _
           InStr(LCase(rowBlock), "free") > 0 Or _
           InStr(LCase(rowBlock), "shipping") > 0 Or _
           InStr(LCase(rowBlock), "payment") > 0 Or _
           InStr(LCase(rowBlock), "total") > 0 Or _
           InStr(LCase(rowBlock), "address") > 0 Then
            GoTo NextRow
        End If

        productCell = ExtractCell(rowBlock, 1)
        quantityCell = ExtractCell(rowBlock, 2)

        posColour = InStr(1, productCell, "Colour:", vbTextCompare)
        If posColour > 0 Then
            itemText = Trim(Left(productCell, posColour - 1))
            colourText = Trim(Mid(productCell, posColour + 7))
        Else
            itemText = productCell
            colourText = "N/A"
        End If

        If Len(itemText) > 0 And Len(quantityCell) > 0 Then
            output = output & "<tr><td>" & itemText & "</td><td>" & colourText & "</td><td>" & quantityCell & "</td></tr>"
        End If
NextRow:
    Loop

    If output <> "" Then
        GetProductTableRows = _
            "<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;font-family:sans-serif;'>" & _
            "<thead><tr><th>Item</th><th>Colour</th><th>Quantity</th></tr></thead><tbody>" & _
            output & "</tbody></table>"
    Else
        GetProductTableRows = "<p><em>No products found.</em></p>"
    End If
End Function



