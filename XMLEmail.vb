Imports System.Net.Mail
Imports System.Xml.XPath
Imports System.Text.RegularExpressions



Public Class XMLEmail

    Private Sub SendButtonClick(panel As Object, e As EventArgs) Handles BtnSend.Click

        Try
            ' Declared here so that the variables would be available to multiple functions	
            Dim custAccountNumberXMLData As String = ""
            Dim custFirstNameXMLData As String = ""
            Dim custLastNameXMLData As String = ""
            Dim custAddress1XMLData As String = ""
            Dim custAddress2XMLData As String = ""
            Dim custCityXMLData As String = ""
            Dim custStateXMLData As String = ""
            Dim custPostCodeXMLData As String = ""
            Dim custCountryXMLData As String = ""
            Dim custEmailXMLData As String = ""
            Dim HTMLDataPopulated As String = ""

            ' Reading in template for Regex
            Dim HTMLTemplateRead As String = My.Computer.FileSystem.ReadAllText("C:\CCM\Templates\binTemplate.html") 'Reading in the HTML template as text

            ' Sender email credential management
            Dim eml As String = "testjunke@gmail.com" 'This will hold the sender's email address
            Dim pwd As String = TbxPass.Text 'Senders password will need to be entered. Later sent over SSL


            ' Looping through the customers in the XML file
            Dim count As Integer = 1 ' This is used as the starting array index for the XMLParse function
            Do While count <= 6

                ' Parsing the XML and selecting the data based on the parameter
                custAccountNumberXMLData = XMLParse("CustAccountNumber", count)  'I can call this for any of the XML fields
                custFirstNameXMLData = XMLParse("CustFirstName", count)
                custLastNameXMLData = XMLParse("CustLastName", count)
                custAddress1XMLData = XMLParse("CustAddress1", count)
                custAddress2XMLData = XMLParse("CustAddress2", count)
                custCityXMLData = XMLParse("CustCity", count)
                custStateXMLData = XMLParse("CustState", count)
                custPostCodeXMLData = XMLParse("CustPostCode", count)
                custCountryXMLData = XMLParse("CustCountry", count)
                custEmailXMLData = XMLParse("CustEmail", count)
                count += 1

                ' Regex cutomer details placement via CustomerWrite function
                Dim HTMLWithCustAccountNumber As String = CustomerWrite("CustAccountNumber", custAccountNumberXMLData, HTMLTemplateRead)
                Dim HTMLWithCustFirstName As String = CustomerWrite("CustFirstName", custFirstNameXMLData, HTMLWithCustAccountNumber)
                Dim HTMLWithCustLastName As String = CustomerWrite("CustLastName", custLastNameXMLData, HTMLWithCustFirstName)
                Dim HTMLWithCustAddress1 As String = CustomerWrite("CustAddress1", custAddress1XMLData, HTMLWithCustLastName)
                Dim HTMLWithCustAddress2 As String = CustomerWrite("CustAddress2", custAddress2XMLData, HTMLWithCustAddress1)
                Dim HTMLWithCustCity As String = CustomerWrite("CustCity", custCityXMLData, HTMLWithCustAddress2)
                Dim HTMLWithCustState As String = CustomerWrite("CustState", custStateXMLData, HTMLWithCustCity)
                Dim HTMLWithCustPostcode As String = CustomerWrite("CustPostCode", custPostCodeXMLData, HTMLWithCustState)
                Dim HTMLWithCustCountry As String = CustomerWrite("CustCountry", custCountryXMLData, HTMLWithCustPostcode)
                HTMLDataPopulated = HTMLWithCustCountry

                ' Creating SMTP based email objects
                Dim Smtp As New SmtpClient
                Dim mail As New MailMessage()

                ' Sender's mail server details, in this case it's Gmail
                Smtp.UseDefaultCredentials = False
                Smtp.Credentials = New Net.NetworkCredential(eml, pwd)
                Smtp.Port = 587
                Smtp.EnableSsl = True 'Using SSL for security purposes
                Smtp.Host = "smtp.gmail.com"

                ' Image embedding starts here
                Dim logoBin = New System.Net.Mail.Attachment("C:\CCM\Templates\logo_bin.PNG") 'Reading in the logo as a jpeg
                logoBin.ContentId = "logo_bin.PNG"
                Dim sigCEO = New System.Net.Mail.Attachment("C:\CCM\Templates\sig_mfitzgerald_mc_300.jpeg") 'Reading in the CEO's Signature as a jpeg
                sigCEO.ContentId = "sig_mfitzgerald_mc_300.jpeg" 'Giving the image an identifier for use in the HTML file

                ' Setting the email properties
                mail.Body = HTMLDataPopulated 'Custom email content based on HTML Template
                mail.From = New MailAddress("testjunke@gmail.com")
                mail.To.Add(custEmailXMLData) 'Destination email address based on XML
                mail.Subject = "Welcome, " & custFirstNameXMLData
                mail.Attachments.Add(sigCEO)
                mail.Attachments.Add(logoBin)
                mail.IsBodyHtml = True
                Smtp.Send(mail)
                Application.Exit() 'Closing off the application once the email has been sent
            Loop

        Catch err As Exception 'Basic error handling
            MsgBox(err.ToString)
            Application.Exit()
        End Try

    End Sub

    ' Parsing XML tags for each customer
    Function XMLParse(element, count) As String
        Dim regSelect = "./Customer[" & count & "]"
        Dim document As XElement = XElement.Load("C:\CCM\Input\customerList.xml")
        Dim oneElement As XElement = document.XPathSelectElement(regSelect.ToString).XPathSelectElement(element)   'Selecting the details of the customer
        Dim oneElementAsString = oneElement.ToString
        Dim oneElementAsStringRegex As Regex = New Regex("<.*?>") 'Regex filtering is being used to tidy XML tags
        Dim openTag As Match = oneElementAsStringRegex.Match(oneElementAsString) 'This is where we're looking for the above characters. 
        Dim openTagTrimmed As String = oneElementAsString.Replace(openTag.ToString, "") 'Replacing unwanted open tag with ""
        Dim closeTag As Match = oneElementAsStringRegex.Match(openTagTrimmed) 'This is looking for the second xml tag in the trimmed string
        Dim finalTrimmedXML As String = openTagTrimmed.Replace(closeTag.ToString, "") 'Replacing the unwanted closing xml tag with ""
        Return finalTrimmedXML
    End Function

    ' Placing customer data onto the in-memory HTML template
    Function CustomerWrite(dataLoc, XMLField, HTMLTemplateRead) As String
        Dim textLocationRegex As Regex = New Regex(dataLoc)
        Dim custData As Match = textLocationRegex.Match(HTMLTemplateRead)
        Dim HTMLWithCustomData As String = HTMLTemplateRead.Replace(custData.ToString, XMLField) 'Replacing generic placeholder text with actual customer name.
        Return HTMLWithCustomData
    End Function

    Private Sub XMLEmail_Load(panel As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TbxPass_TextChanged(panel As Object, e As EventArgs) Handles TbxPass.TextChanged

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles LblPass.Click

    End Sub
End Class
