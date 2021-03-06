Imports System.Net.Mail
Imports System.Xml.XPath
Imports System.Text.RegularExpressions


Public Class XMLEmail

    Private Sub SendButtonClick(panel As Object, e As EventArgs) Handles BtnSend.Click

        Try
            ' Declared early so that the variables will be available anywhere within the Sub
            Dim custAccountNumberXMLData, custFirstNameXMLData, custLastNameXMLData, custAddress1XMLData, custAddress2XMLData, custCityXMLData,
            custStateXMLData, custPostCodeXMLData, custCountryXMLData, custEmailXMLData, custFundXMLData, HTMLTemplateRead, logo, sigCEO, HTMLDataPopulated

            ' Sender email credential management
            Dim eml As String = "testjunke@gmail.com" 'This will hold the sender's email address
            Dim pwd As String = TbxPass.Text 'Senders password as entered in input box. Later sent to Gmail over SSL

            ' Looping through the customers in the XML file
            Dim counter As Integer = 1 ' This is used as the starting array index for the XMLParse function
            Do While counter <= 6

                ' Parsing the XML and selecting the data based on the parameter
                custAccountNumberXMLData = XMLParse("CustAccountNumber", counter)  'I can call this for any of the XML fields
                custFirstNameXMLData = XMLParse("CustFirstName", counter)
                custLastNameXMLData = XMLParse("CustLastName", counter)
                custAddress1XMLData = XMLParse("CustAddress1", counter)
                custAddress2XMLData = XMLParse("CustAddress2", counter)
                custCityXMLData = XMLParse("CustCity", counter)
                custStateXMLData = XMLParse("CustState", counter)
                custPostCodeXMLData = XMLParse("CustPostCode", counter)
                custCountryXMLData = XMLParse("CustCountry", counter)
                custFundXMLData = XMLParse("CustFund", counter)
                custEmailXMLData = XMLParse("CustEmail", counter)
                counter += 1

                ' Checking which fund each customer is with, based on the XML
                If custFundXMLData Like "Fund1" Then

                    ' Reading in template for Regex
                    HTMLTemplateRead = My.Computer.FileSystem.ReadAllText(".\CCM\Templates\template_binhealth.html") 'Reading in the HTML template as text

                    ' Image embedding starts here
                    logo = New System.Net.Mail.Attachment(".\CCM\Templates\logo_binhealth.PNG") 'Reading in the logo as a PNG
                    logo.ContentId = "logoTop"
                    sigCEO = New System.Net.Mail.Attachment(".\CCM\Templates\sig_mfitzgerald.jpeg") 'Reading in the CEO's Signature as a jpeg
                    sigCEO.ContentId = "sigCEO" 'Giving the image an identifier for use in the HTML file

                ElseIf custFundXMLData Like "Fund2" Then
                    ' Reading in template for Regex
                    HTMLTemplateRead = My.Computer.FileSystem.ReadAllText(".\CCM\Templates\template_badhealth.html") 'Reading in the HTML template as text

                    ' Image embedding starts here
                    logo = New System.Net.Mail.Attachment(".\CCM\Templates\logo_badhealth.PNG") 'Reading in the logo as a PNG
                    logo.ContentId = "logoTop"
                    sigCEO = New System.Net.Mail.Attachment(".\CCM\Templates\sig_devil.jpeg") 'Reading in the CEO's Signature as a jpeg
                    sigCEO.ContentId = "sigCEO" 'Giving the image an identifier for use in the HTML file
                End If

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

                ' Setting the email properties
                mail.Body = HTMLDataPopulated 'Custom email content based on HTML Template
                mail.From = New MailAddress("testjunke@gmail.com")
                mail.To.Add(custEmailXMLData) 'Destination email address based on XML
                mail.Subject = "Welcome, " & custFirstNameXMLData
                mail.Attachments.Add(sigCEO)
                mail.Attachments.Add(logo)
                mail.IsBodyHtml = True
                Smtp.Send(mail)
                Application.Exit() 'Closing off the application once the email has been sent
            Loop

        Catch err As Exception 'Basic error handling
            MsgBox(err.ToString)
            Application.Exit()
        End Try

    End Sub

    ' Parsing XML tags for each customer and tidying with Regex
    Function XMLParse(element, counter) As String
        Dim regSelect = "./Customer[" & counter & "]"
        Dim document As XElement = XElement.Load(".\CCM\Input\list_customers.xml")
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

    ' Default empty Sub provided by Visual Studio
    Private Sub XMLEmail_Load(panel As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    ' Default empty Sub provided by Visual Studio
    Private Sub TbxPass_TextChanged(panel As Object, e As EventArgs) Handles TbxPass.TextChanged
    End Sub

    ' Cancel button exits the application
    Private Sub BtnCancel_Click(sender As Object, e As EventArgs) Handles BtnCancel.Click
        Application.Exit()
    End Sub

End Class
