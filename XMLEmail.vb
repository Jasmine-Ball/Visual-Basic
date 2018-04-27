Imports System.Net.Mail
Imports System.Xml.XPath
Imports System.Text.RegularExpressions
Public Class XMLEmail

    Private Sub Form1_Load()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

            ' Parsing the XML and selecting the data based on the parameter
 
            ' Reading in template for Regex

            ' Regex customer details placement in document starts here

            ' Sender email credential management

            ' Creating SMTP based email objects

            ' Sender's mail server details, in this case it's Gmail

            ' Image embedding starts here

            ' Setting the email properties

    End Sub


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        MessageBox.Show("TextBox1_TextChanged")
    End Sub

End Class