module Email

open System.Net
open System.Net.Mail

type SendEmailRequest =
    { fromEmail: string
      toEmail: string
      pathFileAttachment: string
      smtpServer: string
      emailCredentials: {| email: string; password: string |} }

let sendEmailWithAttachment request =
    try
        let mail = new MailMessage()
        mail.From <- MailAddress(request.fromEmail)

        mail.To.Add(request.toEmail)
        mail.Subject <- "Quittance de loyer"
        mail.Body <- "This is a test email with an attachment sent from an F# application."

        // If you want to send HTML email
        // mail.IsBodyHtml <- true

        // Specify the file to attach
        let attachment = new Attachment(request.pathFileAttachment)
        mail.Attachments.Add(attachment)

        // Configure the SMTP client
        let smtpClient = new SmtpClient(request.smtpServer, 587)
        smtpClient.Credentials <- NetworkCredential(request.emailCredentials.email, request.emailCredentials.password)
        smtpClient.DeliveryMethod <- SmtpDeliveryMethod.Network
        smtpClient.UseDefaultCredentials  <- false
        smtpClient.EnableSsl <- true

        // Send the email
        smtpClient.Send(mail)
    with ex ->
        printfn "Error sending email: %s" ex.Message
