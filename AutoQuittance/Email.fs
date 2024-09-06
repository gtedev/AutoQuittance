module Email

open System
open System.IO
open System.Threading
open Google.Apis.Auth.OAuth2
open Google.Apis.Gmail.v1
open Google.Apis.Gmail.v1.Data
open Google.Apis.Services
open Google.Apis.Util.Store
open MimeKit

let Scopes = [| GmailService.Scope.GmailSend |]
let ApplicationName = "Gmail API F# Quickstart"

let sendEmailWithAttachment () =
    let credential =
        use stream = new FileStream("", FileMode.Open, FileAccess.Read)
        let credPath = "token.json"

        GoogleWebAuthorizationBroker
            .AuthorizeAsync(
                GoogleClientSecrets.FromStream(stream).Secrets,
                Scopes,
                "user",
                CancellationToken.None,
                new FileDataStore(credPath, true)
            )
            .Result

    // Create Gmail API service.
    let service =
        new GmailService(
            new BaseClientService.Initializer(HttpClientInitializer = credential, ApplicationName = ApplicationName)
        )

    // Create the email to send
    let msg = new MimeMessage()
    msg.From.Add(MailboxAddress("User1", "user1@gmail.com"))
    msg.To.Add(MailboxAddress("User2", "user2.te@gmail.com"))
    msg.Subject <- "Test Email with PDF Attachment from Gmail API"

    // Create the body of the email
    let body = new TextPart("plain")
    body.Text <- "This is a test email sent using Gmail API with F#."

    // Attach the PDF file
    let attachment = new MimePart("application", "pdf")
    attachment.Content <- new MimeContent(File.OpenRead("C:\\toto.pdf"))
    attachment.ContentDisposition <- new ContentDisposition(ContentDisposition.Attachment)
    attachment.ContentTransferEncoding <- ContentEncoding.Base64
    attachment.FileName <- "toto.pdf"

    // Create the multipart email
    let multipart = new Multipart("mixed")
    multipart.Add body
    multipart.Add attachment
    msg.Body <- multipart

    // Encode the email as base64url and prepare it for sending
    use memoryStream = new MemoryStream()
    msg.WriteTo(memoryStream)

    let raw =
        Convert
            .ToBase64String(memoryStream.ToArray())
            .Replace('+', '-')
            .Replace('/', '_')
            .Replace("=", "")

    let gmailMessage = new Message(Raw = raw)

    // Send the email
    service.Users.Messages.Send(gmailMessage, "me").Execute() |> ignore
    printfn "Email with PDF attachment sent successfully!"
