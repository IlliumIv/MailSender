using CLIArgsHandler;
using CsvHelper;
using CsvHelper.Configuration;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimeKit.Text;
using System.Net.Mime;
using System.Text;

namespace MailSender;

internal class Program
{
    static void Main(string[] args)
    {
        try
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var nonParams = ArgumentsHandler<Parameters>.Parse(args, $"Usage: {nameof(MailSender)} <directory>");

            var pathes = new HashSet<FileInfo>();

            foreach (var path in nonParams)
            {
                var attrs = File.GetAttributes(path);

                if (attrs.HasFlag(FileAttributes.Directory))
                    pathes = [.. pathes, .. Directory.GetFiles(path).Select(x => new FileInfo(x))];
                else
                    pathes.Add(new(path));
            }

            var csvConfig = new CsvConfiguration(Parameters.File_Culture.Value)
            {
                Delimiter = Parameters.Column_Delimeter.Value,
                MissingFieldFound = null,
            };

            var csvFileInfo = pathes.First(x => x.Extension == ".csv");

            using var reader = new StreamReader(csvFileInfo.FullName, Parameters.File_Encoding.Value);
            using var csv = new CsvReader(reader, csvConfig);
            csv.Context.RegisterClassMap<PersonMap>();

            var emailFileInfo = pathes.First(x => x.Extension == ".txt");

            try
            {
                var emailText = File.ReadAllText(emailFileInfo.FullName);
                var persons = csv.GetRecords<Person>();

                var counter = 0;

                foreach (var person in persons)
                {
                    try
                    {
                        Console.Write($"Sending email to {person.FullName} ({person.Email})...");

                        var pdfFileInfo = pathes.First(x => x.Extension == ".pdf" && x.Name.Contains(person.FullName));
                        SendEmail(person, pdfFileInfo, emailText);

                        Console.WriteLine(" Success.");

                        if (Parameters.MailingDelay.Value > 0 && counter >= Parameters.MailingBunch.Value)
                        {
                            Console.Write($"Wait {Parameters.MailingDelay.Value} ms...");
                            Task.Delay(Parameters.MailingDelay.Value);
                            Console.WriteLine(" Done.");
                        }

                        counter++;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($" Error: {e.Message}{Environment.NewLine}{Environment.NewLine}" +
                            $"Type C to continue or A to abort mailing.");
                        var key = Console.ReadKey();

                        if (key.Key == ConsoleKey.C)
                        {
                            Console.WriteLine(Environment.NewLine);
                            continue;
                        }

                        if (key.Key == ConsoleKey.A)
                        {
                            return;
                        }
                    }
                }
            }
            catch (FieldValidationException e)
            {
                var error = e.Message.Split(Environment.NewLine)[0];
                var rawRecord = e?.Context?.Reader?.Parser.RawRecord;
                var row = e?.Context?.Reader?.Parser.Row;
                var delimeter = e?.Context?.Reader?.Parser.Delimiter;

                Console.WriteLine($"Error on parse file. {error}{Environment.NewLine}" +
                    $"  with delimeter \"{delimeter}\"{Environment.NewLine}" +
                    $"  at \"{rawRecord?.TrimEnd('\r', '\n')}\"{Environment.NewLine}" +
                    $"  at {csvFileInfo}:line {row}");
            }
            catch (HeaderValidationException e)
            {
                var error = e.Message.Split(Environment.NewLine)[0];
                var rawRecord = e?.Context?.Reader?.Parser.RawRecord;
                var row = e?.Context?.Reader?.Parser.Row;
                var delimeter = e?.Context?.Reader?.Parser.Delimiter;

                Console.WriteLine($"Error on parse file. {error}{Environment.NewLine}" +
                    $"  with delimeter \"{delimeter}\"{Environment.NewLine}" +
                    $"  at \"{rawRecord?.TrimEnd('\r', '\n')}\"{Environment.NewLine}" +
                    $"  at {csvFileInfo}:line {row}");
            }
        }
        catch (Exception e)
        {
            Console.WriteLine($"{e.Message}{Environment.NewLine}{e.StackTrace}{Environment.NewLine}");
            ArgumentsHandler<Parameters>.ShowHelp();
        }
    }

    public static void SendEmail(Person person, FileInfo pdf, string mailText)
    {
        var multipart = new Multipart("mixed");

        var textPart = new TextPart(TextFormat.Text)
        {
            Text = mailText,
            ContentTransferEncoding = ContentEncoding.Base64
        };

        multipart.Add(textPart);
        
        var pdfBytes = File.ReadAllBytes(pdf.FullName);
        var memoryStream = new MemoryStream(pdfBytes);
        var attachmentPart = new MimePart(MediaTypeNames.Application.Pdf)
        {
            Content = new MimeContent(memoryStream),
            ContentId = pdf.Name,
            ContentTransferEncoding = ContentEncoding.Base64,
            FileName = pdf.Name
        };
        multipart.Add(attachmentPart);

        var message = new MimeMessage();
        message.From.Add(new MailboxAddress(Parameters.SenderName.Value, Parameters.Login.Value));
        message.To.Add(new MailboxAddress(person.FullName, person.Email));
        message.Subject = Parameters.EmailSubject.Value;
        message.Body = multipart;

        using var client = new SmtpClient();
        client.Connect(Parameters.Address.Value, Parameters.Port.Value, SecureSocketOptions.StartTls);

        client.Authenticate(Parameters.Login.Value, Parameters.Password.Value);

        client.Send(message);
        client.Disconnect(true);
    }
}
