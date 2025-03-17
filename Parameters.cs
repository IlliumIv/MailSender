using CLIArgsHandler;
using System.Globalization;
using System.Text;

namespace MailSender;

public class Parameters
{
    public static Parameter ShowHelpMessage { get; } =
        new(prefixes: ["--help", "-h", "-?"],
            format: string.Empty,
            descriptionFormatter: () => "Show this message and exit.",
            sortingOrder: 11,
            parser: (args, i) =>
            {
                ArgumentsHandler<Parameters>.ShowHelp();
                Environment.Exit(0);
                return args.RemoveAt(i, 1);
            });

    public static Parameter<string> Address { get; } =
        new(prefixes: ["--server", "-s"],
            value: "mx1.macroscoptrade.com",
            format: "url",
            descriptionFormatter: () => $"Mail server address. Current value is \"{Address?.Value}\".",
            sortingOrder: 2,
            parser: (args, i) =>
            {
                if (Address is not null)
                    Address.Value = args[i + 1];
                return args.RemoveAt(i, 2);
            });

    public static Parameter<ushort> Port { get; } =
        new(prefixes: ["--port", "-p"],
            value: 587,
            format: "number",
            descriptionFormatter: () => $"Mail server port. Current value is \"{Port?.Value}\".",
            sortingOrder: 2,
            parser: (args, i) =>
            {
                if (Port is not null)
                    Port.Value = ushort.Parse(args[i + 1]);
                return args.RemoveAt(i, 2);
            });

    public static Parameter<int> MailingDelay { get; } =
        new(prefixes: ["--delay"],
            value: 0,
            format: "number",
            descriptionFormatter: () => $"Delay between sending emails in ms. Set up this if mail server return error by spam cause. Current value is \"{MailingDelay?.Value}\".",
            sortingOrder: 2,
            parser: (args, i) =>
            {
                if (MailingDelay is not null)
                    MailingDelay.Value = int.Parse(args[i + 1]);
                return args.RemoveAt(i, 2);
            });

    public static Parameter<int> MailingBunch { get; } =
        new(prefixes: ["--bunch"],
            value: 1,
            format: "number",
            descriptionFormatter: () => $"Number of emails before the delay timeout will triggered. Current value is \"{MailingBunch?.Value}\".",
            sortingOrder: 2,
            parser: (args, i) =>
            {
                if (MailingBunch is not null)
                    MailingBunch.Value = int.Parse(args[i + 1]);
                return args.RemoveAt(i, 2);
            });

    public static Parameter<string> Login { get; } =
        new(prefixes: ["--login", "-l"],
            value: string.Empty,
            format: "string",
            isRequired: true,
            descriptionFormatter: () => $"Login. Current value is \"{Login?.Value}\".",
            sortingOrder: 0,
            parser: (args, i) =>
            {
                if (Login is not null)
                    Login.Value = args[i + 1];
                return args.RemoveAt(i, 2);
            });

    public static Parameter<string> Password { get; } =
        new(prefixes: ["--password"],
            value: string.Empty,
            format: "string",
            isRequired: true,
            descriptionFormatter: () => $"Password. Current value is \"{Password?.Value}\".",
            sortingOrder: 0,
            parser: (args, i) =>
            {
                if (Password is not null)
                    Password.Value = args[i + 1];
                return args.RemoveAt(i, 2);
            });

    public static Parameter<string> SenderName { get; } =
        new(prefixes: ["--sender-name",],
            value: string.Empty,
            format: "string",
            descriptionFormatter: () => $"Sender name. Current value is \"{SenderName?.Value}\".",
            sortingOrder: 0,
            parser: (args, i) =>
            {
                if (SenderName is not null)
                    SenderName.Value = args[i + 1];
                return args.RemoveAt(i, 2);
            });

    public static Parameter<string> EmailSubject { get; } =
        new(prefixes: ["--subject",],
            value: "Сертификат о прохождении базового курса",
            format: "string",
            descriptionFormatter: () => $"Email subject. Current value is \"{EmailSubject?.Value}\".",
            sortingOrder: 1,
            parser: (args, i) =>
            {
                if (EmailSubject is not null)
                    EmailSubject.Value = args[i + 1];
                return args.RemoveAt(i, 2);
            });

    public static Parameter ShowEncodings { get; } =
        new(prefixes: ["--show-encodings"],
            format: string.Empty,
            descriptionFormatter: () => "Show all possible encodings and exit.",
            sortingOrder: 9,
            parser: (args, i) =>
            {
                Console.WriteLine(string.Join(", ", Encoding
                    .GetEncodings()
                    .Select(e => new[] { e.Name })
                    .SelectMany(e => e)));
                Environment.Exit(0);
                return args.RemoveAt(i, 1);
            });

    public static Parameter ShowCultures { get; } =
        new(prefixes: ["--show-cultures"],
            format: string.Empty,
            descriptionFormatter: () => "Show all possible cultures and exit.",
            sortingOrder: 9,
            parser: (args, i) =>
            {
                Console.WriteLine(string.Join(", ", CultureInfo
                    .GetCultures(CultureTypes.AllCultures)
                    .Select(c => new[] { c.Name })
                    .Skip(1).SelectMany(c => c)));
                Environment.Exit(0);
                return args.RemoveAt(i, 1);
            });

    public static Parameter<Encoding> File_Encoding { get; } =
        new(prefixes: ["--encoding"],
            value: Encoding.UTF8,
            format: "string",
            descriptionFormatter: () => $"File encoding. Current value is \"{File_Encoding?.Value.HeaderName}\". " +
                $"To see all possible encodings specify {string.Join(", ", ShowEncodings.Prefixes)}.",
            sortingOrder: 7,
            parser: (args, i) =>
            {
                if (File_Encoding is not null)
                    File_Encoding.Value = Encoding.GetEncoding(args[i + 1]);
                return args.RemoveAt(i, 2);
            });

    public static Parameter<CultureInfo> File_Culture { get; } =
        new(prefixes: ["--culture"],
            value: CultureInfo.InvariantCulture,
            format: "string",
            descriptionFormatter: () => $"File culture. Current value is \"{File_Culture?.Value.DisplayName} ({File_Culture?.Value.Name})\". " +
                $"To see all possible cultures specify {string.Join(", ", ShowCultures.Prefixes)}.",
            sortingOrder: 7,
            parser: (args, i) =>
            {
                if (File_Culture is not null)
                    File_Culture.Value = new CultureInfo(args[i + 1]);
                return args.RemoveAt(i, 2);
            });

    public static Parameter<string> Column_Delimeter { get; } =
        new(prefixes: ["--delimeter"],
            value: ",",
            format: "string",
            descriptionFormatter: () => $"Columns delimeter." +
                $" Current value is \"{File_Culture.Value.TextInfo.ListSeparator}\"." +
                $" By default it depends on culture, current selected culture is \"{File_Culture.Value.DisplayName} ({File_Culture?.Value.Name})\".",
            sortingOrder: 6,
            parser: (args, i) =>
            {
                if (Column_Delimeter is not null)
                    Column_Delimeter.Value = args[i + 1];
                return args.RemoveAt(i, 2);
            });

    public static Parameter<string> Column_FullName { get; } =
        new(prefixes: ["--names"],
            value: "Full Name",
            format: "string",
            descriptionFormatter: () => $"Header of column that contains names. Current value is \"{Column_FullName?.Value}\".",
            sortingOrder: 6,
            parser: (args, i) =>
            {
                if (Column_FullName is not null)
                    Column_FullName.Value = args[i + 1];
                return args.RemoveAt(i, 2);
            });

    public static Parameter<string> Column_Email { get; } =
        new(prefixes: ["--emails"],
            value: "Email",
            format: "string",
            descriptionFormatter: () => $"Header of column that contains emails. Current value is \"{Column_Email?.Value}\".",
            sortingOrder: 6,
            parser: (args, i) =>
            {
                if (Column_Email is not null)
                    Column_Email.Value = args[i + 1];
                return args.RemoveAt(i, 2);
            });
}
