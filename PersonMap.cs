using CsvHelper.Configuration;
using System.Diagnostics.CodeAnalysis;

namespace MailSender;

public class Person
{
    public string? FullName { get; set; }
    public string? Email { get; set; }
}

public class PersonMap : ClassMap<Person>
{
    public PersonMap()
    {
        Map(m => m.FullName).Name(Parameters.Column_FullName.Value);
        Map(m => m.Email).Name(Parameters.Column_Email.Value);
    }
}