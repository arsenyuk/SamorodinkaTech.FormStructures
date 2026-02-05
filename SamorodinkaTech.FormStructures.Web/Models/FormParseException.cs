namespace SamorodinkaTech.FormStructures.Web.Models;

public sealed class FormParseException : Exception
{
    public FormParseException(string message) : base(message)
    {
    }

    public FormParseException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
