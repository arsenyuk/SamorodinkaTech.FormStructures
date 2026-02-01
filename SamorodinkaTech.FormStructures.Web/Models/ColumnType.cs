namespace SamorodinkaTech.FormStructures.Web.Models;

using System.ComponentModel.DataAnnotations;

public enum ColumnType
{
    String = 0,
    Date = 1,
    DateTime = 2,
    Int = 3,
    Decimal = 4,

    [Display(Name = "Reference book")]
    ReferenceBook = 5
}
