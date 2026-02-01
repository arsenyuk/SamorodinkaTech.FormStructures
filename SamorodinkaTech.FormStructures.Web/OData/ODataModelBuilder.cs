using Microsoft.OData.Edm;
using Microsoft.OData.ModelBuilder;

namespace SamorodinkaTech.FormStructures.Web.OData;

public static class ODataModelBuilder
{
    public static IEdmModel BuildEdmModel()
    {
        var builder = new ODataConventionModelBuilder();

        builder.Namespace = "SamorodinkaTech.FormStructures.Web.Api";

        var uploads = builder.EntitySet<ODataUpload>("Data");
        uploads.EntityType.HasKey(x => x.Id);
        uploads.EntityType.Name = "DataSet";

        var rows = builder.EntitySet<ODataRow>("Rows");
        rows.EntityType.HasKey(x => x.Id);
        // Convention will treat IDictionary<string, object?> as the open-type dynamic property container.
        rows.EntityType.Name = "DataRow";

        var cols = builder.EntitySet<ODataColumn>("Columns");
        cols.EntityType.HasKey(x => x.Id);
        cols.EntityType.Name = "DataColumn";

        // Navigation properties are discovered from CLR properties.

        return builder.GetEdmModel();
    }
}
