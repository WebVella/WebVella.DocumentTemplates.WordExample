using System.Data;
using WebVella.DocumentTemplates.Engines.DocumentFile;

namespace WebVella.DocumentTemplates.WordExample;

internal class Program
{
	static void Main(string[] args)
	{
		var templateFile = "TemplateDoc1.docx";
		var ds = new DataTable();
		ds.Columns.Add("position", typeof(int));
		ds.Columns.Add("sku", typeof(string));
		ds.Columns.Add("item", typeof(string));
		ds.Columns.Add("price", typeof(decimal));

		for (int i = 1; i < 6; i++)
		{
			var dsrow = ds.NewRow();
			dsrow["position"] = i;
			dsrow["sku"] = $"SKU{i}";
			dsrow["item"] = $"item {i} description text";
			dsrow["price"] = i * (decimal)0.98;
			ds.Rows.Add(dsrow);
		}

		var template = new WvDocumentFileTemplate
		{
			Template = new Utils().LoadFileAsStream(templateFile)
		};
		WvDocumentFileTemplateProcessResult? result = template.Process(ds);

		new Utils().SaveFileFromStream(result.ResultItems[0].Result!, $"result-{templateFile}");
	}
}
