using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebVella.DocumentTemplates.WordExample;
public class Utils
{
	public MemoryStream LoadFileAsStream(string fileName)
	{
		var path = Path.Combine(Environment.CurrentDirectory, $"{fileName}");
		var fi = new FileInfo(path);
		if (!fi.Exists) throw new FileNotFoundException();
		MemoryStream ms = new MemoryStream();
		using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
			file.CopyTo(ms);
		return ms;
	}
	public void SaveFileFromStream(MemoryStream content, string fileName)
	{
		DirectoryInfo? debugFolder = Directory.GetParent(Environment.CurrentDirectory);
		if (debugFolder is null) throw new Exception("debugFolder not found");
		var projectFolder = debugFolder.Parent?.Parent;
		if (projectFolder is null) throw new Exception("projectFolder not found");

		var path = Path.Combine(projectFolder.FullName,fileName);
		using (FileStream fs = new FileStream(path, FileMode.Create))
		{
			content.WriteTo(fs);
		}
	}
}
