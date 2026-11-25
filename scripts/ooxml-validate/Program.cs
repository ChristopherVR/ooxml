using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var version = FileFormatVersions.Microsoft365;
var inputs = new List<string>();

for (var i = 0; i < args.Length; i++)
{
	if (args[i] == "--version" && i + 1 < args.Length)
	{
		if (!Enum.TryParse(args[++i], ignoreCase: true, out version))
		{
			Console.Error.WriteLine($"Unknown Office version: {args[i]}");
			return 2;
		}
		continue;
	}
	inputs.Add(args[i]);
}

if (inputs.Count == 0)
{
	Console.Error.WriteLine("Usage: ooxml-validate [--version Microsoft365] <file-or-dir> [...]");
	return 2;
}

var extensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".docx", ".xlsx", ".pptx" };
var files = new List<string>();
foreach (var input in inputs)
{
	if (Directory.Exists(input))
	{
		files.AddRange(
			Directory
				.EnumerateFiles(input, "*.*", SearchOption.AllDirectories)
				.Where(path => extensions.Contains(Path.GetExtension(path)))
				.OrderBy(path => path, StringComparer.Ordinal)
		);
	}
	else if (File.Exists(input))
	{
		files.Add(input);
	}
	else
	{
		Console.Error.WriteLine($"Path not found: {input}");
		return 2;
	}
}

if (files.Count == 0)
{
	Console.Error.WriteLine("No .docx, .xlsx, or .pptx files found to validate.");
	return 2;
}

var validator = new OpenXmlValidator(version);
var failed = 0;

foreach (var file in files)
{
	try
	{
		using OpenXmlPackage package = Path.GetExtension(file).ToLowerInvariant() switch
		{
			".docx" => WordprocessingDocument.Open(file, false),
			".xlsx" => SpreadsheetDocument.Open(file, false),
			".pptx" => PresentationDocument.Open(file, false),
			var ext => throw new InvalidOperationException($"Unsupported extension: {ext}"),
		};

		var errors = validator.Validate(package).ToList();
		if (errors.Count == 0)
		{
			Console.WriteLine($"OK   {Path.GetFileName(file)}");
			continue;
		}

		failed++;
		Console.Error.WriteLine($"FAIL {Path.GetFileName(file)} ({errors.Count} error(s)):");
		foreach (var error in errors)
		{
			var part = error.Part?.Uri?.ToString() ?? "<package>";
			var path = error.Path?.XPath ?? "";
			Console.Error.WriteLine($"  [{error.ErrorType}] {part} {path}: {error.Description}");
		}
	}
	catch (Exception ex)
	{
		failed++;
		Console.Error.WriteLine($"FAIL {Path.GetFileName(file)}: {ex.GetType().Name}: {ex.Message}");
	}
}

Console.WriteLine($"Validated {files.Count} OOXML package(s) against {version}: {files.Count - failed} passed, {failed} failed.");
return failed == 0 ? 0 : 1;
