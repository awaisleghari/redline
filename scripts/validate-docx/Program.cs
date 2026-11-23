// Validates Word files against the Office Open XML schema using Microsoft's
// Open XML SDK validator (check 7 in PLAN.md: "the file validates against the
// Office Open XML schema").
//
// Usage: dotnet run -- <folder>
// Validates every .docx in the folder. Prints one line per file, then each
// problem found. Exit code 0 when every file is valid, 1 otherwise.
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var folder = args.Length > 0 ? args[0] : ".";
var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
var invalid = 0;
foreach (var path in Directory.GetFiles(folder, "*.docx").OrderBy(p => p, StringComparer.Ordinal))
{
    var name = Path.GetFileName(path);
    try
    {
        using var document = WordprocessingDocument.Open(path, false);
        var errors = validator.Validate(document).ToList();
        Console.WriteLine($"{(errors.Count == 0 ? "VALID  " : "INVALID")} {name}");
        foreach (var error in errors.Take(20))
        {
            Console.WriteLine($"    [{error.Part?.Uri}] {error.Description}");
            Console.WriteLine($"      at {error.Path?.XPath}");
        }
        if (errors.Count > 20) Console.WriteLine($"    ... and {errors.Count - 20} more");
        if (errors.Count > 0) invalid++;
    }
    catch (Exception e)
    {
        Console.WriteLine($"INVALID {name}");
        Console.WriteLine($"    could not be opened: {e.Message}");
        invalid++;
    }
}
return invalid == 0 ? 0 : 1;
