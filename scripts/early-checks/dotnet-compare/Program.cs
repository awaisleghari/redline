// Early check 1 (see PLAN.md): run the .NET build of the Docxodus engine over
// the same document pairs the browser (WebAssembly) build compared, so the
// two outputs can be compared file by file.
//
// Usage: dotnet run -- <folder containing NAME-original.docx / NAME-modified.docx>
// Writes NAME-dotnet.docx next to each pair.
using System.Diagnostics;
using Docxodus;

var folder = args.Length > 0 ? args[0] : ".";
foreach (var original in Directory.GetFiles(folder, "*-original.docx").OrderBy(p => p))
{
    var name = Path.GetFileName(original).Replace("-original.docx", "");
    var modified = Path.Combine(folder, name + "-modified.docx");
    var clock = Stopwatch.StartNew();
    var redline = DocxCompare.Compare(
        new WmlDocument(original),
        new WmlDocument(modified),
        new DocxDiffSettings { AuthorForRevisions = "Redline" });
    clock.Stop();
    redline.SaveAs(Path.Combine(folder, name + "-dotnet.docx"));
    Console.WriteLine($"{name}: {clock.ElapsedMilliseconds} ms");
}
