using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
var validator = new OpenXmlValidator();
using var doc = PresentationDocument.Open(@"output\sample-slide-narrated3.pptx", false);
var errors = validator.Validate(doc);
Console.WriteLine($"Total errors: {errors.Count()}");
foreach (var e in errors.Take(30))
{
    Console.WriteLine($"---");
    Console.WriteLine($"Part: {e.Part?.Uri}");
    Console.WriteLine($"Path: {e.Path?.XPath}");
    Console.WriteLine($"Id:   {e.Id}");
    Console.WriteLine($"Desc: {e.Description}");
}
