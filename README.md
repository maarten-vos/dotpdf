# dotpdf [![dotpdf MyGet Build Status](https://www.myget.org/BuildSource/Badge/dotpdf?identifier=0cccd0c6-a15d-4650-9fd8-a11ce5a61875)](https://www.myget.org/)
*Create PDF documents using templates written in JSON*

Currently there isn't much documentation for dotpdf, but because it uses pdfsharp for document creation its documentation is a good source of information.

## Features
* Standard static templates
* Data injection during runtime

## Installation

Install it using NuGet!

Just add `https://www.myget.org/F/dotpdf/api/v2` as a package source in your NuGet package manager.

## Usage
First you need a template, here's a basic one as an example:
```JSON
{
  "PageSetup": {
    "LeftMargin": "1cm",
    "RightMargin": "1cm",
    "TopMargin": "1cm"
  },
  "Children": [
    {
      "Type": "Paragraph",
      "@Text": "\"Hello, \" + Obj.Name"
    }
  ]
}
```
After that you need to create your `DocumentBuilder` and provide the required arguments to the `GetDocumentRenderer` method.
```CSharp
var template = File.ReadAllText($"./basic_template.json");
var data = JObject.FromObject(new {Name = "John Smith", Time = DateTime.Now});
var builder = new DocumentBuilder();
builder.GetDocumentRenderer(data, templateJson).Save("./document.pdf");
```