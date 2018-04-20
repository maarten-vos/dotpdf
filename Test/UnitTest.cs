using System;
using DotPDF;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json.Linq;
using System.IO;

namespace Test
{
    [TestClass]
    public class UnitTest
    {
        public const string TemplatesDir = "../../Templates/";

        [TestMethod]
        public void TestBasicTemplate()
        {
            var templateJson = File.ReadAllText($"{TemplatesDir}basic_template.json");
            var data = JObject.FromObject(new {Name = "John Smith", Time = DateTime.Now});

            var builder = new DocumentBuilder();

            builder.GetDocumentRenderer(data, templateJson).Save("./doc.pdf");
        }
    }
}
