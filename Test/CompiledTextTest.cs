using System;
using DotPDF;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json.Linq;
using System.IO;

namespace Test
{
    [TestClass]
    public class CompiledTextTest
    {
        public const string TemplatesDir = "../../Templates";

        [TestMethod]
        public void TestCompiledText()
        {
            SaveDocument("compiled_text", new { Name = "John Smith", Time = DateTime.Now });
        }

        private void SaveDocument(string templateName, object data)
        {
            var templateJson = File.ReadAllText($"{TemplatesDir}/{templateName}.json");
            var builder = new DocumentBuilder();

            builder.GetDocumentRenderer(JObject.FromObject(data), templateJson).Save($"./{templateName}.pdf");
        }
    }
}
