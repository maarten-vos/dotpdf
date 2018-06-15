using System;
using DotPDF;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Dynamic;
using System.Linq;

namespace Test
{
    [TestClass]
    public class ProductList
    {

        [TestMethod]
        public void TestProductList()
        {
            var products = Enumerable.Range(0, 50).Select(n => new
            {
                Description = $"Product {n}",
                Quantity = 3,
                Price = 2.00M
            }).ToArray();

            SaveDocument("product_list", new
            {
                Products = products,
                Total = products.Sum(p => p.Price * p.Quantity)
            });
        }

        private void SaveDocument(string templateName, object data)
        {
            var templateJson = File.ReadAllText($"../../Templates/{templateName}.json");
            var builder = new DocumentBuilder
            {
                Imports = new[] { "System.Linq" }
            };

            builder.GetDocumentRenderer(JObject.FromObject(data), templateJson).Save($"./{templateName}.pdf");
        }

    }
}
