using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using CSScriptLibrary;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;

namespace DotPDF
{
    public class DocumentBuilder : DocumentBuilder<DocumentBuilder.Globals>
    {
        public class Globals : IGlobals
        {
            public dynamic Obj { get; set; }
            public dynamic Item { get; set; }
            public JArray Array { get; set; }

            public bool IsEven(JArray array)
            {
                return array.IndexOf(Item) % 2 == 0;
            }
            
            public int IndexOf(JArray array, dynamic obj)
            {
                for (var i = 0; i < array.Count; i++)
                {
                    var e = array[i];
                    if (e == obj)
                        return i;
                }
                return -1;
            }
        }
    }

    public class DocumentBuilder<T> where T : class, IGlobals, new()
    {
        private readonly T _globals = new T();
        private readonly Dictionary<Tuple<string, Type>, Delegate> _dictionary = new Dictionary<Tuple<string, Type>, Delegate>();

        public PdfDocumentRenderer GetDocumentRenderer(dynamic obj, string templateJson)
        {
            var pdfRenderer = new PdfDocumentRenderer(true);
            var document = CreateDocument(obj, templateJson);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();
            return pdfRenderer;
        }

        private Document CreateDocument(JToken obj, string templateJson)
        {
            var template = JObject.Parse(templateJson);
            var pdfDocument = new Document();
            _globals.Obj = obj;

            void Start()
            {
                var section = pdfDocument.AddSection();
                var pageSetup = (JObject)template["PageSetup"];
                var styles = (JObject)template["Styles"];
                if (styles != null)
                    SetDefaultProperties(pdfDocument.Styles, styles);
                if (pageSetup != null)
                    SetDefaultProperties(section.PageSetup, pageSetup);
                ParseChildren(section, (JArray)template["Children"]);
            }

            if (obj.Type == JTokenType.Array)
            {
                _globals.Array = (JArray) obj;
                foreach (var e in obj)
                {
                    _globals.Obj = e;
                    Start();
                }
            }
            else
                Start();

            return pdfDocument;
        }

        private void ParseChildren(dynamic obj, JArray children)
        {
            foreach (var child in children)
            {
                if (child["Condition"] != null)
                    if (!Compile<bool>((string)child["Condition"]))
                        continue;

                var loop = child["@Repeat"] != null ? Compile<List<object>>((string)child["@Repeat"]) : new List<object> {_globals.Item};

                switch ((string)child["Type"])
                {
                    case "Table":
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            SetTable(obj.AddTable(), (JObject)child);
                        }
                        break;
                    case "Paragraph":
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            SetDefaultProperties(obj.AddParagraph(), (JObject)child);
                        }
                        break;
                    case "Footer":
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            SetTable(obj.Footers.Primary.AddTable(), (JObject)child);
                        }
                        break;
                    case "Header":
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            SetTable(obj.Headers.Primary.AddTable(), (JObject)child);
                        }
                        break;
                    case "Image":
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            var source = (string)child["Source"];
                            var img = obj.AddImage(source.StartsWith("base64:") ? source : ImageToBase64(Compile<byte[]>((string)child["Source"])));
                            SetDefaultProperties(img, (JObject)child);
                        }
                        break;
                    case "TextFrame":
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            SetDefaultProperties(obj.AddTextFrame(), (JObject)child);
                        }
                        break;
                    case "FormattedText":
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            SetDefaultProperties(obj.AddFormattedText(), (JObject)child);
                        }
                        break;
                    case "PageBreak":
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            obj.AddPageBreak();
                        }
                        break;
                    case "PageField":
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            obj.AddPageField();
                        }
                        break;
                    default:
                        throw new NotImplementedException("Unknown type: " + (string)child["Type"] + "\n\n" + child);
                }
            }
        }

        private void SetDefaultProperties<TParent>(TParent obj, JObject child)
        {
            foreach (var item in child.Properties())
                SetProperty(obj, item);
        }

        private void SetTable(Table table, JObject child)
        {
            foreach (var column in child["@Columns"])
            {
                var c = table.AddColumn();
                foreach (var item in ((JObject)column).Properties())
                    SetProperty(c, item);
            }
                

            var rowIndex = 0;
            var virtualRows = new List<VirtualRow>();
            foreach (var row in child["@Rows"])
            {
                if (row["@Repeat"] != null)
                {
                    var vr = new VirtualRow
                    {
                        Kids = Compile<IList>((string) row["@Repeat"]),
                        RowIndex = rowIndex
                    };
                    virtualRows.Add(vr);
                    foreach (var unused in vr.Kids)
                    {
                        var parent = table.AddRow();
                        foreach (var item in ((JObject)row).Properties())
                            SetProperty(parent, item);
                    }
                }
                else
                {
                    var parent = table.AddRow();
                    foreach (var item in ((JObject)row).Properties())
                        SetProperty(parent, item);
                }
                rowIndex++;
            }

            foreach (var cell in child["@Cells"])
            {
                var row = (int)cell["@Row"];
                var vr = virtualRows.Find(v => v.RowIndex == row);
                if (vr != null)
                {
                    var kids = vr.Kids;
                    foreach (var kid in kids)
                    {
                        _globals.Item = kid;
                        foreach (var item in ((JObject)cell).Properties())
                            SetProperty(table[row + kids.IndexOf(kid), (int)cell["@Column"]], item);
                    }
                }
                else
                {
                    foreach (var item in ((JObject)cell).Properties())
                        SetProperty(table[row, (int)cell["@Column"]], item);
                }


            }
            SetDefaultProperties(table, child);
        }

        protected void SetProperty<TParent>(TParent parent, JProperty property)
        {
            var objProperty = parent.GetType().GetProperty(property.Name);
            if (objProperty != null)
            {
                if (property.Value.Type == JTokenType.Object)
                {
                    var value = objProperty.GetValue(parent);
                    foreach (var subProperty in ((JObject)property.Value).Properties())
                        SetProperty(value, subProperty);
                }
                else if (objProperty.PropertyType == typeof(Unit))
                {
                    var key = Tuple.Create(property.Name, typeof(TParent));
                    if (!_dictionary.TryGetValue(key, out var del))
                    {
                        var newDel = (Action<TParent, Unit>)Delegate.CreateDelegate(typeof(Action<TParent, Unit>), objProperty.GetSetMethod());
                        _dictionary.Add(key, newDel);
                        del = newDel;
                    }

                    var genericDel = (Action<TParent, Unit>)del;
                    genericDel(parent, Unit.Parse((string)property.Value));
                }
                else if (objProperty.PropertyType == typeof(string))
                {
                    var key = Tuple.Create(property.Name, typeof(TParent));
                    if (!_dictionary.TryGetValue(key, out var del))
                    {
                        var newDel = (Action<TParent, string>)Delegate.CreateDelegate(typeof(Action<TParent, string>), objProperty.GetSetMethod());
                        _dictionary.Add(key, newDel);
                        del = newDel;
                    }

                    var genericDel = (Action<TParent, string>)del;
                    genericDel(parent, (string)property.Value);
                }
                else if (objProperty.PropertyType == typeof(bool))
                {
                    var key = Tuple.Create(property.Name, typeof(TParent));
                    if (!_dictionary.TryGetValue(key, out var del))
                    {
                        var newDel = (Action<TParent, bool>)Delegate.CreateDelegate(typeof(Action<TParent, bool>), objProperty.GetSetMethod());
                        _dictionary.Add(key, newDel);
                        del = newDel;
                    }

                    var genericDel = (Action<TParent, bool>)del;
                    genericDel(parent, (bool)property.Value);
                }
                else if (objProperty.PropertyType == typeof(Color))
                {
                    var key = Tuple.Create(property.Name, typeof(TParent));
                    if (!_dictionary.TryGetValue(key, out var del))
                    {
                        var newDel = (Action<TParent, Color>)Delegate.CreateDelegate(typeof(Action<TParent, Color>), objProperty.GetSetMethod());
                        _dictionary.Add(key, newDel);
                        del = newDel;
                    }

                    var genericDel = (Action<TParent, Color>)del;
                    genericDel(parent, Color.Parse((string)property.Value));
                }
                else if (objProperty.PropertyType.IsEnum)
                    objProperty.SetValue(parent, Enum.Parse(objProperty.PropertyType, (string)property.Value));
                else if (objProperty.PropertyType == typeof(TopPosition))
                {
                    var key = Tuple.Create(property.Name, typeof(TParent));
                    if (!_dictionary.TryGetValue(key, out var del))
                    {
                        var newDel = (Action<TParent, TopPosition>)Delegate.CreateDelegate(typeof(Action<TParent, TopPosition>), objProperty.GetSetMethod());
                        _dictionary.Add(key, newDel);
                        del = newDel;
                    }

                    var genericDel = (Action<TParent, TopPosition>)del;
                    genericDel(parent, TopPosition.Parse((string)property.Value));
                }
                else if (objProperty.PropertyType == typeof(LeftPosition))
                {
                    var key = Tuple.Create(property.Name, typeof(TParent));
                    if (!_dictionary.TryGetValue(key, out var del))
                    {
                        var newDel = (Action<TParent, LeftPosition>)Delegate.CreateDelegate(typeof(Action<TParent, LeftPosition>), objProperty.GetSetMethod());
                        _dictionary.Add(key, newDel);
                        del = newDel;
                    }

                    var genericDel = (Action<TParent, LeftPosition>)del;
                    genericDel(parent, LeftPosition.Parse((string)property.Value));
                }
                else
                    throw new NotSupportedException($"Unknown property type: {objProperty.PropertyType}");
            }
            else
            {
                switch (property.Name)
                {
                    case "Text":
                        {
                            dynamic text = parent;
                            var value = (string) property.Value;
                            if (value != null)
                                text.AddText(value);
                            break;
                        }
                    case "@Text":
                        {
                            dynamic text = parent;
                            var value = Compile<string>((string) property.Value);
                            if (value != null)
                                text.AddText(value);

                            break;
                        }
                    case "Children":
                        ParseChildren(parent, (JArray)property.Value);
                        break;
                    case "@Color":
                        var p = parent.GetType().GetProperty(property.Name.Substring(1));
                        p.SetValue(parent, Color.Parse(Compile<string>((string)property.Value)));
                        break;
                }
            }
        }

        private static string ImageToBase64(byte[] image)
        {
            return "base64:" + Convert.ToBase64String(image);
        }

        private T Compile<T>(string code)
        {
            var script = CSScript.Evaluator.CreateDelegate(code);
            return (T) script.Invoke(_globals);
        }
    }

}
