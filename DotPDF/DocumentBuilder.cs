using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using CSScriptLibrary;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;

namespace DotPDF
{
    public class DocumentBuilder
    {
        private readonly Globals _globals = new Globals();

        private readonly Dictionary<string, dynamic> _cache = new Dictionary<string, dynamic>();

        private readonly Dictionary<Tuple<string, Type>, Delegate> _dictionary = new Dictionary<Tuple<string, Type>, Delegate>();

        private static readonly MethodInfo _setPropertyMethod;

        static DocumentBuilder()
        {
            _setPropertyMethod = typeof(DocumentBuilder).GetMethod("SetProperty", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.FlattenHierarchy);
        }

        public PdfDocumentRenderer GetDocumentRenderer(JToken token, string templateJson)
        {
            var pdfRenderer = new PdfDocumentRenderer(true);
            var document = CreateDocument(token, templateJson);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();
            return pdfRenderer;
        }

        public Document CreateDocument(JToken token, string templateJson)
        {
            var template = JObject.Parse(templateJson);
            var pdfDocument = new Document();
            _globals.Obj = token;

            void Start()
            {
                var section = pdfDocument.AddSection();
                var pageSetup = (JObject)template[Tokens.PageSetup];
                var styles = (JObject)template[Tokens.Styles];
                if (styles != null)
                    SetDefaultProperties(pdfDocument.Styles, styles);
                if (pageSetup != null)
                    SetDefaultProperties(section.PageSetup, pageSetup);
                ParseChildren(section, (JArray)template[Tokens.Children]);
            }

            if (token.Type == JTokenType.Array)
            {
                _globals.Array = (JArray)token;
                foreach (var e in token)
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
                if (child[Tokens.Condition] != null)
                    if (!Compile<bool>((string)child[Tokens.Condition]))
                        continue;

                var loop = child[Tokens.Repeat] != null ? Compile<List<object>>((string)child[Tokens.Repeat]) : new List<object> { _globals.Item };

                switch ((string)child[Tokens.Type])
                {
                    case Tokens.Table:
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            SetTable(obj.AddTable(), (JObject)child);
                        }
                        break;
                    case Tokens.Paragraph:
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            SetDefaultProperties(obj.AddParagraph(), (JObject)child);
                        }
                        break;
                    case Tokens.Footer:
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            SetTable(obj.Footers.Primary.AddTable(), (JObject)child);
                        }
                        break;
                    case Tokens.Header:
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            SetTable(obj.Headers.Primary.AddTable(), (JObject)child);
                        }
                        break;
                    case Tokens.Image:
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            var source = (string)child["Source"];
                            var img = obj.AddImage(source.StartsWith("base64:") ? source : ImageToBase64(Compile<byte[]>((string)child["Source"])));
                            SetDefaultProperties(img, (JObject)child);
                        }
                        break;
                    case Tokens.TextFrame:
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            SetDefaultProperties(obj.AddTextFrame(), (JObject)child);
                        }
                        break;
                    case Tokens.FormattedText:
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            SetDefaultProperties(obj.AddFormattedText(), (JObject)child);
                        }
                        break;
                    case Tokens.PageBreak:
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            obj.AddPageBreak();
                        }
                        break;
                    case Tokens.PageField:
                        foreach (var item in loop)
                        {
                            _globals.Item = item;
                            obj.AddPageField();
                        }
                        break;
                    default:
                        throw new NotImplementedException("Unknown type: " + (string)child[Tokens.Type] + "\n\n" + child);
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
            foreach (var column in child[Tokens.Columns])
            {
                var addColumn = table.AddColumn();
                foreach (var item in ((JObject)column).Properties())
                    SetProperty(addColumn, item);
            }
            var rowIndex = 0;
            var virtualRows = new List<VirtualRow>();
            foreach (var row in child[Tokens.Rows])
            {
                if (row[Tokens.Repeat] != null)
                {
                    var virtualRow = new VirtualRow
                    {
                        Items = Compile<IList>((string)row[Tokens.Repeat]),
                        RowIndex = rowIndex
                    };
                    virtualRows.Add(virtualRow);
                    foreach (var unused in virtualRow.Items)
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
            foreach (var cell in child[Tokens.Cells])
            {
                var row = (int)cell[Tokens.Row];
                var virtualRow = virtualRows.Find(v => v.RowIndex == row);
                if (virtualRow != null)
                {
                    var items = virtualRow.Items;
                    foreach (var kid in items)
                    {
                        _globals.Item = kid;
                        foreach (var item in ((JObject)cell).Properties())
                            SetProperty(table[row + items.IndexOf(kid), (int)cell[Tokens.Column]], item);
                    }
                }
                else
                {
                    foreach (var item in ((JObject)cell).Properties())
                        SetProperty(table[row, (int)cell[Tokens.Column]], item);
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
                    var genericSubMethod = _setPropertyMethod.MakeGenericMethod(value.GetType());
                    foreach (var subProperty in ((JObject)property.Value).Properties())
                        genericSubMethod.Invoke(this, new[] { value, subProperty });
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
                    case Tokens.Text:
                        {
                            dynamic text = parent;
                            var value = (string)property.Value;
                            if (value != null)
                                text.AddText(value);
                            break;
                        }
                    case Tokens.CompiledText:
                        {
                            dynamic text = parent;
                            var value = Compile<string>((string)property.Value);
                            if (value != null)
                                text.AddText(value);

                            break;
                        }
                    case Tokens.Children:
                        ParseChildren(parent, (JArray)property.Value);
                        break;
                    case Tokens.Color:
                        parent.GetType().GetProperty("Color").SetValue(parent, Color.Parse((string)property.Value));
                        break;
                }
            }
        }

        private static string ImageToBase64(byte[] image)
        {
            return "base64:" + Convert.ToBase64String(image);
        }

        private TR Compile<TR>(string code)
        {
            if (!_cache.TryGetValue(code, out var myClass))
            {
                var newCode = $"using DotPDF; class Stub {{ public object Eval(Globals globals) {{ var Obj = globals.Obj; var Array = globals.Array; var Item = globals.Item; return {code}; }} }}";
                _cache[code] = myClass = CSScript.Evaluator.LoadCode(newCode);
            }

            return (TR)myClass.Eval(_globals);
        }
    }
}
