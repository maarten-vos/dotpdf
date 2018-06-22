using CSScriptLibrary;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace DotPDF
{
    public class DocumentBuilder
    {
        private readonly Globals _globals = new Globals();

        private readonly Dictionary<string, dynamic> _cache = new Dictionary<string, dynamic>();

        private readonly Dictionary<Tuple<string, Type>, Delegate> _dictionary = new Dictionary<Tuple<string, Type>, Delegate>();

        private static readonly MethodInfo _setPropertyMethod;

        public string[] Imports { get; set; }

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

        private void ParseChildren(dynamic currentObj, JArray children)
        {
            foreach (JObject child in children)
            {
                if (child[Tokens.Condition] != null)
                    if (!Compile<bool>((string)child[Tokens.Condition]))
                        continue;

                var forEach = child[Tokens.LegacyRepeat] ?? child[Tokens.ForEach];

                var items = forEach != null
                    ? Compile<List<object>>((string)forEach)
                    : new List<object> { _globals.Item };

                foreach (var item in items)
                {
                    _globals.Item = item;
                    switch ((string)child[Tokens.Type])
                    {
                        case Tokens.Table:
                            SetTable(currentObj.AddTable(), child);
                            break;
                        case Tokens.Paragraph:
                            if (currentObj is Paragraph)
                                SetDefaultProperties(currentObj, child);
                            else
                                SetDefaultProperties(currentObj.AddParagraph(), child);
                            break;
                        case Tokens.Footer:
                            SetDefaultProperties(((Section)currentObj).Footers.Primary, child);
                            break;
                        case Tokens.Header:
                            SetDefaultProperties(((Section)currentObj).Headers.Primary, child);
                            break;
                        case Tokens.Image:
                            var source = (string)child["Source"];
                            var imageData = source.StartsWith("base64:")
                                ? source
                                : ImageToBase64(Compile<byte[]>((string)child["Source"]));

                            var image = currentObj.AddImage();
                            SetDefaultProperties(image, child);
                            break;
                        case Tokens.TextFrame:
                            SetDefaultProperties(currentObj.AddTextFrame(), child);
                            break;
                        case Tokens.FormattedText:
                            SetDefaultProperties(currentObj.AddFormattedText(), child);
                            break;
                        case Tokens.PageBreak:
                            currentObj.AddPageBreak();
                            break;
                        case Tokens.PageField:
                            currentObj.AddPageField();
                            break;
                        case Tokens.NumPagesField:
                            currentObj.AddNumPagesField();
                            break;
                        default:
                            throw new NotSupportedException($"Unknown type: {(string)child[Tokens.Type]} \n\n {child}");
                    }
                }

            }
        }

        private void SetDefaultProperties<T>(T parent, JObject child)
        {
            foreach (var item in child.Properties())
                SetProperty(parent, item);
        }

        private void SetTable(Table table, JObject child)
        {
            foreach (JObject column in child[Tokens.Columns])
                SetDefaultProperties(table.AddColumn(), column);

            var rowIndex = 0;
            var tableRows = new List<TableRow>();
            foreach (JObject row in child[Tokens.Rows])
            {
                var forEach = row[Tokens.LegacyRepeat] ?? row[Tokens.ForEach];
                if (forEach != null)
                {
                    var tableRow = new TableRow
                    {
                        Items = Compile<IList>((string)forEach),
                        RowIndex = rowIndex
                    };
                    tableRows.Add(tableRow);
                    foreach (var item in tableRow.Items)
                        SetDefaultProperties(table.AddRow(), row);
                }
                else
                    SetDefaultProperties(table.AddRow(), row);

                rowIndex++;
            }

            foreach (JObject cell in child[Tokens.Cells])
            {
                var rowId = (int)cell[Tokens.Row];
                var tableRow = tableRows.FirstOrDefault(v => v.RowIndex == rowId);
                if (tableRow != null)
                {
                    foreach (var item in tableRow.Items)
                    {
                        _globals.Item = item;
                        SetDefaultProperties(table[rowId + tableRow.Items.IndexOf(item), (int)cell[Tokens.Column]], cell);
                    }
                }
                else
                    SetDefaultProperties(table[rowId, (int)cell[Tokens.Column]], cell);


            }
            SetDefaultProperties(table, child);
        }

        protected void SetProperty<T>(T parent, JProperty property)
        {
            switch (property.Name)
            {
                case Tokens.Cells:
                case Tokens.Rows:
                case Tokens.Columns:
                case Tokens.Condition:
                case Tokens.Column:
                case Tokens.Row:
                case Tokens.ForEach:
                case Tokens.Type:
                    break;

                case Tokens.Text:
                    {
                        dynamic text = parent;
                        var value = (string)property.Value;
                        if (value != null)
                            text.AddText(value);
                        break;
                    }
                case Tokens.CompiledText:
                case Tokens.LegacyCompiledText:
                    {
                        dynamic text = parent;
                        var value = Compile<string>((string)property.Value);
                        if (value != null)
                            text.AddText(value);

                        break;
                    }
                case Tokens.Children:
                    {
                        ParseChildren(parent, (JArray)property.Value);
                        break;
                    }
                case Tokens.Color:
                case Tokens.LegacyColor:
                    {
                        var color = Color.Parse((string)property.Value);
                        parent.GetType()
                            .GetProperty(Tokens.Color)
                            .SetValue(parent, color);
                        break;
                    }
                default:
                    {
                        var info = parent.GetType().GetProperty(property.Name)
                            ?? throw new NotSupportedException($"Invalid property: {property.Name}");

                        SetPdfSharpProperty(parent, info, property);
                        break;
                    }
            }

        }

        private void SetPdfSharpProperty<T>(T parent, PropertyInfo info, JProperty property)
        {
            if (property.Value.Type == JTokenType.Object)
            {
                var value = info.GetValue(parent);
                var genericSubMethod = _setPropertyMethod.MakeGenericMethod(value.GetType());
                foreach (var subProperty in ((JObject)property.Value).Properties())
                    genericSubMethod.Invoke(this, new[] { value, subProperty });
            }
            else if (info.PropertyType.IsEnum)
                info.SetValue(parent, Enum.Parse(info.PropertyType, (string)property.Value));
            else if (info.PropertyType == typeof(Unit))
            {
                var key = Tuple.Create(property.Name, typeof(T));
                if (!_dictionary.TryGetValue(key, out var del))
                {
                    var newDel = (Action<T, Unit>)Delegate.CreateDelegate(typeof(Action<T, Unit>), info.GetSetMethod());
                    _dictionary.Add(key, newDel);
                    del = newDel;
                }

                var genericDel = (Action<T, Unit>)del;
                genericDel(parent, Unit.Parse((string)property.Value));
            }
            else if (info.PropertyType == typeof(string))
            {
                var key = Tuple.Create(property.Name, typeof(T));
                if (!_dictionary.TryGetValue(key, out var del))
                {
                    var newDel = (Action<T, string>)Delegate.CreateDelegate(typeof(Action<T, string>), info.GetSetMethod());
                    _dictionary.Add(key, newDel);
                    del = newDel;
                }

                var genericDel = (Action<T, string>)del;
                genericDel(parent, (string)property.Value);
            }
            else if (info.PropertyType == typeof(bool))
            {
                var key = Tuple.Create(property.Name, typeof(T));
                if (!_dictionary.TryGetValue(key, out var del))
                {
                    var newDel = (Action<T, bool>)Delegate.CreateDelegate(typeof(Action<T, bool>), info.GetSetMethod());
                    _dictionary.Add(key, newDel);
                    del = newDel;
                }

                var genericDel = (Action<T, bool>)del;
                genericDel(parent, (bool)property.Value);
            }
            else if (info.PropertyType == typeof(Color))
            {
                var key = Tuple.Create(property.Name, typeof(T));
                if (!_dictionary.TryGetValue(key, out var del))
                {
                    var newDel = (Action<T, Color>)Delegate.CreateDelegate(typeof(Action<T, Color>), info.GetSetMethod());
                    _dictionary.Add(key, newDel);
                    del = newDel;
                }

                var genericDel = (Action<T, Color>)del;
                genericDel(parent, Color.Parse((string)property.Value));
            }
            else if (info.PropertyType == typeof(TopPosition))
            {
                var key = Tuple.Create(property.Name, typeof(T));
                if (!_dictionary.TryGetValue(key, out var del))
                {
                    var newDel = (Action<T, TopPosition>)Delegate.CreateDelegate(typeof(Action<T, TopPosition>), info.GetSetMethod());
                    _dictionary.Add(key, newDel);
                    del = newDel;
                }

                var genericDel = (Action<T, TopPosition>)del;
                genericDel(parent, TopPosition.Parse((string)property.Value));
            }
            else if (info.PropertyType == typeof(LeftPosition))
            {
                var key = Tuple.Create(property.Name, typeof(T));
                if (!_dictionary.TryGetValue(key, out var del))
                {
                    var newDel = (Action<T, LeftPosition>)Delegate.CreateDelegate(typeof(Action<T, LeftPosition>), info.GetSetMethod());
                    _dictionary.Add(key, newDel);
                    del = newDel;
                }

                var genericDel = (Action<T, LeftPosition>)del;
                genericDel(parent, LeftPosition.Parse((string)property.Value));
            }
            else
                throw new NotSupportedException($"Unknown property type: {info.PropertyType}");
        }

        private static string ImageToBase64(byte[] image)
            => $"base64:{Convert.ToBase64String(image)}";

        private T Compile<T>(string code)
        {
            if (!_cache.TryGetValue(code, out var myClass))
            {
                var usings = Imports?.Select(s => $"using {s};");
                var imports = usings == null ? string.Empty : string.Join(" ", usings);
                var newCode = $"using DotPDF; {imports} class Stub {{ public object Eval(Globals globals) {{ var Obj = globals.Obj; var Array = globals.Array; var Item = globals.Item; return {code}; }} }}";
                _cache[code] = myClass = CSScript.Evaluator.LoadCode(newCode);
            }

            return (T)myClass.Eval(_globals);
        }
    }
}
