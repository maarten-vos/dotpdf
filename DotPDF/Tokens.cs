namespace DotPDF
{
    internal static class Tokens
    {

        /// <summary>
        /// A token which is used to define children on objects that support it
        /// </summary>
        public const string Children = nameof(Children);

        /// <summary>
        /// Tokens that are used at the root of the object
        /// </summary>
        public const string PageSetup = nameof(PageSetup);
        public const string Styles = nameof(Styles);

        /// <summary>
        /// Tokens that are related to creating tables
        /// </summary>
        public const string Columns = "@Columns";
        public const string Rows = "@Rows";
        public const string Cells = "@Cells";
        public const string Row = "@Row";
        public const string Column = "@Column";

        /// <summary>
        /// Tokens that are used as properties
        /// </summary>
        public const string Condition = nameof(Condition);
        public const string Type = nameof(Type);
        public const string Text = nameof(Text);
        public const string CompiledText = nameof(CompiledText);
        public const string Color = nameof(Color);
        public const string ForEach = nameof(ForEach);

        /// <summary>
        /// Tokens that define which types you can use
        /// </summary>
        public const string Table = nameof(Table);
        public const string Paragraph = nameof(Paragraph);
        public const string Footer = nameof(Footer);
        public const string Header = nameof(Header);
        public const string Image = nameof(Image);
        public const string TextFrame = nameof(TextFrame);
        public const string FormattedText = nameof(FormattedText);
        public const string PageBreak = nameof(PageBreak);
        public const string PageField = nameof(PageField);
        public const string NumPagesField = nameof(NumPagesField);

        /// <summary>
        /// Legacy tokens, we advise you no longer use these
        /// </summary>
        public const string LegacyCompiledText = "@Text";
        public const string LegacyColor = "@Color";
        public const string LegacyRepeat = "@Repeat";
    }
}
