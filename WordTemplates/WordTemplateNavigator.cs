using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace WordTemplates
{
    public class WordTemplateNavigator : ITemplateNavigator
    {
        // i.e.: CompanyName, CompanyName_1 etc.
        private static readonly Regex BookmarkNameRegex =
            new Regex(@"(?<name>[a-zA-Z0-9]+)(?<number>_\d+)?");

        private readonly _Application _word;
        private readonly _Document _doc;

        private readonly Dictionary<string, List<Bookmark>> _bookmarksByName =
            new Dictionary<string, List<Bookmark>>();

        public WordTemplateNavigator(_Application word, _Document doc)
        {
            _word = word;
            _doc = doc;

            GroupSynonymousBookmarks();
        }

        private void GroupSynonymousBookmarks()
        {
            foreach (Bookmark bookmark in _doc.Bookmarks)
            {
                var match = BookmarkNameRegex.Match(bookmark.Name);
                if (match.Success)
                {
                    var name = match.Groups["name"].Value;
                    if (_bookmarksByName.ContainsKey(name))
                        _bookmarksByName[name].Add(bookmark);
                    else
                        _bookmarksByName[name] = new List<Bookmark> { bookmark };
                }
            }
        }

        public void SetField(string field, object value)
        {
            List<Bookmark> bookmarks;
            if (_bookmarksByName.TryGetValue(field, out bookmarks))
            {
                foreach (var bookmark in bookmarks)
                    SetBookmarkValue(bookmark, value);
            }
        }

        public void SetFields(string[] fields, object[][] values)
        {
            if (!SelectRowWithFields(fields)) return;

            _word.Selection.Range.Copy();

            var table = _word.Selection.Tables[1];
            var lastDataRow = _word.Selection.Rows[1].Index;
            for (int rowIndex = values.Length - 1; rowIndex >= 0; rowIndex--)
            {
                FillRow(fields, values[rowIndex]);

                if (rowIndex > 0)
                    table.Rows[lastDataRow].Range.Paste();
            }
        }

        private bool SelectRowWithFields(string[] columns)
        {
            var bookmark = _bookmarksByName
                .Where(o => columns.Contains(o.Key))
                .Select(o => o.Value[0])
                .FirstOrDefault();

            if (bookmark == null)
                return false;

            bookmark.Select();
            _word.Selection.SelectRow();

            if (_word.Selection.Tables.Count != 1 || _word.Selection.Rows.Count != 1)
                throw new NotSupportedException(
                    "Not supported bookmark location '" + bookmark.Name + "'");

            return true;
        }

        private void FillRow(string[] columns, object[] row)
        {
            for (int columnIndex = 0; columnIndex < columns.Length; columnIndex++)
            {
                var name = columns[columnIndex];
                if (_doc.Bookmarks.Exists(name))
                    SetBookmarkValue(_doc.Bookmarks[name], row[columnIndex]);
            }
        }

        private void SetBookmarkValue(Bookmark bookmark, object value)
        {
            // no need to keep bookmarks and they get in the way in tables
            var checkBox = _doc.FormFields[bookmark.Name].CheckBox;
            if (checkBox.Valid)
            {
                bookmark.Delete();
                checkBox.Value = value.ToString() == Boolean.TrueString;
            }
            else
            {
                bookmark.Select();
                _word.Selection.TypeText(value.ToString());
            }
        }
    }
}
