using System;
using System.Collections.Generic;
using System.Linq;

namespace WordTemplates
{
    public class TemplateEngine
    {
        private readonly ITemplateNavigator _doc;

        public TemplateEngine(ITemplateNavigator doc)
        {
            _doc = doc;
        }

        public void SetFields(IEnumerable<TemplateField> fields)
        {
            foreach (var group in fields.GroupBy(o => o.Group).ToList())
            {
                var iteratedGroup = group
                    .Where(o => o.Iteration.HasValue)
                    .OrderBy(o => o.Iteration)
                    .ThenBy(o => o.Name)
                    .GroupBy(o => o.Iteration)
                    .ToList();

                if (iteratedGroup.Any())
                    HandleIteratedGroup(iteratedGroup);
                else
                    HandleSimpleFields(group);
            }
        }

        private void HandleIteratedGroup(IList<IGrouping<int?, TemplateField>> iterations)
        {
            var columns = iterations[0].Select(o => o.Name).ToArray();
            var rows = new object[iterations.Count][];
            for (int i = 0; i < iterations.Count; i++)
            {
                var row = iterations[i].Select(o => o.Value).ToArray();
                if (row.Length != columns.Length)
                    throw new NotSupportedException(
                        string.Format(
                            "Iteration ({2}) has different number of fields ({0}) than expected ({1})",
                            row.Length,
                            columns.Length,
                            i));

                rows[i] = row;
            }

            _doc.SetFields(columns, rows);
        }

        private void HandleSimpleFields(IEnumerable<TemplateField> fields)
        {
            foreach (var field in fields)
                _doc.SetField(field.Name, field.Value);
        }
    }
}