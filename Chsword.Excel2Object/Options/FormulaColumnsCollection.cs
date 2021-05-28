using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using Chsword.Excel2Object.Functions;

namespace Chsword.Excel2Object.Options
{
    public class FormulaColumnsCollection : ICollection<FormulaColumn>
    {
        public int Count => FormulaColumns.Count;
        public List<FormulaColumn> FormulaColumns { get; set; } = new List<FormulaColumn>();

        public bool IsReadOnly => false;

        public void Add(string columnTitle, Expression<Func<ColumnCellDictionary, object>> func)
        {
            FormulaColumns.Add(new FormulaColumn()
            {
                Title = columnTitle,
                Formula = func
            });
        }

        public void Add(FormulaColumn item)
        {
            if (item == null)
            {
                throw new Excel2ObjectException("item must not null");
            }

            if (FormulaColumns.Any(c => c.Title == item.Title))
            {
                throw new Excel2ObjectException("same title has existsed in options.FormulaColumns");
            }

            if (!string.IsNullOrWhiteSpace(item.AfterColumnTitle) &&
                FormulaColumns.Any(c => c.AfterColumnTitle == item.AfterColumnTitle))
            {
                throw new Excel2ObjectException($"There is a formula column after {item.AfterColumnTitle} already.");
            }

            FormulaColumns.Add(item);
        }

        public void Clear()
        {
            FormulaColumns.Clear();
        }

        public bool Contains(FormulaColumn item)
        {
            return FormulaColumns.Contains(item);
        }

        public void CopyTo(FormulaColumn[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public IEnumerator<FormulaColumn> GetEnumerator()
        {
            return FormulaColumns.GetEnumerator();
        }

        public bool Remove(FormulaColumn item)
        {
            return FormulaColumns.Remove(item);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return FormulaColumns.GetEnumerator();
        }
    }
}