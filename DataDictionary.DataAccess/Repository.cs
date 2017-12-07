using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataDictionary.DataAccess
{
    public class Repository
    {
		public void AddEntry(DataDictionaryEntity DataDict)
		{
			using (var context = new Context())
			{
				context.DBDicts.Add(DataDict);
				context.SaveChanges();
			}
		}

		public IEnumerable<DataDictionaryEntity> GetEntryByTableName(string TableName)
		{
			using (var context = new Context())
			{
				return context.DBDicts.Where(d => d.TableName.Equals(TableName, StringComparison.InvariantCultureIgnoreCase)).ToList();
			}
		}

		public DataDictionaryEntity GetEntryByTableColumnName(string TableName, string ColumnName)
		{
			using (var context = new Context())
			{
				return context.DBDicts.Where(
					d => d.TableName.Equals(TableName, StringComparison.InvariantCultureIgnoreCase) 
					&& 
					d.ColumnName.Equals(ColumnName, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
			}
		}
    }
}
