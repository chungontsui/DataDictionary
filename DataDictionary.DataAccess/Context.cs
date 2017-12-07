using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataDictionary.DataAccess
{
	public class Context : DbContext
	{
		public Context() :base("DataDictConn")
		{}

		public DbSet<DataDictionaryEntity> DBDicts { get; set; }
	}

	public class DataDictionaryEntity
	{
		[Key]
		public int Id { get; set; }
		[Required]
		public string TableName { get; set; }
		[Required]
		public string ColumnName { get; set; }
		public string Decription { get; set; }
		public string PrxNo { get; set; }
	}
}
