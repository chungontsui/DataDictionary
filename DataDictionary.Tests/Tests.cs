using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using DataDictionary.DataAccess;

namespace DataDictionary.Tests
{
	public class Tests
	{
		[Test]
		public void TestAddEntry()
		{
			Repository repository = new Repository();

			string tName = "TestTable", cName = "TestColumn";

			DataDictionaryEntity newEntry = new DataDictionaryEntity()
			{
				TableName = tName,
				ColumnName = cName,
				Decription = "Testing Testing",
				PrxNo = "Testing 12345"
			};

			repository.AddEntry(newEntry);

			var result = repository.GetEntryByTableColumnName(tName, cName);

			Assert.That(result != null, "Testing DataDict didn't get Inserted");

		}
	}
}
