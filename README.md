Open Xml Imports
=================================================

OpenXmlImports is a robust library that simplifies importing and exporting data to and from OpenXml spreadsheet documents.

To export this class:

	public class PersonWorkbook
	{
		public List<Person> People { get; set; }
	}

	public class Person
	{
		public string Name { get; set; }
		public DateTime DateOfBirth { get; set; }
	}

You can write the following:

	WorkbookConfiguration<PersonWorkbook> config = OpenXmlConfiguration.Workbook<PersonWorkbook>().Create();

The above creates a configuration which you can then use to import/export data from OpenXml spreadsheets.

To export data:

	var data = new PersonWorkbook
	{
		People = new List<Person>()
		{
			new Person{ Name = "Chris", DateOfBirth = new DateTime(1981, 8, 28) },
			new Person { Name = "Fred", DateOfBirth = new DateTime(1984, 3, 7) }
		}
	};

	using(var output = File.OpenWrite("export.xlsx"))
		config.Export(data, output);


By default, this will generate an OpenXml spreadsheet document which contains one worksheet named "People". The People worksheet will contain two columns, one named "Name" and one named "Date Of Birth". It is possible to customize the appearance, formatting, and names of sheets/columns by using the included DSL.

For an import:


	using(var input = File.OpenRead("import.xlsx"))
		var personWorkbook = (PersonWorkbook)config.Import(data, input);

This will import the data from the file import.xlsx and build a PersonWorkbook object from it.


Stylesheets
--------------------------------------------------
Stylesheets are also used to format text, such as bold, italics, font size, and to format numbers as dates. You can specify an non-default stylesheet when the configuration is created:

	var workbook = OpenXmlConfiguration.Create<PersonWorkbook, MyStylesheet>();

By default, dates are formatted using the DefaultStylesheetProvider.DateFormat format. This provider also declares DateTimeFormat and TimeFormat formats for dates/times. You can create your own custom IStylesheetProvider if you need more options.

Possible Future Features
--------------------------------------------------
1. Figure out how to force column/sheet orders.
	- Probably needs to be a partial ordering
		- columns set to have the same order can appear in any order
2. More strongly type importer/exporter objects
	- Prevents the wrong object from being specified as the source
	- Eliminates casting on imports
3. Add styles for accounting/money formats, etc...
4. Support for identifying objects
	- Would allow for cross-sheet references w/ out duplicating objects.
5. Support for composite types
	- Ex:

			class LimitSet
			{
				public int Occurrence { get; set; }
				public int Aggregate { get; set; }
			}

6. 	Add support for more Boolean types
	- maybe a custom one that allows something like
			new MyBooleanType(string[] trueValues, string[] falseValues) or something
	- probably a few kinds of these, Y/N, yes/no, true/false
		- These could be implemented through the one in the first bullet
7. Build up x:cols in the worksheets so that columns can be best fit automatically on exports.