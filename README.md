Open Xml Imports
=================================================

OpenXmlImports is a robust library that simplifies importing and exporting data to and from Excel spreadsheets.

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

	OpenXmlConfiguration config = OpenXmlConfiguration.Workbook<PersonWorkbook>().Create();

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

	using(var fs = File.OpenWrite("export.xlsx"))
		config.Export(data, fs);


By default, this will generate an OpenXml spreadsheet document which contains one worksheet named "People". The People worksheet will contain two columns, one named "Name" and one named "Date Of Birth". It is possible to customize the appearance, formatting, and names of sheets/columns by using the included DSL.

For an import:


	using(var fs = File.OpenWrite("import.xlsx"))
		var personWorkbook = (PersonWorkbook)config.Import(data, fs);

This will import the data from the file import.xlsx and build a PersonWorkbook object from it.


Stylesheets
--------------------------------------------------
OpenXml stores dates/times as OLE Automation (OA) dates, which means they are stored as numbers. Special formatting is applied to the number which causes it to appear as a date/time.

By default, dates are formatted using the DefaultStylesheetProvider.DateFormat format. This provider also declares DateTimeFormat and TimeFormat formats for dates/times. You can create your own custom IStylesheetProvider if you need more options.

Stylesheets are also used to format text, such as bold, italics, and font size. Support for this is not provided by default (yet).

Possible Future Features
--------------------------------------------------
1. Need good style sheet support. Necessary to format dates correctly.
	- This is implemented, but doesn't feel great yet.
	- Would be nice if styling were more strongly typed in the configuration DSL.
	- Add support for bold/italics/underline and all permutations out of the box.
2. Support for sheets which are not collection bound (No IList<>).
3. Need to support shared strings correctly.
	- Mostly important on the import side. Exporting is handled by open xml SDK.
	- Use the shared string table during exports.
4. Support for non-nullable/required members
5. Support for composite types
	1. Ex:

			class LimitSet
			{
				public int Occurrence { get; set; }
				public int Aggregate { get; set; }
			}
	2. This is an enhancement to implementing NH-style ITypes.
6. Support for identifying objects
	- Would allow for cross-sheet references w/ out creating multiple objects
7. Ability to ignore members.
	- WTH was I thinking w/ this one? If data is from DTOs, all eligible members should be considered. I don't think it is a good idea to allow direct importation to domain entities, but I could change my mind here.
8. Support for max length on string columns.
	- Will probably be in DSL for all columns, just only used by string columns.
9. Support ignoring rows with no cells.
	- Could maybe just assume this is always ok.
10. User-friendly aliases for type names. This is probably part of a larger effort to support NH-like ITypes for coercion to/from OpenXml formats.
	- May also need special support for decimals since excel may store them using scientific notation.