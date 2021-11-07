Manage translations
===================

A spreadsheet and macro for app developers to manage translations more easily.
Keep track of the list of captions and label texts, have translators or your users
translate the texts, and finally use a macro to generate the source files used in your project.

It contains a VBA and LibreOffice script that outputs the .json, .xml, .strings or .resx files
which can be used in JavaScript, Eclipse, XCode or Visual Studio.
Includes examples for both MS-Excel and LibreOffice Calc.

![preview screenshot](/translations_preview.png?raw=true "Translations spreadsheet preview")

How to use
----------
To use this sheet for your project, first add all labels you need for your app 
in column A. To keep the sheet structured, you can use empty rows or comment 
rows starting with double slash (//). You can also change or add columns to add 
more translations.

Hire a translator or have your superusers make the necessary translations.
When all the translations are ready, you can generate the JSON, Eclipse XML files, the
XCode .strings files or the Visual Studio resource .resx files. Press ALT+F8 and run one
of the following macro's:

	GenerateLocalisationJson
	GenerateLocalisationEclipse
	GenerateLocalisationVisualStudio
	GenerateLocalisationXcode

The macros will create a folder called "json" or "xcode" or "eclipse" or
"visualstudio" in the same folder where the spreadsheet file is located.
Visual Studio .resx output is untested, the files might need some tweaking.

Sheet content
-------------
Column A contains the keys you will use in your code. These are the keys for
NSLocalizedString in XCode, or the string ID's in Eclipse. Column B and further
contain the actual translations which you can display somewhere in your app.

The first 4 rows contain information about each translation. Language code is a
ISO 639-1 code, a two letter code for the language. For example es for Spanish,
see full [ISO 639-1 list](http://en.wikipedia.org/wiki/List_of_ISO_639-1_codes)

The displayname and translator cells only serve for comment/credits purposes.
The cell colors are for display only, the macro doesn't use them.


Export preview
--------------
The macro can export to different formats, here is a preview of the output.

JavaScript `JSON` files

	{
		"en": {
			"Start": "Start game",
			"Editor": "Editor",
			..etc

XCode `Localization.strings` files

	// menu items
	"Start" = "Start game";
	"Editor" = "Editor";
	..etc

Eclipse `string.xml` files

	<?xml version="1.0" encoding="utf-8"?>
	<resources>
		<!-- menu items -->
		<string name="start">Start game</string>
		<string name="editor">Editor</string>
		..etc

Visual Studio `.resx` files export (untested)

	<?xml version="1.0" encoding="utf-8"?>
	<root>
		<!-- menu items -->
		<data name="start">
			<value>Start game</value>
		</data>
		<data name="editor">
			<value>Editor</value>
		</data>
		..etc
		
Questions
---------

* Why use a spreadsheet when there are translation services?

App localization services can work for large scale projects or when you're constantly adding content and new texts to be translated.
But for most apps it doesn't make a lot of sense to pay a monthly fee for a service that basically manages a list of phrases.

* Why not just use automatic tanslations like Google Translate?

Automatic translations can give pretty impressive results, but translation errors are still [very common](https://www.reddit.com/r/Translation_Fails/).
For text labels and captions in an app the probability of errors is even higher, because they're typically short texts without any context.
As a real world example of translations errors, in an industrial machinery app the English label `Running time installation` was automatically translated to Russian as basically `"Duration of the theater show"`,
and in a retail inventory report `Stock shift` was automatically translated to `"Stock shares moved"`.

* Why not just query an online translations api?

Besides the translation errors, this adds a lot of complexity;
Do you require the user to be always-online? Do you initially cache the translation results?
What if the translation server is temporarily offline?
This requires a disproportionate amount of effort for what is essentially a static list of phrases and words.


Questions, comments - Bas de Reuver bdr1976@gmail.com
