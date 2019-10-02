# localizable-sheet-script
A Google Sheets script that will take a sheet in a specific format and return iOS and Android localization files.

## What it does

For Android it creates an XML resources file with all of the strings. For iOS it creates a Localized enum with String enum, a .strings and a .stringsdict files.

## Installing

1. Open your sheet.
2. Go to **Tools -> Script Editor**
3. Copy `Code.js`, make your edits if needed, and **Save**.

## Usage

1. Open your sheet.
2. Go to **Custom Export** and select your **iOS** or **Android**.

## Sheet format

The script expects the sheet to be formatted in a specific way.

| ... arbitrary number of columns before iOS keys | **Identifier iOS** | **Identifier Android** | English text | German text | ... |
|-------------------------------------------------|--------------------|------------------------|--------------|-------------|-----|
| place whatever you want in here                 | login_button_title | login_button_title     | Login        | Einloggen   |     |
|                                                 | ...                | ...                    | ...          | ...         |     |

**The texts in bold cannot be changed!** The script depends on them to know which identifier is which. The other texts don't matter.

The **first row** must **always** contain headers, and not the actual strings.

The number of languages depends on the `NUMBER_OF_LANGUAGES` variable in the script, and new languages can be added by adding a new column on the right and incrementing that number in the script.

The position of the first (iOS) column that is relevant to the script is changed with the `FIRST_COLUMN_POSITION` variable in the script. By default it's `1` (the first column).

Please use dot separated identifier for iOS to get autogenerated structured enum. Apply **=SUBSTITUTE({Android_column}{Row_NUMBER}; "."; "_")** formula on Android identifier column as dots are not supported on it. It is being constructed based on the first language, so it has to have all the strings for the autocompletition to work across yours iOS project.

Formatted example:

| **Identifier iOS** | **Identifier Android** | English text |
| --- | --- | --- |
distance.format | distance_format| %1$s away |									

%s and %1$s are transformed to %@ and %1$@ for iOS by the export script

Plurals example:
| **Identifier iOS** | **Identifier Android** | English text |
| --- | --- | --- |
| missions.missions_format[p] | missions_missions_format[p] | You have %#@MISSIONS@ to do |
| d[]MISSIONS | d[]MISSIONS | 
| zero | zero | no missions
| one | one | %d mission
| two | two |	
| few | few |
| many | many |
| other | other | %d missions |

## Configuring

 - `NUMBER_OF_LANGUAGES`: The number of language columns to use.
 - `FIRST_COLUMN_POSITION`: The position of the iOS identifiers (the first column relevant to the script). Starting from 1.
 ~~- `IOS_INCLUDES_LOCALIZABLE_ENUM`: Whether or not to create an `Localizable` `enum` containing all of the keys as `static let` constants.~~
 
## Exported files

The exported files are the standard format (`strings.xml` or `Localizable.strings` + `Localizable.stringsdict`) for the specific platforms. iOS also includes a `Localized` `enum` which contains all of the keys as `case` statements for code-completion and less typos. Use dot separated format for the nested enums generation. The **Localizable** protocol:

```swift
protocol Localizable {
	var rawValue: String { get }
	func get(_: String?) -> String
}

protocol LocalizableFormat: Localizable { //detected by a "format" suffix in the identifier
	func getFormatted(arguments: CVarArg ...) -> String
}

extension Localizable {
	func get(_ comment: String? = nil) -> String {
		return "use NSLocalizedString or other custom logic like bundle?.localizedString(forKey: key, value: comment, table: nil) ?? key
	}
}

extension LocalizableFormat {
	func getFormatted(arguments: CVarArg...) -> String {
		let format = self.get(nil)
		return String(format: format, arguments: arguments)
	}
}
```
Usage:
```swift
Localized.сhange.get() //returns the localized string for the "change" identifier
Localized.Popup.Validate.Issues.end_of_world.get()    //popup.validate.issues.end_of_world

let formatTextStart = Localized.aScreen.Format_example._1.get("start")
let formatTextEnd = Localized.aScreen.Format_example._2.get("end")
Localized.aScreen.Format_example.format.getFormatted(arguments: formatTextStart, formatTextEnd)
```

## License: MIT

Created by COBE http://cobeisfresh.com/
Copyright 2017 COBE

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
