# Paragraph
Represents a single paragraph in a selection, range or document. Its a member of the paragraphs collection. The paragraphs collection includes all the paragrpahs ina selection range or document. The Paragraph object is a member of the Paragraphs collection.

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`parentContentControl`|  [ContentControl](contentControl.md)   |Returns the content control wrapping the object, if any. | Returns null if no content control|
|`font`|  [Font](font.md) | Entry point for formatting content.|  Exposes font name, size, color, and other properties. |
|`style`| String |Name of the style been used. | This is the name of an pre-installed or custom style.|




## Relationships
The Worksheet resource has the following relationships defined:

| Relationship     | Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`contentControls`](#contentcontrols)| [ContentControls](contentControls.md) collection |Collection of [contentControl](#contentcontrol.md) objects  in the current document | Includes content controls on the headers/footer and in the body of the document.  | 


## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`clearContent()`](#clearcontent)| Void | Clears the content of the calling object. | Undo operation by the user is supported. | 
|[`deleteElement(paragraphText: string, insertLocation: string)`](#insertparagraph)| [Paragraph](paragraph.md)  |Inserts a paragraph on the specified location. |All locations may not apply. See method details. | 
|[`getText()`](#gettext)| String |Gets the plain text of the calling object. | | 
|[`getHtml()`](#gethtml)| String  | Gets the HTML representation  of the calling object. | | 
|[`getOoxml()`](#getooxml)| String  | Gets the Office Open XML (OOXML) representation  of the calling object. |  | 
|[`insertContentControl()`](#insertcontentcontrol)| [ContentControl](contentcontrol.md)  |Wraps the calling object with a Rich Text content control. |  | 
|[`insertFile(fileLocation:string, location:string)`](#insertfile)| String |Inserts the complete specified document into the specified location. | | 
|[`insertBreak(paragraphText: string, insertLocation: string)`](#insertBreak)| [Paragraph](paragraph.md)  |Inserts a paragraph on the specified location. |All locations may not apply. See method details. | 
|[`insertParagraph(paragraphText: string, insertLocation: string)`](#insertparagraph)| [Paragraph](paragraph.md)  |Inserts a paragraph on the specified location. |All locations may not apply. See method details. | 
|[`insertPictureBase64(url: string, insertLocation: string)`](#insertPictureBase64)| [Paragraph](paragraph.md)  |Inserts a paragraph on the specified location. |All locations may not apply. See method details. | 
|[`insertPictureUrl(base64: string, insertLocation: string)`](#insertPictureUrl)| [Paragraph](paragraph.md)  |Inserts a paragraph on the specified location. |All locations may not apply. See method details.| \
|[`insertText(text: string, insertLocation: string)`](#inserttext)| [Range](range.md) | Inserts the specified text on the specified location. | All locations may not apply. See method details. | 
|[`insertHtml(html: string, insertLocation: string)`](#inserthtml)| [Range](range.md)  |Inserts the specified html on the specified location. | All locations may not apply. See method details.| 
|[`insertOoxml(ooxml: string, insertLocation: string)`](#insertooxml)| [Range](range.md)  |Inserts the specified ooxml on the specified location.  | All locations may not apply.See method details.| 
|[`search(text: string)`](#search)| [Ranges](ranges.md) |Executes a search on the scope of the calling object | Search results are a ranges collection. | 
|[`select(paragraphText: string, insertLocation: string)`](#select)| [Paragraph](paragraph.md)  | Selects and Navigates to the paragraph ||
|Paragraph Properties|
|[`getAlignment()`](#getAlignment)| float  |  Returns or sets an Alignment constant that represents the alignment for the specified paragraphs.     |||
|[`setAlignment(points: float)`](#insertparagraph)| void  |  Sets an Alignment constant that represents the alignment for the specified paragraphs.     |||
|[`getFirstLineIndent()`](#getFirstLineIndent)| float  |Returns the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write  | |
|[`setFirstLineIndent(points: float)`](#setFirstLineIndent)| void |Sets the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.   | |
|[`getLeftIndent()`](#getLeftIndent)| float  |Returns a Single that represents the left indent value (in points) for the specified paragraph.  || 
|[`setLeftIndent(points: float)`](#setLeftIndent)| void  |Sets a Single that represents the left indent value (in points) for the specified paragraph. || 
|[`getLineUnitBefore()`](#getLineUnitBefore)| float  |  Returns the amount of spacing (in gridlines) before the specified paragraph.    ||
|[`setLineUnitBefore(points: float)`](#setLineUnitBefore)| void  | Sets the amount of spacing (in gridlines) before the specified paragraph.    |||
|[`getLineSpacing()`](#insertparagraph)| float  |Returns or sets the line spacing (in points) for the specified paragraphs.     |||
|[`setLineSpacing(points: float)`](#insertparagraph)| void |  Returns or sets the line spacing (in points) for the specified paragraphs.    |||
|[`getOutlineLevel()`](#insertparagraph)|float   |  Returns the outline level for the specified paragraph.    |||
|[`setOutlineLevel(points: float)`](#insertparagraph)| void  |  Sets the outline level for the specified paragraph.    ||| 
|[`getRightIndent()`](#getRightIndent)| float  |  Returns a Single that represents the right indent value (in points) for the specified paragraph.    |||
|[`setRightIndent(points: float)`](#setRightIndent)|void  |   Sets a Single that represents the right indent value (in points) for the specified paragraph.   |||
|[`getSpaceAfter()`](#insertparagraph)|float  |    Returns the spacing (in points) before the specified paragraphs.   |||
|[`setSpaceAfter(points: float)`](#insertparagraph)|void  |    Sets the spacing (in points) before the specified paragraphs.   ||| 
|[`getSpaceBefore()`](#insertparagraph)| float   |  Returns the spacing (in points) before the specified paragraphs.    |||
|[`setSpaceBefore(points: float)`](#insertparagraph)| void  | Sets the spacing (in points) before the specified paragraphs.    |||


      
  


### Charts 

Get Charts collection that contains each of the chart objects contained in the worksheet. Each item contains the following properties. 

#### Syntax
```js
worksheetObject.charts;
```

#### Returns

[Chart](resources/chart.md) collection.

#### Examples

```js
var wSheetName = 'Sheet1';
var ctx = new Excel.ExcelClientContext();
var charts = ctx.workbook.worksheets.getItem(wSheetName).charts;
ctx.load(charts);
ctx.executeAsync().then(function () {
	for (var i = 0; i < charts.items.length; i++)
	{
		Console.log(charts.items[i].name);
	}
});
```
[Back](#relationships)



### getCell

Get the Cell (as a Range object) object based on row and column address relative to a top of worksheet. 

#### Syntax

```js
worksheetObject.getCell(row, column);
```

#### Parameters 

Parameter      | Type   | Description
-------------- | ------ | ------------
`row`          | Number | Required. Row number of the cell to be retrieved. Zero indexed. 
`col`          | Number | Required. Column number of the cell to be retrieved. Zero indexed.

#### Returns

[Range](resources/range.md) object.

#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "D5:F8";
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var cell = worksheet.cell(0,0);
ctx.load(cell);
ctx.executeAsync().then(function() {
	Console.log(cell.address);
});
```
[Back](#methods)


### getUsedRange

Get the used-range of a worksheet. 

#### Syntax
```js
worksheetObject.getUsedRange();
```
#### Parameters

None

#### Returns

[Range](resources/r.md) object.


#### Examples

```js
var ctx = new Excel.ExcelClientContext();
var wSheetName = 'Sheet1';
var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
var usedRange = worksheet.getUsedRange();
ctx.load(usedRange);
ctx.executeAsync().then(function () {
		Console.log(usedRange.address);
});
```
[Back](#methods)