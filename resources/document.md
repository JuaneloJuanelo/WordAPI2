# Document 
 Represents a Word document. Main entry point to all interactions with the document. A document is composed of one or more sections(resources/section.md), and a body where the main content of the document resides.

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`body`|  `[Body'](body.md)   |Represents the body of the document, not includes the header/footer and other section metadata | |
|`saved`|  bool |Indicates if the document is dirty, and requires to be saved | |
|`selection`| [Range'](range.md) |Represents the continous current selection of the document. Since it can expand multiple paragraphs its considered to be a Range Object. 
If there is no selection, it represents the insertion point in the document.
 |Multiple selection is not supported|



## Relationships
The Worksheet resource has the following relationships defined:

| Relationship     | Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`sections`](#sections)| [Section](section.md) collection |Collection of sections in the current document |Document.Section  |       

## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`getContentControlById(id: string)`](#getContentControlById)| [ContentControl](contentControl.md) object |Returns the content control with the specified Id, returns null if the content control does not exist|  |
|[`getContentControlByName(name: string)`](#getContentControlByName)| [ContentControls](contentControls.md) collection |Returns the collection of the content controls matching the specified name| Since there could be many Content Controls with the same name, this method returns a collection|  
|[`getContentControlByTag(tag: string)`](#getContentControlByTag)| [ContentControls](contentControls.md) collection |Returns the collection of the content controls matching the specified tag| Since there could be many Content Controls with the same name, this method returns a collection |
|[`save(void)`](#save)| Void |Saves the Document | If document has not saved before it will use Word default names (i.e. Document1.docx, etc.) |     



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



### getContentControlById

Creates a content control, gets the Id, then retrieves the content control and changes it appearance. 

#### Syntax

```js
var ctx = new 
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