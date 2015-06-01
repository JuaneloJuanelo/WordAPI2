# Body 
 Represents the body of document or section. If used in a document context represents the entire document body. If used in a section context is limited to the section boundaries.

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
|[`paragraphs`](#paragraphs)| [Paragraphs](paragraphs.md) collection |Collection of [paragraph](#paragraph.md) objects within the body. |  |      
|[`inlinePictures`](#inlinepictures)| [InlinePictures](inlinepictures.md) collection |Collection of [inlinePicture](#picture.md) objects within the body. |Does not include floating images.  |       


## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`clearContent()`](#clearcontent)| Void | Clears the content of the calling object. | Undo operation by the user is supported. | 
|[`getText()`](#gettext)| String |Gets the plain text of the calling object. | | 
|[`getHtml()`](#gethtml)| String  | Gets the HTML representation  of the calling object. | | 
|[`getOoxml()`](#getooxml)| String  | Gets the Office Open XML (OOXML) representation  of the calling object. |  | 
|[`insertText(text: string, insertLocation: string)`](#inserttext)| [Range](range.md) | Inserts the specified text on the specified location. | All locations may not apply. See method details. | 
|[`insertHtml(html: string, insertLocation: string)`](#inserthtml)| [Range](range.md)  |Inserts the specified html on the specified location. | All locations may not apply. See method details.| 
|[`insertOoxml(ooxml: string, insertLocation: string)`](#insertooxml)| [Range](range.md)  |Inserts the specified ooxml on the specified location.  | All locations may not apply.See method details.| 
|[`insertParagraph(paragraphText: string, insertLocation: string)`](#insertparagraph)| [Paragraph](paragraph.md)  |Inserts a paragraph on the specified location. |All locations may not apply. See method details. | 
|[`insertContentControl()`](#insertcontentcontrol)| [ContentControl](contentcontrol.md)  |Wraps the calling object with a Rich Text content control. |  | 
|[`search(text: string)`](#search)| [Ranges](ranges.md) |Executes a search on the scope of the calling object | Search results are a ranges collection. | 



### ContentControls 

The colection holds all the content controls in the document.

#### Syntax
```js
  document.contentControls

```

#### Returns

[ContentControls](contentControls.md) collection.

#### Examples

```js
// enumerates all the content controls in the document
var ctx = new Word.WordClientContext();
var cCtrls = ctx.document.body.contentControls;
ctx.load(cCtrls);

ctx.executeAsync().then(
	function () {
		var results = new Array();
		for (var i = 0; i < cCtrls.count; i++) {
			results.push(cCtrls.getItemAt(i));
		}
		ctx.executeAsync().then(
			function () {
				for (var i = 0; i < results.length; i++) {
					console.log("contentControl[" + i + "].length = " + results[i].text.length);
				}
			}
		);
	},
	function (result) {
		console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
		console.log(result.traceMessages);
	}
);


```
[Back](#relationships)


### Paragraphs 

The colection holds all the ccontent controls in the scope.

#### Syntax
```js
  document.body.paragraphs  // returns the paragraphs on the body of the document.
  document.sections.getItemAt(0).paragraphs  //returns the paragraphs in the first section of the document.
  document.selection.paragraphs   //returns the paragraphs contained in the selection.

```

#### Returns

[Paragraphs](paragraphs.md) collection.

#### Examples

```js

// this example iterates all the paragraphs in the documents and reports back the lenght and text of each paragraph in the document

var ctx = new Word.WordClientContext();
ctx.customData = OfficeExtension.Constants.iterativeExecutor;

var paras = ctx.document.body.paragraphs;
ctx.load(paras);

ctx.executeAsync().then(
    function () {
        var results = new Array();
        for (var i = 0; i < paras.count; i++) {
            results.push(paras.getItemAt(i).getPlainText());
        }
        ctx.executeAsync().then(
            function () {
                for (var i = 0; i < results.length; i++) {
                    console.log("paras[" + i + "].length = " + results[i].value.length + " " + results[i].value);
                }
            }
        );
    },
    function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
    }
);

```
[Back](#relationships)


### InlinePictures 

The colection holds all the ccontent controls in the scope.

#### Syntax
```js
  document.body.paragraphs  // returns the paragraphs on the body of the document.
  document.sections.getItemAt(0).paragraphs  //returns the paragraphs in the first section of the document.
  document.selection.paragraphs   //returns the paragraphs contained in the selection.

```

#### Returns

[Paragraphs](paragraphs.md) collection.

#### Examples

```js

// this example iterates all the paragraphs in the documents and reports back the lenght and text of each paragraph in the document

var ctx = new Word.WordClientContext();
ctx.customData = OfficeExtension.Constants.iterativeExecutor;

var paras = ctx.document.body.paragraphs;
ctx.load(paras);

ctx.executeAsync().then(
    function () {
        var results = new Array();
        for (var i = 0; i < paras.count; i++) {
            results.push(paras.getItemAt(i).getPlainText());
        }
        ctx.executeAsync().then(
            function () {
                for (var i = 0; i < results.length; i++) {
                    console.log("paras[" + i + "].length = " + results[i].value.length + " " + results[i].value);
                }
            }
        );
    },
    function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
    }
);

```
[Back](#relationships)

### Methods 

#### Examples

### save

Saves the current document. 

#### Syntax

```js

ctx.document.save();
```

#### Parameters 

None

#### Returns

Void

#### Examples

```js
var ctx = new Word.WordClientContext();
ctx.document.save();
```
[Back](#methods)
