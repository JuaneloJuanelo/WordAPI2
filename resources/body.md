# Body 
 Represents the body of document or section. If used in a document context represents the entire document body. If used in a section context is limited to the section boundaries.

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`parentContentControl`|  [ContentControl](contentControl.md)   |Represents the body of the document, not includes the header/footer and other section metadata. | |
|`font`|  [Font](font.md) | Entry point for formatting content.| The scope is the entire document. |
|`style`| String |Name of the style been used. | The scope is the entire document|




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
|[`clearContent()`](#clearcontent)| Void | Clears the content of the calling object. |  | 
|[`getText()`](#gettext)| String |Gets the plain text of the calling object. | | 
|[`getHtml()`](#gethtml)| String  | Gets the HTML  of the calling object. || 
|[`getOoxml()`](#getooxml)| String  | Gets the OOXML  of the calling object. |  | 
|[`insertText(text: string, insertLocation: string)`](#inserttext)| [Range](range.md) | Inserts the specified text on the specified location. | Returns the inserted text as Range | 
|[`InsertHtml(html: string, insertLocation: string)`](#inserthtml)| [Range](range.md)  |Inserts the specified html on the specified location. | | 
|[`InsertOoxml(ooxml: string, insertLocation: string)`](#insertooxml)| [Range](range.md)  |Inserts the specified ooxml on the specified location.  | | 
|[`InsertParagraph(paragraphText: string, insertLocation: string)`](#insertparagraph)| [Paragraph](paragraph.md)  |Inserts a paragraph on the specified location. | | 
|[`InsertContentControl()`](#insertcontentcontrol)| [ContentControl](contentcontrol.md)  |Wraps the calling object with a Rich Text content control. |  | 
|[`search(text: string)`](#search)| [Ranges](ranges.md) |Executes a search on the scope of the calling object | | 



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


### Sections 

Contains each of the section objects composing the document.

#### Syntax
```js
  document.sections

```

#### Returns

[Section](section.md) collection.

#### Examples

```js
// gets the paragprahs of the first section in the document. 
var ctx = new Word.WordClientContext();
ctx.customData = OfficeExtension.Constants.iterativeExecutor;

var paras = Ctx.document.sections.getItemAt(0).body.paragraphs;
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
                    console.log("paras[" + i + "].length = " + results[i].value.length);
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
