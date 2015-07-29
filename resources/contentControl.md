# ContentControl

Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as dates, lists, or paragraphs of formatted text. Currently, only rich text content controls are supported. 


## Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|appearance|  string |Gets or sets the appearance of the content control. The value can be 'boundingBox', 'tags' or 'hidden'. |
|cannotDelete|  bool |Gets or sets a value that indicates whether the user can delete a content control from the active document.|
|cannotEdit|  bool | Gets or sets a value that indicates whether the user can edit the contents of a content control.|
|color|  string |   Gets or sets the color of the content control. Color is set in "#FFFFFF" format or by using the color name.|
|font|  [Font](font.md) | Gets the text format of the content control. Use this to get and set font name, size, color, and other properties. |
|id|  string |Gets a string that represents the content control identifier. |
|parentContentControl|  [ContentControl](contentControl.md)   |Gets the content control that contains the content control. Returns null if there isn't a parent content control.|
|placeholderText|  string   | Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the  content control is empty.|
|removeWhenEdited|  bool | Gets or sets a value that indicates whether the content control is removed after it is edited.|
|title|  string  |  Gets or sets the title for a content control.   | 
|text|  string  |  Gets or sets the text of the content control. |
|type|  string  | Gets or sets the content control type. Only rich text content controls are supported|
|style| string |Gets or sets the style used for the content control. This is the name of the pre-installed or custom style.|
|tag| string |Gets or sets a value to identify a content control. |



## Relationships

| Relationship     | Type    |Description|
|:-----------------|:--------|:----------|
|contentControls | [contentControlCollection](contentControlCollection.md)  | Gets the collection of content control objects in the current content control. | 
|inlinePictures | [inlinePictureCollection](inlinePictureCollection.md)  | Gets the collection of inlinePicture objects in the current content control. The collection does not include floating images.  | 
|paragraphs| [paragraphCollection](paragraphCollection.md)  | Get the collection of paragraph objects in the content control. |      

       

## Methods


| Method     | Return Type    |Description|
|:-----------------|:--------|:----------|
|[clear()](#clear)| void | Clears the contents of the content control. The user can perform the undo operation on the cleared content. |
|[delete(keepContent: bool)](#deletekeepcontent-bool)| void  | Deletes the content control and its content from the document. If keepContent is set to true, the content is not deleted. | 
|[getHtml()](#gethtml)| string  | Gets the HTML representation  of the content control object. | 
|[getOoxml()](#getooxml)| string  | Gets the Office Open XML (OOXML) representation  of the content control object. | 
|[insertFileFromBase64(base64File: string, insertLocation: string)](#insertfilefrombase64base64file-string-insertlocation-string)| string |Inserts a document into the current content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.| 
|[insertBreak(breakType: string, insertLocation: string)](#insertbreakbreaktype-string-insertlocation-string)| void | Inserts a break at the specified location. The insertLocation value can be 'Start' or 'End'. | 
|[insertParagraph(paragraphText: string, insertLocation: string)](#insertparagraphparagraphtext-string-insertlocation-string)| [Paragraph](paragraph.md)  |Inserts a paragraph at the specified location. The insertLocation value can be 'Start' or 'End'. | 
|[insertText(text: string, insertLocation: string)](#inserttexttext-string-insertlocation-string)| [Range](range.md) | Inserts text into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'. | 
|[insertHtml(html: string, insertLocation: string)](#inserthtmlhtml-string-insertlocation-string)| [Range](range.md)  |Inserts HTML into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'. | 
|[insertOoxml(ooxml: string, insertLocation: string)](#insertooxmlooxml-string-insertlocation-string)| [Range](range.md)  |Inserts OOXML into the content control at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'. | 
|[load(param: object)](#loadparam-object)|void|Fills the content control proxy object created in the JavaScript layer with property and object values specified in the parameter.|

|[search(searchText : string, searchOptions: searchOptions)](#searchsearchtext-string-searchoptions-searchoptions)| [searchResultCollection](searchResultCollection.md) |Performs a search with the specified searchOptions on the scope of the content control object. The search results are a collection of range objects. | 
|[select()](#select)| void |Selects the content control. This causes Word to scroll to the selection.  | 
  
## API Specification


### clear

Clears the content of the calling object.

#### Syntax
```js
ctx.document.body.clearContent();

```
#### Parameters

None

#### Returns

void.


#### Examples

```js

//Clear content of the body of the document...

var ctx = new Word.RequestContext();

ctx.document.body.clear();
ctx.executeAsync().then(
   function () {
     console.log("Success!!");
   },
   function (result) {
     console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
     console.log(result.traceMessages);
   }
);
```
[Back](#methods)


### getHtml()

Gets the HTML representation  of the calling object.

#### Syntax
```js
var myTHTML  = document.body.Html();
```
#### Parameters

None

#### Returns

[Range](range.md).


#### Examples

```js
var myHTML  = document.body.getHtml();
```
[Back](#methods)

### getOoxml

Gets the Office Open XML (OOXML) representation  of the calling object.

#### Syntax
```js
var myOOXML  = document.body.getOoxml();
```
#### Parameters

None

#### Returns

[Range](range.md).


#### Examples

```js
var myOOXML  = document.body.getOoxml();
```
[Back](#methods)

### insertText()

Inserts the specified text on the specified location.

#### Syntax
```js
var myText = document.body.insertText("Hello World!", "End");
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`text`          | string | Required. Text to be inserted.
`insertLocation`          | string | Either "Start" "End"  the body of the document.

#### Returns

[Range](range.md).


#### Examples

```js

//inserts some text at the end of the document.
var ctx = new Word.RequestContext();
ctx.document.body.insertText("new text", "end");
ctx.executeAsync().then(
    function () {
    console.log("Success!!");    
    },
    function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
    }
);
```
[Back](#methods)

### insertHtml()

Inserts the specified HTML on the specified location.

#### Syntax
```js
var myRange = document.body.insertHtml("<b>This is some bold text</b>", "End");
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`html`          | string | Required. the HTML to be inserted in the document.
`insertLocation`          | string | Either "Start" "End"  the body of the document

#### Returns

[Range](range.md) .


#### Examples

```js
//inserts some html at the end of the doc :) 
var ctx = new Word.RequestContext();
ctx.document.body.insertHtml("<b>This is some bold text</b>", "End");
ctx.executeAsync().then(
    function () {
    console.log("Success!!");    
    },
    function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
    }
);
```
[Back](#methods)

### insertOoxml()

Inserts the specified OOXML on the specified location.

#### Syntax
```js
range.insertOoxml(ooxmlText, Word.InsertLocation.end);

```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`ooxml`          | string | Required. OOXML to be inserted.
`insertLocation`          | string | Either "Start" "End"  the body of the document
 
#### Returns

[Range](range.md) collection.


#### Examples

```js
// this code inserts some formatted text into the document!
var ctx = new Word.RequestContext();
var range = ctx.document.getSelection();

var ooxmlText =
  "<w:p xmlns:w='http://schemas.microsoft.com/office/word/2003/wordml'><w:r><w:rPr><w:b/><w:b-cs/><w:color w:val='FF0000'/><w:sz w:val='28'/><w:sz-cs w:val='28'/></w:rPr><w:t>Hello world (this should be bold, red, size 14).</w:t></w:r></w:p>";

range.insertOoxml(ooxmlText, Word.InsertLocation.end);

ctx.executeAsync().then(
   function () {
     console.log("Success");
   },
   function (result) {
     console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
     console.log(result.traceMessages);
   }
);

  ```
[Back](#methods)

### insertParagraph()

Inserts a paragraph on the specified location.

#### Syntax
```js
var ccs = document.insertParagraph("Some initial text", "Start");
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`paragraphText`          | string | Paragrph text. null for blank Paragraph.
`insertLocation`          | string | Either "Start" "End"  the body of the document


#### Returns

[Paragraph](Paragraph.md).


#### Examples

```js
//Inserting paragraphs at the end of the document.

var ctx = new Word.RequestContext();

var myPar = ctx.document.body.insertParagraph("Bibliography","end");
myPar.style = "Heading 1";

var myPar2 = ctx.document.body.insertParagraph("this is my first book","end");
myPar2.style = "Normal"



ctx.executeAsync().then(
     function () {
         console.log("Success!!");
     },
     function (result) {
         console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        // console.log(result.traceMessages);
     }
);
```
[Back](#methods)

### insertContentControl()

Wraps the calling object with a Rich Text content control.

#### Syntax
```js
var ccs = document.body.insertContentControl();
```
#### Parameters

None

#### Returns

[ContentControl](contentControl.md).


#### Examples

```js
//Insert a Content Control (on user's selection)  and changing the properties by using the selection

var ctx = new Word.RequestContext();
var range = ctx.document.getSelection();

var myContentControl = range.insertContentControl();
myContentControl.tag = "Customer-Address";
myContentControl.title = "Enter Customer Address Here:";
myContentControl.style = "Heading 1";
myContentControl.insertText("One Microsoft Way,Redmond,WA,98052",'replace');
myContentControl.cannotEdit = true;
myContentControl.appearance = "tags";



ctx.load(myContentControl);

ctx.executeAsync().then(
   function () {
     console.log("Content control Id: " + myContentControl.id);
   },
   function (result) {
     console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
     console.log(result.traceMessages);
   }
);



```
[Back](#methods)

### search

Executes a search on the scope of the calling object.

#### Syntax
```js
var searchResults = document.body.search("Sales Report");
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`text`          | string | Required. Text to be searched.

#### Returns

[Ranges](searchResultCollection.md) collection.


#### Examples

```js
///Search example! 

var ctx = new Word.RequestContext();
var options = Word.SearchOptions.newObject(ctx);

options.matchCase = false

var results = ctx.document.body.search("Video", options);
ctx.load(results, {select:"text, font/color", expand:"font"});
ctx.references.add(results);

ctx.executeAsync().then(
  function () {
    console.log("Found count: " + results.items.length + " " + results.items[0].font.color );
    for (var i = 0; i < results.items.length; i++) {
      results.items[i].font.color = "#FF0000"    // Change color to Red
      results.items[i].font.highlightColor = "#FFFF00";
      results.items[i].font.bold = true;
      if (i == 3)
        results.items[i].select();
    }
    ctx.references.remove(results);
    ctx.executeAsync().then(
      function () {
        console.log("Deleted");
      }
    );
  }
);
```
[Back](#methods)


### insertFile()

Inserts the specified file on the specified location.

#### Syntax
```js
TBD
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`fileLocation`          | string | Required. Full path to the file to be inserted. Can be on the hard drive, or a url.
`insertLocation`          | string | Either "Start" "End"  the body of the document.


#### Returns

[Range](range.md) collection.


#### Examples

```js
TBD

```
[Back](#methods)

### insertBreak()

Inserts the specified [type of break](breakType.md) on the specified location.

#### Syntax
```js
ctx.document.body.insertBreak("page", "End");
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`breakType`          | string | Required.  [Type of break](breakType.md)
`insertLocation`          | string | Either "Start" "End"  the body of the document.


#### Returns

[Range](range.md) collection.


#### Examples

```js
//inserts a page break and then adds a paragraph!

var ctx = new Word.RequestContext();

ctx.document.body.insertBreak("page", "End");
ctx.document.body.insertParagraph("Hello after break!","End");

ctx.executeAsync().then(
  function () {
    console.log("Success");
  },
  function (result) {
    console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
    console.log(result.traceMessages);
  }
);


```
[Back](#methods)
