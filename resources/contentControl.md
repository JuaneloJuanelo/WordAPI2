# ContentControl

An individual content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as dates, lists, or paragraphs of formatted text. On this release, only rich text content controls are supported. The ContentControl object is a member of the ContentControls collection.


## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`appearance`|  String |Returns or sets the appearance of the content control. |RW. Can be 'boundingBox', 'tags' or 'hidden' |
|`cannotDelete`|  boolean |Returns or sets a Boolean that represents whether the user can delete a content control from the active document |RW. |
|`cannotEdit`|  boolean | Returns or sets a Boolean that represents whether the user can edit the contents of a content control. |RW. |
|`color`|  Number |   Returns or sets the color of the content control.        | Color is set in "#FFFFFF" format or color name|
|`font`|  [Font](font.md) | Entry point for formatting content.|  Exposes font name, size, color, and other properties. |
|`id`|  String |Returns a String that represents the identification for a content control. |Read-only|\
|`parentContentControl`|  [ContentControl](contentControl.md)   |Returns the content control wrapping the object, if any. | Returns null if no content control|
|`removeWhenEdited`|  boolean |  Removes the content control after edited.         ||
|`title`|  String  |  Returns or sets a String that represents the title for a content control.   | |
|`text`|  String  |  Returns or sets the text of the Content Control  | |
|`type`|  String  | Returns or sets  the type for a content control.          |Only rich text content controls are supported|\
|`style`| String |Name of the style been used. | This is the name of an pre-installed or custom style.|
|`tag`| String |Returns or sets a String that represents a value to identify a content control. | RW and might be duplicated|



## Relationships
The Content Control resource has the following relationships defined:

| Relationship     | Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`contentControls`](#contentcontrols)| [ContentControls](contentControlCollection.md) collection |Collection of [contentControl](contentControl.md) objects  in the current document | Includes content controls on the headers/footer and in the body of the document.  | 
|[`inlinePictures`](#inlinepictures)| [InlinePictures](inlinePictureCollection.md) collection |Collection of [inlinePicture](inlinePicture.md) objects within the body. |Does not include floating images.  | 
|[`paragraphs`](#paragraphs)| [Paragraphs](paragraphCollection.md) collection |Collection of [paragraph](paragraph.md) objects within the content control. |  |      

       

## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`clear()`](#clear)| Void | Clears the content of the calling object. | Undo operation by the user is supported. | 
|[`delete(keepContent:boolean )`](#deleteelement)| Void  |Deletes the content control and its content from the document, users may keep the content if send true as parameter. | | 
|[`getHtml()`](#gethtml)| String  | Gets the HTML representation  of the calling object. | IMPORTANT: we are deprecating this method in favor of the property| 
|[`getOoxml()`](#getooxml)| String  | Gets the Office Open XML (OOXML) representation  of the calling object. | IMPORTANT: we are deprecating this method in favor of the property | 
|[`insertContentControl()`](#insertcontentcontrol)| [ContentControl](contentcontrol.md)  |Wraps the calling object with a Rich Text content control. |  | 
|[`insertFile(fileLocation:String, insertLocation:String)`](#insertfile)| String |Inserts the complete specified document into the specified location. | | 
|[`insertBreak(breakType: String, insertLocation: String)`](#insertBreak)|void  | Inserts the specified [type of break](breakType.md) on the specified location. |All locations may not apply. See method details. | 
|[`insertParagraph(paragraphText: String, insertLocation: String)`](#insertparagraph)| [Paragraph](paragraph.md)  |Inserts a paragraph on the specified location. |All locations may not apply. See method details. | 
|[`insertPictureBase64(url: String, insertLocation: String)`](#insertPictureBase64)| [Paragraph](paragraph.md)  |Inserts a picture on the specified location. |All locations may not apply. See method details. | 
|[`insertText(text: String, insertLocation: String)`](#inserttext)| [Range](range.md) | Inserts the specified text on the specified location. | All locations may not apply. See method details. | 
|[`insertHtml(html: String, insertLocation: String)`](#inserthtml)| [Range](range.md)  |Inserts the specified html on the specified location. | All locations may not apply. See method details.| 
|[`insertOoxml(ooxml: String, insertLocation: String)`](#insertooxml)| [Range](range.md)  |Inserts the specified ooxml on the specified location.  | All locations may not apply.See method details.| 
|[`select(paragraphText: String, insertLocation: String)`](#select)| [Paragraph](paragraph.md)  | Selects and Navigates to the paragraph ||
  


### ContentControls 

The colection holds all the content controls in the document.

#### Syntax
```js
  document.contentControls

```

#### Returns

[ContentControls](contentControlCollection.md) collection. See ContentControl(contentControl.md) object.

#### Examples

```js
// enumerates all the content controls in the document
var ctx = new Word.RequestContext();
var cCtrls = ctx.document.body.contentControls;
ctx.load(cCtrls,{select:'appearance,text'});  // just need these properties!

ctx.executeAsync().then(
    function () {
        var results = new Array();
     
        for (var i = 0; i < cCtrls.items.length; i++) {
           console.log("contentControl[" + i + "].text = " + cCtrls.items[i].text + " Appearance:" +cCtrls.items[i].appearance );
      }
        ctx.executeAsync().then(
            function () {
               console.log("Success!!");
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

The colection holds all the paragraphs in the scope.

#### Syntax
```js
  document.body.paragraphs  // returns the paragraphs on the body of the document.
  document.sections.getItemAt(0).paragraphs  //returns the paragraphs in the first section of the document.
  document.selection.paragraphs   //returns the paragraphs contained in the selection.

```

#### Returns

[Paragraphs](paragraphCollection.md) collection. See [Paragraph](paragrph.md) object.

#### Examples

```js

// this example iterates all the paragraphs in the documents and reports back the text of each paragraph in the document
var ctx = new Word.RequestContext();
var paras = ctx.document.body.paragraphs;
ctx.load(paras,{select:"text"});

ctx.executeAsync().then(
  function () {
    for (var i = 0; i < paras.items.length; i++) {
      console.log("paras[" + i + "].content  = " + paras.items[i].text);
    }
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

### clear

Clears the content of the calling object.

#### Syntax
```js
ctx.document.body.clearContent();

```
#### Parameters

None

#### Returns

Void.


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
`text`          | String | Required. Text to be inserted.
`insertLocation`          | String | Either "Start" "End"  the body of the document.

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
`html`          | String | Required. the HTML to be inserted in the document.
`insertLocation`          | String | Either "Start" "End"  the body of the document

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
`ooxml`          | String | Required. OOXML to be inserted.
`insertLocation`          | String | Either "Start" "End"  the body of the document
 
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
`paragraphText`          | String | Paragrph text. null for blank Paragraph.
`insertLocation`          | String | Either "Start" "End"  the body of the document


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
`text`          | String | Required. Text to be searched.

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
`fileLocation`          | String | Required. Full path to the file to be inserted. Can be on the hard drive, or a url.
`insertLocation`          | String | Either "Start" "End"  the body of the document.


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
`breakType`          | String | Required.  [Type of break](breakType.md)
`insertLocation`          | String | Either "Start" "End"  the body of the document.


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