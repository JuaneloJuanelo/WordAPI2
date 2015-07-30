# Paragraph
Represents a single paragraph in a selection, range, document, or document body.

## Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|parentContentControl|  [ContentControl](contentControl.md) | Gets the content control that contains the paragraph. Returns null if there isn't a parent content control.|
|font|  [Font](font.md) | Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties. |
|alignment| string |Gets or sets the alignment for a paragraph. The value can  be "left", "centered", "right", or "justified". |
|firstLineIndent| number |Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.|
|leftIndent| number | Gets or sets the left indent value, in points, for the paragraph.|
|lineSpacing| number | Gets or sets the line spacing, (in points) for the specified paragraph. In the Word UI, this value is divided by 12. |
|lineUnitAfter| number |Gets or sets the amount of spacing, in grid lines. after the paragraph.|
|lineUnitBefore| number |Gets or sets the amount of spacing, in grid lines, before the paragraph.
|outlineLevel| number |Gets or sets the outline level for the paragraph.
|rightIndent| number |Gets or sets the right indent value, in points, for the paragraph.
|spaceAfter| number |Gets or sets the spacing, in points, after the paragraph. |
|spaceBefore| number |Gets or sets the spacing, in points, before the paragraph. |
|text|  string  |  Gets or sets the text of the paragraph.  | 


## Relationships

| Relationship     | Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|contentControls| [contentControlCollection](contentControlCollection.md)  | Gets the collection of content control objects that are in the paragraph.|
|inlinepictures| [inlinePictureCollection](inlinePictureCollection.md)  |Gets the collection of inlinePicture objects that are in the paragraph. The collection does not include floating images.  | 


## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`clear()`](#clear)| Void | Clears the content of the calling object. | Undo operation by the user is supported. | 
|[`delete()`](#delete)| Void  |Deletes the content control and its content from the document | | 
|[`getHtml()`](#gethtml)| string  | Gets the HTML representation  of the calling object. | | 
|[`getOoxml()`](#getooxml)| string  | Gets the Office Open XML (OOXML) representation  of the calling object. |  | 
|[`insertContentControl()`](#insertcontentcontrol)| [ContentControl](contentcontrol.md)  |Wraps the calling object with a Rich Text content control. |  | 
|[`insertFile(fileLocation:string, insertLocation:string)`](#insertfile)| string |Inserts the complete specified document intopaoar the specified location. | | 
|[`insertBreak(breakType: string, insertLocation: string)`](#insertBreak)| Void | Inserts the specified [type of break](breakType.md) on the specified location. |All locations may not apply. See method details. | 
|[`insertParagraph(paragraphText: string, insertLocation: string)`](#insertparagraph)| [Paragraph](paragraph.md)  |Inserts a paragraph on the specified location. |All locations may not apply. See method details. | 
|[`insertPictureUrl(base64: string, insertLocation: string)`](#insertPictureUrl)| [Paragraph](paragraph.md)  |Inserts a picture  on the specified location. |All locations may not apply. See method details.| \
|[`insertText(text: string, insertLocation: string)`](#inserttext)| [Range](range.md) | Inserts the specified text on the specified location. | All locations may not apply. See method details. | 
|[`insertHtml(html: string, insertLocation: string)`](#inserthtml)| [Range](range.md)  |Inserts the specified html on the specified location. | All locations may not apply. See method details.| 
|[`insertOoxml(ooxml: string, insertLocation: string)`](#insertooxml)| [Range](range.md)  |Inserts the specified ooxml on the specified location.  | All locations may not apply.See method details.| 
|[`search(text: string)`](#search)| [Ranges](searchResultCollection.md) |Executes a search on the scope of the calling object | Search results are a ranges collection. | 
|[`select(paragraphText: string, insertLocation: string)`](#select)| [Paragraph](paragraph.md)  | Selects and Navigates to the paragraph ||


### Setting Paragraph Properties 
```js
  // playing with a few parapgraph properties, check out how it modifies your first paragrpahs settings!
  var ctx = new Word.RequestContext();
var paras = ctx.document.body.paragraphs;
ctx.load(paras, {select:"text"});
ctx.references.add(paras);


ctx.executeAsync().then(
  function () {
var par = paras.items[0];
par.lineSpacing = 45;
par.alignment = "justified";
par.spaceAfter = 45;
par.firstLineIndent = 1;
par.leftIndent = 2;
par.lineUnitAfter = 2;
par.lineUnitBefore = 5;
par.outlineLevel = 10;

 ctx.executeAsync().then(
      function () {
        console.log("Success!!!" + par.lineSpacing);
     ctx.references.remove(par);
      }
    );

    //console.log("Success! Setting paragraph line spacing to " + par.lineSpacing);
  },
  function (result) {
    console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
    console.log(result.traceMessages);
  }
);



```

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


### InlinePictures 

The colection holds all the inline pictures contained in the scope.

#### Syntax
```js
  document.body.paragraphs  // returns the paragraphs on the body of the document.
  document.sections.getItemAt(0).paragraphs  //returns the paragraphs in the first section of the document.
  document.selection.paragraphs   //returns the paragraphs contained in the selection.

```

#### Returns

[InlinePictures](inlinePictureCollection.md) collection. See [InlinePicture](inlinePicture.md) object.

#### Examples

```js

//gets all the images in the body of the document and then gets the base64 for each.
var ctx = new Word.RequestContext();


var pics = ctx.document.body.inlinePictures;
ctx.load(pics);
ctx.references.add(pics);

ctx.executeAsync().then(
  function () {
    var results = new Array();
  
    for (var i = 0; i < pics.items.length; i++) {
      results.push(pics.items[i].getBase64ImageSrc());
    }
    ctx.executeAsync().then(
      function () {
        for (var i = 0; i < results.length; i++) {
          console.log("pics[" + i + "].base64 = " + results[i].value);
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

### clear

Clears the content of the calling object.

#### Syntax
```js
ctx.document.body.clear();

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


### getHtml

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

### insertText

Inserts the specified text on the specified location.

#### Syntax
```js
var myText = document.body.insertText("Hello World!", "End");
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`text`          | string | Required. Text to be inserted.
`location`          | string | Either "Start" "End"  the body of the document.

#### Returns

[Range](range.md).


#### Examples

```js
var myText = document.body.insertText("Hello World!", "End");

```
[Back](#methods)

### insertHtml

Inserts the specified HTML on the specified location.

#### Syntax
```js
var myRange = document.body.insertHtml("<b>This is some bold text</b>", "End");
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`html`          | string | Required. the HTML to be inserted in the document.
`location`          | string | Either "Start" "End"  the body of the document

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

### insertOoxml

Inserts the specified OOXML on the specified location.

#### Syntax
```js
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
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`ooxml`          | string | Required. OOXML to be inserted.
`location`          | string | Either "Start" "End"  the body of the document
 
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

### insertParagraph

Inserts a paragraph on the specified location.

#### Syntax
```js
var ccs = document.insertParagraph("Some initial text", "Start");
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`text`          | string | Paragrph text. null for blank Paragraph.
`location`          | string | Either "Start" "End"  the body of the document


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

### insertContentControl

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
// wraps the current selection with a content control, then sets a few properties.
var ctx = new Word.RequestContext();
var range = ctx.document.getSelection();

var myContentControl = range.insertContentControl();
myContentControl.tag = "Customer-Address";
myContentControl.title = "Enter Customer Address Here:";
myContentControl.style = "Heading 1";
myContentControl.insertText("One Microsoft Way, Redmond, WA 98052", 'replace');
myContentControl.cannotEdit = true;
myContentControl.appearance = "tags";

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

### search()

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
`location`          | string | Either "Start" "End"  the body of the document.


#### Returns

[Range](range.md) collection.


#### Examples

```js
TBD

```
[Back](#methods)