# Body 
Represents the body of a [document](document.md) or a [section](section.md).

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`font`|  [Font](font.md) | Contains the text format of the body.| Use this to get and set font name, size, color, and other properties. |
|`parentContentControl`|  [ContentControl](contentControl.md)   |Returns the content control that contains the body.| Returns null if there isn't a parent content control.|
|`style`| String |Name of the style been used. | This is the name of an pre-installed or custom style.|
|`text`| String |Retrieves the document body as plain text. | Read-Only.  Use the [insertText](inserttext) method to insert text. [Go to the sample](#gettext). |

## Relationships
The Worksheet resource has the following relationships defined:

| Relationship     | Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`contentControls`](#contentcontrols)| [ContentControls](contentControlCollection.md) collection |Collection of [contentControl](contentControl.md) objects  in the current document. | Includes content controls in the body of the document.|
|[`inlinePictures`](#inlinepictures)| [InlinePictures](inlinePictureCollection.md) collection |Collection of [inlinePicture](inlinePicture.md) objects within the body. |Does not include floating images.  | 
|[`paragraphs`](#paragraphs)| [Paragraphs](paragraphCollection.md) collection |Collection of [paragraph](paragraph.md) objects within the body. |   |      
    


## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`clear()`](#clear)| Void | Clears the contents of the body object. | Undo operation by the user is supported. | 
|[`getHtml()`](#gethtml)| String  | Gets the HTML representation  of the body object. | IMPORTANT: we are deprecating this method in favor of the property.| 
|[`getOoxml()`](#getooxml)| String  | Gets the Office Open XML (OOXML) representation  of the body object. | IMPORTANT: we are deprecating this method in favor of the property. | 
|[`insertBreak(breakType: String, insertLocation: String)`](#insertbreak)| Void | Inserts the specified [break type](breakType.md) at the specified location. | All locations may not apply. See method details. | 
|[`insertContentControl()`](#insertcontentcontrol)| [ContentControl](contentcontrol.md)  |Wraps the body object with a Rich Text content control. |  | 
|[`insertFile(fileLocation:String, insertLocation:String)`](#insertfile)| String |Inserts the a document into the specified location. | This methood may get deprecated for security resons.| 
|[`insertText(text: String, insertLocation: String)`](#inserttext)| [Range](range.md) | Inserts text at the specified location. | All locations may not apply. See method details. | 
|[`insertHtml(html: String, insertLocation: String)`](#inserthtml)| [Range](range.md)  |Inserts html at the specified location. | All locations may not apply. See method details.| 
|[`insertOoxml(ooxml: String, insertLocation: String)`](#insertooxml)| [Range](range.md)  |Inserts ooxml at the specified location.  | All locations may not apply.See method details.| 
|[`insertParagraph(paragraphText: String, insertLocation: String)`](#insertparagraph)| [Paragraph](paragraph.md)  |Inserts a paragraph at the specified location. |All locations may not apply. See method details. | 
|[`search(searchText : String, searchOptions: searchOptions)`](#search)| [Ranges](searchResultCollection.md) |Performs a search with the specified [searchOptions](searchOptions.m) on the scope of the calling object | Search results are a ranges collection. | 



### ContentControls 

The collection holds all the content controls in the document.

#### Syntax
```js
  document.contentControls

```

#### Returns

[ContentControls](contentControlCollection.md) collection. See the [ContentControl](contentControl.md) object for more information.

#### Examples

```javascript

    // enumerates all the content controls in the document
    var ctx = new Word.RequestContext();
    var cCtrls = ctx.document.body.contentControls;
    ctx.load(cCtrls,{select:'appearance,text'});  // just need these properties!

    ctx.executeAsync().then(
    function () {
            var results = new Array();

            for (var i = 0; i < cCtrls.items.length; i++) {
               console.log("contentControl[" + i + "].text = " + cCtrls.items[i].text + " Appearance:" +
                            cCtrls.items[i].appearance );
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

The collection holds all the paragraphs in the scope.

#### Syntax
```js
    // returns the paragraphs in the document body.
    document.body.paragraphs  
    
    //returns the paragraphs in the first section of the document.
    document.sections.getItemAt(0).paragraphs  
    
    //returns the paragraphs contained in the selection.
    document.selection.paragraphs   

```

#### Returns

[Paragraphs](paragraphCollection.md) collection. See [Paragraph](paragraph.md) object.

#### Examples

```js

    // this example iterates all the paragraphs in the documents and reports back the 
    // length and text of each paragraph in the document
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

The collection contains all of the inline pictures contained in the body.

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

### getText

Gets the plain text value  of the calling object.

#### Syntax
```js
myBody.text
```
#### Parameters

None

#### Returns

[Range](range.md).


#### Examples

```js

//gets the text of the entire body.
var ctx = new Word.RequestContext();
var myBody = ctx.document.body
ctx.load(myBody, {select:'text'});
ctx.executeAsync().then(
    function () {
    console.log(myBody.text);    
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
var myTHTML  = document.body.getHtml();
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
//get inserts some text at the end of the document.
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

var results = ctx.document.body.search("Hello", options);  //searches for hello in the document
```

#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`searchText`          | String | Required. Text to be searched.
`searchOptions` | [SearchOptions](searchOptions.md) | Required. Options for the search.

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



### select()

Selects the specified Range. Scrolls to the selection. 

#### Syntax
```js
results.items[i].select();
```
#### Parameters

No Parameters.

#### Returns

Void


#### Examples

```js
///Search and selects the first occurrence! 

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
      if (i == 0)
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