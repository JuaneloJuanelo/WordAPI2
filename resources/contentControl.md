# ContentControl

An individual content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as dates, lists, or paragraphs of formatted text. On this release, only rich text content controls are supported. The ContentControl object is a member of the ContentControls collection.


## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`appearance`|  string |Returns or sets the appearance of the content control. |RW. Can be 'boundingBox', 'tags' or 'hidden' |
|`cannotDelete`|  boolean |Returns or sets a Boolean that represents whether the user can delete a content control from the active document |RW. |
|`cannotEdit`|  boolean | Returns or sets a Boolean that represents whether the user can edit the contents of a content control. |RW. |
|`color`|  Number |   Returns or sets the color of the content control.        | Color is set in "#FFFFFF" format or color name|
|`font`|  [Font](font.md) | Entry point for formatting content.|  Exposes font name, size, color, and other properties. |
|`id`|  string |Returns a String that represents the identification for a content control. |Read-only|\
|`parentContentControl`|  [ContentControl](contentControl.md)   |Returns the content control wrapping the object, if any. | Returns null if no content control|
|`removeWhenEdited`|  boolean |  Removes the content control after edited.         ||
|`title`|  string  |  Returns or sets a String that represents the title for a content control.   | |
|`type`|  string  | Returns or sets  the type for a content control.          |Only rich text content controls are supported|\
|`style`| String |Name of the style been used. | This is the name of an pre-installed or custom style.|
|`tag`| String |Returns or sets a String that represents a value to identify a content control. | RW and might be duplicated|



## Relationships
The Content Control resource has the following relationships defined:

| Relationship     | Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`contentControls`](#contentcontrols)| [ContentControls](contentControls.md) collection |Collection of [contentControl](#contentcontrol.md) objects  in the current document | Includes content controls on the headers/footer and in the body of the document.  | 
|[`inlinePictures`](#inlinepictures)| [InlinePictures](inlinePictures.md) collection |Collection of [inlinePicture](#inlinePicture.md) objects within the body. |Does not include floating images.  | 
|[`paragraphs`](#paragraphs)| [Paragraphs](paragraphs.md) collection |Collection of [paragraph](#paragraph.md) objects within the content control. |  |      

       

## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`clear()`](#clear)| Void | Clears the content of the calling object. | Undo operation by the user is supported. | 
|[`delete(keepContent:boolean )`](#deleteelement)| Void  |Deletes the content control and its content from the document, users may keep the content if send true as parameter. | | 
|[`getText()`](#gettext)| String |Gets the plain text of the calling object. | IMPORTANT: we are deprecating this method in favor of the property | 
|[`getHtml()`](#gethtml)| String  | Gets the HTML representation  of the calling object. | IMPORTANT: we are deprecating this method in favor of the property| 
|[`getOoxml()`](#getooxml)| String  | Gets the Office Open XML (OOXML) representation  of the calling object. | IMPORTANT: we are deprecating this method in favor of the property | 
|[`insertContentControl()`](#insertcontentcontrol)| [ContentControl](contentcontrol.md)  |Wraps the calling object with a Rich Text content control. |  | 
|[`insertFile(fileLocation:string, location:string)`](#insertfile)| String |Inserts the complete specified document into the specified location. | | 
|[`insertBreak(paragraphText: string, insertLocation: string)`](#insertBreak)| [Paragraph](paragraph.md)  |Inserts a paragraph on the specified location. |All locations may not apply. See method details. | 
|[`insertParagraph(paragraphText: string, insertLocation: string)`](#insertparagraph)| [Paragraph](paragraph.md)  |Inserts a paragraph on the specified location. |All locations may not apply. See method details. | 
|[`insertPictureBase64(url: string, insertLocation: string)`](#insertPictureBase64)| [Paragraph](paragraph.md)  |Inserts a paragraph on the specified location. |All locations may not apply. See method details. | 
|[`insertText(text: string, insertLocation: string)`](#inserttext)| [Range](range.md) | Inserts the specified text on the specified location. | All locations may not apply. See method details. | 
|[`insertHtml(html: string, insertLocation: string)`](#inserthtml)| [Range](range.md)  |Inserts the specified html on the specified location. | All locations may not apply. See method details.| 
|[`insertOoxml(ooxml: string, insertLocation: string)`](#insertooxml)| [Range](range.md)  |Inserts the specified ooxml on the specified location.  | All locations may not apply.See method details.| 
|[`select(paragraphText: string, insertLocation: string)`](#select)| [Paragraph](paragraph.md)  | Selects and Navigates to the paragraph ||
  


### ContentControls 

The colection holds all the content controls in the document.

#### Syntax
```js
  document.contentControls

```

#### Returns

[ContentControls](contentControls.md) collection. See ContentControl(contentControl.md) object.

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

The colection holds all the paragraphs in the scope.

#### Syntax
```js
  document.body.paragraphs  // returns the paragraphs on the body of the document.
  document.sections.getItemAt(0).paragraphs  //returns the paragraphs in the first section of the document.
  document.selection.paragraphs   //returns the paragraphs contained in the selection.

```

#### Returns

[Paragraphs](paragraphs.md) collection. See [Paragraph](paragrph.md) object.

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

### clearContent

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

//the follwoing snippet clears the content of the document's body.
var ctx = new Word.WordClientContext();

ctx.document.body.clearContent();

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

### getText

Gets the plain text value  of the calling object.

#### Syntax
```js
var myText  = document.body.getText();
```
#### Parameters

None

#### Returns

[Range](range.md).


#### Examples

```js
var ctx = new Word.WordClientContext();
var text = ctx.document.body.getText();
ctx.load(text);

ctx.executeAsync().then(
    function () {
        console.log("Document Text:" + text);
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
var myRange = document.body.insertHtml("<b>This is some bold text</b>", "End");

```
[Back](#methods)

### insertOoxml

Inserts the specified OOXML on the specified location.

#### Syntax
```js
var myRange = document.body.insertOoxml("<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document mc:Ignorable="w14 w15 wp14" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <w:body>
          <w:p>
            <w:pPr>
              <w:spacing w:before="360" w:after="0" w:line="480" w:lineRule="auto"/>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
              <w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t>
            </w:r>
            <w:bookmarkStart w:id="0" w:name="_GoBack"/>
            <w:bookmarkEnd w:id="0"/>
          </w:p>
          <w:p/>
          <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
            <w:cols w:space="720"/>
          </w:sectPr>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>","End");
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
var myRange = document.body.insertOoxml("<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document mc:Ignorable="w14 w15 wp14" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <w:body>
          <w:p>
            <w:pPr>
              <w:spacing w:before="360" w:after="0" w:line="480" w:lineRule="auto"/>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
              <w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t>
            </w:r>
            <w:bookmarkStart w:id="0" w:name="_GoBack"/>
            <w:bookmarkEnd w:id="0"/>
          </w:p>
          <w:p/>
          <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
            <w:cols w:space="720"/>
          </w:sectPr>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>","End");

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
var ccs = document.insertParagraph("Some initial text", "Start");
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
var ccs = document.body.insertContentControl();

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

[Ranges](ranges.md) collection.


#### Examples

```js
var searchResults = document.body.search("Sales Report");

```
[Back](#methods)


### insertFile

Inserts the specified file on the specified location.

#### Syntax
```js
var myDoc = document.body.insertFile("http://mylibrary/myDoc.docx", "End");

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
var myDoc = document.body.insertFile("http://mylibrary/myDoc.docx", "End");


```
[Back](#methods)