# InlinePicture

Represents an inline picture anchored to a paragraph.
## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|\
|`parentContentControl`|  [ContentControl](contentControl.md)   |Returns the content control wrapping the object, if any. | Returns null if no content control|
|`altTextDescription`| string  | Returns or sets a String that represents the alternative text associated with a shape in a Web page. Read/write. | Read/write. |
|`altTextTitle`| string  | Returns or sets a String that contains a title for the specified inline shape. |Read/write. |
|`height`| number  |  Returns or sets the height of an inline shape. | |
|`hyperlink`| string  |sets/gets the hyperlink associated with the specified inline shape.  | |
|`id`| number  | | A session-wise identifier of the image |
|`lockAspectRatio`| bool  | True if the specified image retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it. | R/W |
|`width`| number  | Returns or sets the width of an inline shape.  | |



#### Examples

#### Syntax
```js
// grabs the first paragraph in the document and inserts an image at the end of it, then sets a
// few props.
var ctx = new Word.WordClientContext();
var paras = ctx.document.body.paragraphs;
ctx.load(paras);

var myImage = paras.getItem(0).insertInlinePictureFromUrl("http://dev.office.com/Media/Default/App%20Awards/AppAwards.png", Word.InsertLocation.end, false, true);

myImage.width = 100;
myImage.height = 100;
myImage.lockAspectRatio = true;
myImage.hyperlink = "http://dev.office.com";



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



