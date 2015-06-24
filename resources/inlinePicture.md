# InlinePicture

Represents an inline picture anchored to a paragraph.
## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|\
|`parentContentControl`|  [ContentControl](contentControl.md)   |Returns the content control wrapping the object, if any. | Returns null if no content control|
|`altTextDescription`| string  | Returns or sets a String that represents the alternative text associated with a shape in a Web page. Read/write. | Read/write. |
|`altTextTitle`| string  | Returns or sets a String that contains a title for the specified inline shape. |Read/write. |
|`height`| number  |  Returns or sets the height of an inline shape. | Read/write.|
|`hyperlink`| string  |sets/gets the hyperlink associated with the specified inline shape.  |Read/write. |
|`lockAspectRatio`| bool  | True if the specified image retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it. | Read/write.|
|`width`| number  | Returns or sets the width of an inline shape.  | Read/write.|

## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`getBase64ImageSrc()`](#getbase64imagesrc)| String | Gets the base64 encoded string of the image | | 
|[`insertContentControl()`](#insertcontentcontrol)| [ContentControl](contentcontrol.md)  |Wraps the calling object with a Rich Text content control. |  | 


  


#### Examples
### getBase64ImageSrc
Gets the base64 encoded string of the image

#### Syntax
```js
pics.items[i].getBase64ImageSrc();

```
#### Parameters

None

#### Returns

String.


#### Examples

```js
//gets all the images in the body of the document and then gets the base64 for each.
var ctx = new Word.WordClientContext();


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
// grabs the first paragraph in the document and inserts an image at the end of it, then sets a
// few props, then wraps it inside a content control to finally adjust a few properties of the content control.
var ctx = new Word.WordClientContext();
var paras = ctx.document.body.paragraphs;
ctx.load(paras);

var myImage = paras.getItem(0).insertInlinePictureFromBase64("iVBORw0KGgoAAAANSUhEUgAAAIAAAACABAMAAAAxEHz4AAAAJFBMVEX///9GRkZGRkZGRkZGRkZGRkZGRkZGRkYBpO9/ugDyUCL/uQGm4PjWAAAACHRSTlMBCQ0RFRknMx7uViEAAAB3SURBVGje7dcxCYBQGEXhi6izYBHB0RIiiAXkzW5iAMEKFnCwguVscJd/ecM5Ab79SNHK5FqlZXeNql/XIx23awMAAAAAAAAAAAAAAAAAyBwIvzNJxeyapLZ3Naou1ykNn6sDAAAAAAAAAAAAAAAAAMgcCL9ztB/UhshWs1l/WAAAAABJRU5ErkJggg==", Word.InsertLocation.end);


myImage.width = 100;
myImage.height = 100;
myImage.lockAspectRatio = true;
myImage.hyperlink = "http://dev.office.com";
var myCC = myImage.insertContentControl();
myCC.title = "My Image";
myCC.appearance = "tags";

ctx.references.add(myImage);

ctx.executeAsync().then(
	function () {
		console.log("*" + myImage.id);
		console.log("Success");
	},
	function (result) {
		console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
		console.log(result.traceMessages);
	}
);
```
[Back](#methods)


