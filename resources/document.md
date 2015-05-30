# Document 
 Represents a Word document. Main entry point to all interactions with the document. A document is composed of one or more sections(resources/section.md), and a body where the main content of the document resides.

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`body`|  [Body](body.md)   |Represents the body of the document, not includes the header/footer and other section metadata | |
|`saved`|  Bool |Indicates if the document is dirty, and requires to be saved | |
|`selection`| [Range](range.md) |Represents the continous current selection of the document. Since it can expand multiple paragraphs its considered to be a Range Object. | Multiple selection is not supported|



## Relationships
The Worksheet resource has the following relationships defined:

| Relationship     | Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`sections`](#sections)| [Section](section.md) collection |Collection of sections in the current document |Document.Section  |       

## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`getContentControlById(id: string)`](#getcontentcontrolbyid)| [ContentControl](contentControl.md) object |Returns the content control with the specified Id, returns null if the content control does not exist|  |
|[`getContentControlByName(name: string)`](#getcontentcontrolbyname)| [ContentControls](contentControls.md) collection |Returns the collection of the content controls matching the specified name| Since there could be many Content Controls with the same name, this method returns a collection|  
|[`getContentControlByTag(tag: string)`](#getcontentcontrolbytag)| [ContentControls](contentControls.md) collection |Returns the collection of the content controls matching the specified tag| Since there could be many Content Controls with the same name, this method returns a collection |
|[`save(void)`](#save)| Void |Saves the Document | If document has not saved before it will use Word default names (i.e. Document1.docx, etc.) |     



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

### getContentControlById

Gets the content control with the specified ID. 

#### Syntax

```js
var myContentContolId = myContentControl.id;

```

#### Parameters 

Parameter      | Type   | Description
-------------- | ------ | ------------
`id`          | string | Required. Id of the content control.

#### Returns

[ContentControl](contentContol.md) object.

#### Examples

```js
// this is an example of inserting a content control then getting the content control by ID and changing its title. 
var ctx = new Word.WordClientContext();
var myContentControl = ctx.document.body.paragraphs.getItemAt(1).insertContentControl();
var myContentContolId = myContentControl.id;
ctx.executeAsync().then(
    function() {
    }
);


var myCC = ctx.document.getContentControlById(myContentContolId);
ctx.load(myCC);
ctx.executeAsync().then(
    function () {
        var results = new Array();
    	 myCC.title = "this is the new title";
},  function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
    }

);
```
[Back](#methods)


### getContentControlByName

Gets a collection of content controls with the same name/title.

#### Syntax
```js
var ccs = document.getContentControlByName("Address");
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`name`          | string | Required. Name/title of the content control(s) to retrieve.

#### Returns

[ContentControls](contentControls.md) collection.


#### Examples

```js
var ccs = document.getContentControlByName("Address");
```
[Back](#methods)


### getContentControlByTag

Gets a collection of content controls with the same tag.

#### Syntax
```js
var ccs = document.getContentControlByTag("TagForName");
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`tag`          | string | Required. Tag of the content control(s) to retrieve.


#### Returns

[ContentControls](contentControls.md) collection.


#### Examples

```js
var ccs = document.getContentControlByTag("TagForName");
```
[Back](#methods)