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
|[`contentControls`](#contentcontrols)| [ContentControl](contentControl.md) collection |Collection of content controls in the current document | Includes content controls on the headers/footer and in the body of the document.  | 

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
