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
|[`contentControls`](#contentcontrols)| [ContentControls](contentControlCollection.md) collection |Collection of [contentControl](contentcontrol.md) objects  in the current document | Includes content controls on the headers/footer and in the body of the document.  | 
|[`sections`](#sections)| [Sections](sectionCollection.md) collection |Collection of [section](sectionCollection.md) objects in the  document |Document.Section  |       


## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`save()`](#save)| Void |Saves the Document | If document has not saved before it will use Word default names (i.e. Document1.docx, etc.) |     
|[`getSelection()`](#getselection)| [Range](range.md) |Represents the continous current selection of the document. Since it can expand multiple paragraphs its considered to be a Range Object. | Multiple selection is not supported|


### ContentControls 

The colection holds all the content controls in the document.

#### Syntax
```js
  document.contentControls

```

#### Returns

[ContentControls](contentControlCollection.md) collection.

#### Examples

```js
// enumerates all the content controls in the document
var ctx = new Word.WordClientContext();
var cCtrls = ctx.document.body.contentControls;
ctx.load(cCtrls);

ctx.executeAsync().then(
    function () {
        var results = new Array();
        
        for (var i = 0; i < cCtrls.items.length; i++) {
            results.push(cCtrls.getItemAt(i).getText());
        }
        ctx.executeAsync().then(
            function () {
                for (var i = 0; i < results.length; i++) {
                    console.log("contentControl[" + i + "].length = " + results[i].value);
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
//traversing paragraphs...
var ctx = new Word.WordClientContext();


var mySections = ctx.document.sections;
ctx.load(mySections);

var paras = mySections.getItem(0).body.paragraphs;
ctx.load(paras);
ctx.references.add(paras);


ctx.executeAsync().then(
    function () {
        var results = new Array();
        for (var i = 0; i < paras.items.length; i++) {
            results.push(paras.getItem(i).getText());
        }
        ctx.executeAsync().then(
            function () {
                for (var i = 0; i < results.length; i++) {
                    console.log("paras[" + i + "].content  = " + results[i].value);
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
