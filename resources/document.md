# Document 
The Document object is the top level object. A Document object contains one or more 
sections, content controls, and the body that contains the contents of the document.

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`body`|  [Body](body.md)   |Gets the body of the document. | |
|`saved`|  Bool | Indicates whether the document has been changed. | |



## Relationships
The Document object has the following relationships defined:

| Relationship     | Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`contentControls`](#contentcontrols)| [ContentControlCollection](contentControlCollection.md)  |Collection of [contentControl](contentcontrol.md) objects  in the  document | Includes content controls on the headers/footer and in the body of the document.  | 
|[`sections`](#sections)| [SectionCollection](sectionCollection.md) |Collection of [section](sectionCollection.md) objects in the  document |  |       


## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`getSelection()`](#getselection)| [Range](range.md) |Represents the continuous current selection of the document. Since it can expand multiple paragraphs its considered to be a Range Object. | Multiple selection is not supported|
|load()| Document | Loads the Document |  |
|[`save()`](#save)| Void |Saves the Document | If document has not been saved before it will use the Word default file naming convention. |     

## API Specification

### ContentControls 

Get the content control collection of the document.

#### Syntax
```js
  document.contentControls;

```

#### Returns

[ContentControlCollection](contentControlCollection.md)

#### Example

```js
// Get all of the content controls in the document
var ctx = new Word.WordClientContext();
var cCtrls = ctx.document.contentControls;
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

Gets all of the section objects in the document.

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
