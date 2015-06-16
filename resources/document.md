# Document 
 Represents a Word document. It's the entrypoint to all interactions with the document. A document is composed of one or more sections(resources/section.md) and a body where the main content of the document resides.

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`body`|  [Body](body.md)   |Represents the body of the document, not including the header, footer, and other section metadata | |
|`saved`|  Bool |Indicates if the document is dirty and needs to be saved | |
|`selection`| [Range](range.md) |Represents the continous current selection of the document. It can span multiple paragraphs. | Discontinuous selections are not supported.|



## Relationships
The Worksheet resource has the following relationships defined:

| Relationship     | Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`contentControls`](#contentcontrols)| [ContentControls](contentControls.md) collection |Collection of [contentControl](#contentcontrol.md) objects  in the current document | Includes content controls in the header, footer, and body of the document.  | 
|[`sections`](#sections)| [Sections](sections.md) collection |Collection of [section](#sections.md) in the  document |Document.Section  |       


## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`save()`](#save)| Void |Saves the Document | If document has never been saved it will use Word's default names (eg. Document1.docx). |     


### ContentControls 

This colection holds all the content controls in the document.

#### Syntax
```js
  document.contentControls

```

#### Returns

[ContentControls](contentControls.md) collection.

#### Examples

##### <We should find an example that does not need two executeAsync() calls, which makes it confusing. It's not necessary to get the length of each one, for instance.>

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


### Sections 

Contains each of the section objects composing the document.

#### Syntax
```js
  document.sections

```

#### Returns

[Section](section.md) collection.

#### Examples

##### <Same problem with this example; let's find something simpler.>

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
