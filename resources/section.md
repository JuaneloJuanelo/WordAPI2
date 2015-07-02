# Section 
 Represents a Word document. Main entry point to all interactions with the document. A document is composed of one or more sections(resources/section.md), and a body where the main content of the document resides.

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`body`|  [Body](body.md)   |Represents the body of the section, not includes the header/footer and other section metadata | |




## Relationships
The Worksheet resource has the following relationships defined:
None


## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`getFooter(type:String)`](#getfooter)| [Body](body.md) |Gets the footer of the specified type for the referred section | Type can be: "primary", "firstPage" or  "evenPages" |     
|[`getHeader(type:String)`](#getheader)| [Body](body.md) |Gets the header of the specified type for the referred section | Type can be: "primary", "firstPage" or  "evenPages"|





### Methods 

#### Examples

### getFooter

Gets the footer of the specified type for the referred section

#### Syntax
```js
var myFooter = mySections.getItem(0).getFooter("primary");

```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`type`          | String | Required. Type of header  "primary", "firstPage" or  "evenPages"


#### Returns

[Body](body.md).


#### Examples

```js
//Inserting text in the footer  of the document, also adds a content control :)

var ctx = new Word.RequestContext();

var mySections  = ctx.document.sections;
ctx.load(mySections);

var myFooter = mySections.getItem(0).getFooter("primary");
myFooter.insertText("this is a header!!","end");
myFooter.insertContentControl();

ctx.executeAsync().then(
	function(){
				   
				   console.log("Success!!");
			
                              console.log("Success!!");
	}
	
);

```
[Back](#methods)


### getHeader

Gets the header of the specified type for the referred section

#### Syntax
```js
var myHeader = mySections.getItem(0).getHeader("primary");

```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`type`          | String | Required. Type of header  "primary", "firstPage" or  "evenPages"


#### Returns

[Body](body.md).


#### Examples

```js
//Inserting text in the footer  of the document, also adds a content control :)

var ctx = new Word.RequestContext();

var mySections  = ctx.document.sections;
ctx.load(mySections);

var myHeader = mySections.getItem(0).getHeader("primary");
myHeader.insertText("this is a header!!","end");
myHeader.insertContentControl();

ctx.executeAsync().then(
	function(){
				   
				   console.log("Success!!");
			
                              console.log("Success!!");
	}
	
);

```
[Back](#methods)