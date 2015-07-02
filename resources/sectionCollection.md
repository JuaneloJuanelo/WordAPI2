# Sections
A collection of [Section](section.md) objects in a document

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`items`|  Array |Array containing the [Section](section.md) objects in the given scope. ||


## Relationships
None  

## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`getItem(index:Number)`](#getitem)|[Section](section.md)   | Gets a [Section](section.md)  by its index in the collection. || 

#### Example
```js
//gets access to seciton headers

var ctx = new Word.RequestContext();

var mySections  = ctx.document.sections;
ctx.load(mySections);

var myFooter = mySections.getItem(0).getFooter("primary");
myFooter.insertText("this is a footer!!","end");
myFooter.insertContentControl();

ctx.executeAsync().then(
	function(){
				   
				   console.log("Success!!");
			
                              console.log("Success!!");
	}
	
);


```



