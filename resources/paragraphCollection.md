# Paragraphs
A collection of Paragraph objects in a selection, range, or document

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`items`|  array |Array containing the [Paragraph](paragraph.md) objects in the given scope. ||


## Relationships
None  

## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`getItem(index:integer)`](#getitem)|[Paragraph](paragraph.md)    | Gets a [Paragraph](paragraph.md)  by its index in the collection. || 

#### Example
```js
// this example iterates all the paragraphs in the documents and reports back the length and text of each paragraph in the document
var ctx = new Word.RequestContext();
var paras = ctx.document.body.paragraphs;
ctx.load(paras,{select:"text"});

ctx.executeAsync().then(
  function () {
    for (var i = 0; i < paras.items.length; i++) {
      console.log("paras[" + i + "].content  = " + paras.items[i].text);
    }
  },
  function (result) {
    console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
    console.log(result.traceMessages);
  }
);


```



