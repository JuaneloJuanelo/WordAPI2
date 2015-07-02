# Ranges

A collection of [Range](range.md) objects. Usually a result of a document.search() operation.


## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`items`|  Array |Array containing the [Range](range.md) objects in the given scope. ||
|`count`|  Number |Number of [Range](range.md) objects  in the scope |Read-Only|



## Relationships
None  

## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`getItem(index:Number)`](#getitem)| [Range](range.md)   | Gets a[Range](range.md) by its index in the collection. || 


  



#### Example
```js

///Search example, returns a collection of ranges!


var ctx = new Word.RequestContext();
var options = Word.SearchOptions.newObject(ctx);

options.matchCase = false

var results = ctx.document.body.search("Video", options);
ctx.load(results, {select:"text, font/color", expand:"font"});
ctx.references.add(results);

ctx.executeAsync().then(
  function () {
    console.log("Found count: " + results.items.length + " " + results.items[0].font.color );
    for (var i = 0; i < results.items.length; i++) {
      results.items[i].font.color = "#FF0000"    // Change color to Red
      results.items[i].font.highlightColor = "#FFFF00";
      results.items[i].font.bold = true;
      if (i == 0)
        results.items[i].select();
    }
    ctx.references.remove(results);
    ctx.executeAsync().then(
      function () {
        console.log("Deleted");
      }
    );
  }
);

```



