# Ranges

A collection of [Range](range.md) objects. Usually a result of a document.search() operation.


## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`items`|  array |Array containing the [Range](range.md) objects in the given scope. ||
|`count`|  integer |Number of [Range](range.md) objects  in the scope |Read-Only|



## Relationships
None  

## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`getItem(index:integer)`](#getitem)| [Range](range.md)   | Gets a[Range](range.md) by its index in the collection. || 


  



#### Example
```js

///Search example, returns a collection of ranges!

var ctx = new Word.WordClientContext();

var options = Word.SearchOptions.newObject(ctx);
options.matchCase = false

var results = ctx.document.body.search("video", options);
ctx.load(results);
ctx.references.add(results);

ctx.executeAsync().then(
    function () {
        console.log("found count = " + results.items.length);
        for (var i = 0; i < results.items.length; i++) {
            results.items[i].font.color = "#FF0000"    // Change color to Red
            results.items[i].font.highlightColor = "#FFFF00";
            results.items[i].font.bold = true;
            if(i==3)
                results.items[i].select();
        }
        ctx.references.remove(results);
        ctx.executeAsync().then(
            function () {
                console.log("deleted");
            }
        );
    }
);

```



