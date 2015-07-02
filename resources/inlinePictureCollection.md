# InlinePictures

A collection of [InlinePicture](inlinePicture.md) objects. 


## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`items`|  Array |Array containing the [InlinePicture](inlinePicture.md) objects in the given scope. ||
|`count`|  Number |Number of [InlinePicture](inlinePicture.md) objects  in the scope |Read-Only|



## Relationships
None  

## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`getItem(index:Number)`](#getitem)| [InlinePicture](inlinePicture.md)   | Gets a [InlinePicture](inlinePicture.md) by its index in the collection. || 


  



#### Example
```js
//gets all the images in the body of the document and then gets the base64 for each.
var ctx = new Word.RequestContext();


var pics = ctx.document.body.inlinePictures;
ctx.load(pics);
ctx.references.add(pics);

ctx.executeAsync().then(
    function () {
        var results = new Array();

        for (var i = 0; i < pics.items.length; i++) {
            results.push(pics.items[i].getBase64ImageSrc());
        }
        ctx.executeAsync().then(
            function () {
                for (var i = 0; i < results.length; i++) {
                    console.log("pics[" + i + "].base64 = " + results[i].value);
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



