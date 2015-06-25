# ContentControls

A collection of ContentControl objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain content such as dates, lists, or paragraphs of formatted text.


## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`items`|  array |Array containing the content controls in the given scope. ||
|`count`|  integer |Number of content controls in the scope |Read-Only|

## Relationships
None  

## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[`getById()`](#clear)| [contentControl](contentControl.md) | Gets a content control by its id. | | 
|[`getByTag(tag:string )`](#getbytag)| contentControls(contentControlCollection.md)  |Gets the content controls that have the specified tag. | | 
|[`getByTitle(title:string)`](#getbytitle)| contentControls(contentControlCollection.md) |Gets the content controls that have the specified tag. |  | 
|[`getItemAt(index:integer)`](#getitemat)| [contentControl](contentControl.md)   | Gets a content control by its index in the collection. || 


  



#### Example
```js
// gets Content control by tags and prints its value.
var ctx = new Word.WordClientContext();
var ccs = ctx.document.contentControls.getByTag("Customer-Address");
ctx.load(ccs);
ccs.getItemAt(0).font.italic = true;
 
ctx.executeAsync().then(
     function () {
         var ccText =   ccs.getItemAt(0).getText();
         ctx.executeAsync().then(
             function(){
                  console.log("Content Control Text: " + ccText.value);

             }
         )  ;
        
     },
     function (result) {
         console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
         console.log(result.traceMessages);
     }
);


```



