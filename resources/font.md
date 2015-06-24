# Font

Contains font attributes (such as font name, font size and color) usally applicable to a Range.

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`bold`| bool  | True if the font is formatted as bold. Read/write Long.| |
|`color`| string  | Returns or sets the color for the specified font. Read/write . |  Like "#FF00FF" or color name |
|`doubleStrikeThrough`| boolean  |True if the specified font is formatted as double strikethrough text.| |
|`highlightColor`| string  | | |
|`italic`| bool  | True if the font or range is formatted as italic.  | Read/write |
|`name`| string  | Returns or sets the name of the specified object.  |Read/write |
|`size`| number  | Returns or sets the font size, in points.| Read/write|
|`strikeThrough`| bool  | True if the font is formatted as strikethrough text.|Read/write |
|`subscript`| bool  |True if the font is formatted as subscript. | Read/write |
|`superscript`| bool  | True if the font is formatted as superscript. | Read/write|
|`underline`|  bool  | Returns or sets the type of underline applied to the font. |Read/write |



#### Examples

#### Syntax
```js
var ctx = new Word.WordClientContext();
var para = ctx.document.body.paragraphs.getItemAt(0);
var font = para.font;

font.size = 32;
font.bold = true;
font.color = "#0000ff";
font.highlightColor = "#ffff00";

ctx.executeAsync().then(
    function () {
        console.log("Success");
    },
    function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
    }
);
```



