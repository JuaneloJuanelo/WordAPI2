# Word JavaScript APIs
Welcome to the new Word Javascript API! We hope you enjoy it and find it useful. Please open [issues](https://github.com/JuaneloJuanelo/WordAPI2/issues) if you find errors in the documentation or if you have suggested content or examples that we should add to this documentation. We're open to community contributions if you early adopters have found some useful information.


## Release Notes for build 4220

This is the first release of JavaScript APIs for Word and we focused on the following functional areas:
 1. **Document navigation:** You can traverse the document by accessing the [paragraphs](resources/paragraph.md) and [content control](resources/contentControl.md) collections. You can also access the user's
 current selection to get objects in the document.

 2. **Insert content:** A user's selection can be used to add formatted content into a document. This includes appending 
 	and prepending content to the core Word objects. Inserted content can be formatted text, HTML or Office Open XML.
	You can also insert the entire contents of another Word file. 
 
 3. **Full access to Paragraphs and Content Controls** 

 4. **Search:** Search for content in the document.

 5. **Range:**  Selections, search results, document, paragraph, and content control objects can be accessed by a range object.


## Main Objects  

* [Document](resources/document.md): The Document object is the top level object. A Document objects contains one or more 
[sections](resources/section.md), a body that contains the content of the document, and header/footer information.
* [Paragraph](resources/paragraph.md): A Paragraph object represents a single paragraph in a selection, range or document. 
You can access a paragraph through the paragraphs collection in a selection, range, or document. 
* [ContentControl](resources/contentControl.md): A ContentControl object is a container for content. It is a bound and
 potentially labeled region in a document that serves as a container for specific types of content. For example, content 
 controls can contain contents such as paragraphs of formatted text and other content controls. You can access a 
 content control through the content control collection of the document, document body, paragraph, range, or on a content control.
* [Section](resources/section.md):  A Section object is commonly used to define different header and footers as well as 
different page layout configurations of a document. You can access sections from the Document object. 
* [Range](resources/range.md): A Range object represents a contiguous area in a document. You get a Range object when you
 get a selection, insert content into the body, insert content into a content control, insert content into a paragraph, 
 or get a search result. You can define and manipulate a range without changing the selection.
* [Selection](resources/selection.md): The Selection object represents the user's selection in the document, or the 
 current insertion point.
* [Picture](resources/inlinePicture.md): A Picture object represents an inline image. You can access the inline picture
 collection of the body, content control, paragraph objects.
* [Font](resources/font.md): The Font object provides text formatting to a body, content control, paragraph, or range.

**Figure 1.  Word API object model**

![A simple diagram of the Word API object model. The Word object is at the top level. The next object is the Document object. Under the Document objects you can access the collection of section objects.](resources\images\wordAPIObjects.png)

## Programming notes

 The Word.WordClientContext() method returns the context for working with the Word object model. All actions that target a Word 
 document start by getting this context. 

		var ctx = new Word.WordClientContext();

