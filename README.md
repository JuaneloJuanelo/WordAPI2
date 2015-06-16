# Word JavaScript APIs
Welcome to the new Word Javascript API! We hope you enjoy it an find it useful. Send feedback to juanbl@microsoft.com


## Release Notes 

For this release of the JavaScript APIs for Word we focused on the following functional areas:
 1. **Basic document navigation:** On top of having access to the user's current selection, we are also providing ways to traverse the entire document by exposing collections of two of the most important objects in Word: Paragraphs and Content Controls.

 2. **Insertion of content:** We are enabling developers to insert fully-formatted content into the Word document and to append or prepend inserted content. Developers can insert either formatted text, HTML, or Office Open XML. Developers can also reuse content from other Word documents by inserting a Word file into the current document.

 3. **Full control to Paragraphs and Content Controls:** We are providing access to the most important properties of these objects.

 4.  **Search:** Developers can search for content in the document and then iterate and manipulate the search results.

 5. **Range:**  For selection and search results as well as document, paragraph, and content control objects, the developer can access the represented range and its most relevant properties.


## Main Objects  

* [Document](resources/document.md): Represents a Word document. It's the entrypoint to all interactions with the document. A document is composed of one or more sections(resources/section.md) and a body where the main content of the document resides.
* [Paragraph](resources/paragraph.md):  Represents a single paragraph in a selection, range, or document. It's a member of the Paragraphs collection. The paragraphs collection includes all the paragrpahs in a selection range or document. The Paragraph object is a member of the Paragraphs collection.
* [ContentControl](resources/contentControl.md): Represents an individual content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as paragraphs of formatted text and other content controls. The ContentControl object is a member of the ContentControls collection.
* [Section](resources/section.md):  Represents a single section in a document. Sections are commonly used to define the potentially different headers and footers as well as different page layout configurations that a document can define. 
* [Range](resources/range.md): Represents a contiguous area in a document. Range objects are independent of the selection. That is, you can define and manipulate a range without changing the selection.

* [Picture](resources/inlinePicture.md): Represents a picture anchored to a Paragraph.
* [Font](resources/font.md): Represents an object to provide text formatting to a Range.

## Additional Content Needed

* Explanation of the new pipeline, including the Ctx, local objects, and properties vs relations.
o   ExecuteAsync is the only place where there’s a “cross-boundary” call to the Office application.
o	The purpose of load() is to tell Office what properties are being requested so that the application doesn’t waste time/space returning stuff that’s unnecessary.
o	What’s a “relation”? Is it different than a method? Is it different than a “relationship”? What does it mean to say that relations are not “loaded” by default? I don’t think I understand this part. If entireRow is a relation for Range, then it should be included in the Range object page, right? (maybe that one’s just in progress now)

