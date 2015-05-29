# Word JavaScript APIs
Welcome to the new Word Javascript API! We hope you enjoy it an find it useful. Send feedback to juanbl@microsoft.com


## Release Notes 

This is the first release of JavaScript APIs for Word and we focused on the following functional areas:
 1. **Basic document navigation:** On top of having access to the user's current selection, we are also providing ways to traverse the document by exposing collection of two of the most important objects in Word: Paragraphs and Content Controls, having easy access to the content of the entire document.

 2. **Insertion of content:** Once positioned in a location of the document to add content, we are enabling developers to insert fully formatted content into Word document and capabilities to do append/prepend before/after type of insertions against our main set of objects. Developers can insert either formatted text, HTML, Office Open XML. Developers are also enabled to reuse content from other Word documents by inserting a Word file into the current document.

 3. **Full control to Paragraphs and Content Controls:** We are providing access to the most important properties of these objects.

 4.  **Search:** Developers can search for content in the document and then iterate and manipulate the search results.

 5. **Range notions:**  For selection and search results as well as document, paragraph and content control objects, developer can access the represented range and its most relevant properties.


## Main Objects  

* [Document](resources/document.md): Represents a Word document. Main entry point to all interactions with the document. A document is composed of one or more sections(resources/section.md), and a body where the main content of the document resides.
* [Section](resources/section.md):  Represents a single section in a document. Sections are commonly used to define the potentially different header and footers as well as different page layout configurations that a document can define. 
* [Paragraph](resources/paragraph.md):  Represents a single paragraph in a selection, range or document. Its a member of the paragraphs collection. The paragraphs collection includes all the paragrpahs ina selection range or document. 
* [ContentControl](resources/contentControl.md): 
* [Range](resources/range.md): 
* [Picture](resources/picture.md):
* [Font](resources/font.md):