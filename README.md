A pure Java library for converting Microsoft Word (docx) files to Java2D.

## Quick Start
Create a `DocxReader` and render it to your `GraphicsBuilder` implementation:

```java
Renderer renderer = new DocxRenderer(new File("input.docx"));

renderer.render(new GraphicsBuilder() {
	@Override
	public Graphics2D nextPage(int pageWidth, int pageHeight) {
		BufferedImage bi = new BufferedImage(pageWidth, pageHeight, BufferedImage.TYPE_INT_RGB);

		return (Graphics2D) bi.getGraphics();
	}
});
```

*Note:* A real `GraphicsBuilder` implementation (see examples below) will need to keep a reference to the `Graphics2D` object for later displaying, saving, etc

## Examples
* [DOCX to PDF](examples/docx-to-pdf)
* [DOCX to SVG](examples/docx-to-image)
* [DOCX to Bitmap Images](examples/docx-to-image)

## FAQ
**Why is the output so huge?**

The output is in OOXML's native units ([twips](http://en.wikipedia.org/wiki/Twip)).
The size can be reduced by setting the scale of the `Graphics2D` object:

```java
public class ScaledGraphicsBuilder implements GraphicsBuilder {
	@Override
	public Graphics2D nextPage(int pageWidth, int pageHeight) {
		Graphics2D g2 = ...;
		g2.scale(0.05, 0.05); // twips to actual size (72 DPI)
		return g2;t
	}
}
```