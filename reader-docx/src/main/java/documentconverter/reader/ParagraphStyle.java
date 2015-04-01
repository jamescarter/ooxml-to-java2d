package documentconverter.reader;

import java.awt.Color;
import java.awt.geom.Rectangle2D;

import documentconverter.renderer.FontConfig;
import documentconverter.renderer.FontStyle;

public class ParagraphStyle {
	private int lineSpacing;
	private int spaceBefore;
	private int spaceAfter;
	private Color color = Color.BLACK;
	private FontConfig fontConfig = new FontConfig();
	private Alignment alignment = Alignment.LEFT;

	public ParagraphStyle() { }

	public ParagraphStyle(ParagraphStyle baseStyle) {
		this.lineSpacing = baseStyle.getLineSpacing();
		this.spaceBefore = baseStyle.getSpaceBefore();
		this.spaceAfter = baseStyle.getSpaceAfter();
		this.color = baseStyle.getColor();
		this.fontConfig = new FontConfig(baseStyle.getFontConfig());
		this.alignment = baseStyle.getAlignment();
	}

	public void setLineSpacing(int lineSpacing) {
		this.lineSpacing = lineSpacing;
	}

	public void setSpaceBefore(int spaceBefore) {
		this.spaceBefore = spaceBefore;
	}

	public void setSpaceAfter(int spaceAfter) {
		this.spaceAfter = spaceAfter;
	}

	public void setColor(Color color) {
		this.color = color;
	}

	public void setFontName(String name) {
		fontConfig.setName(name);
	}

	public void setFontSize(float size) {
		fontConfig.setSize(size);
	}

	public void setAlignment(Alignment alignment) {
		this.alignment = alignment;
	}

	public void enableFontStyle(FontStyle style) {
		fontConfig.enableStyle(style);
	}

	public int getLineSpacing() {
		return lineSpacing;
	}

	public int getSpaceBefore() {
		return spaceBefore;
	}

	public int getSpaceAfter() {
		return spaceAfter;
	}

	public Color getColor() {
		return color;
	}

	public FontConfig getFontConfig() {
		return fontConfig;
	}

	public Alignment getAlignment() {
		return alignment;
	}

	public Rectangle2D getStringBoxSize(String text) {
		return fontConfig.getStringBoxSize(text);
	}
}
