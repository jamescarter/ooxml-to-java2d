package documentconverter.renderer;

import java.awt.Font;
import java.awt.font.FontRenderContext;
import java.awt.geom.Rectangle2D;

public class FontConfig {
	private Font font = new Font(Font.SERIF, Font.PLAIN, 1);
	private String name = font.getName();
	private float size;

	public FontConfig() {

	}

	public FontConfig(String name, float size) {
		setName(name);
		setSize(size);
	}

	public void setName(String name) {
		if (!font.getName().equals(name)) {
			this.name = name;
			this.font = new Font(name, font.getStyle(), font.getSize());
			setSize(size);
		}
	}

	public void setSize(float size) {
		this.size = size;
		this.font = font.deriveFont(size);
	}

	public String getName() {
		return name;
	}

	public float getSize() {
		return size;
	}

	public Rectangle2D getStringBoxSize(String text) {
		return font.getStringBounds(text, new FontRenderContext(font.getTransform(), true, true));
	}

	public Font getFont() {
		return font;
	}
} 
