package documentconverter.renderer;

import java.awt.Font;
import java.awt.font.FontRenderContext;
import java.awt.font.TextAttribute;
import java.awt.geom.Rectangle2D;
import java.util.Collections;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

public class FontConfig {
	private Font font = new Font(Font.SERIF, Font.PLAIN, 1);
	private String name = font.getName();
	private float size;
	private Set<FontStyle> styles = new HashSet<>();

	public FontConfig() { }

	public FontConfig(String name, float size, Set<FontStyle> styles) {
		setName(name);
		setSize(size);

		for (FontStyle fs : styles) {
			enableStyle(fs);
		}
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

	public void enableStyle(FontStyle style) {
		if (!hasStyle(style)) {
			styles.add(style);

			switch(style) {
				case BOLD:
					font = font.deriveFont(Font.BOLD);
				break;
				case ITALIC:
					font = font.deriveFont(Font.ITALIC);
				break;
				case STRIKETHROUGH:
					font = setAttribute(TextAttribute.STRIKETHROUGH, TextAttribute.STRIKETHROUGH_ON);
				break;
				case SUPERSCRIPT:
					font = setAttribute(TextAttribute.SUPERSCRIPT, TextAttribute.SUPERSCRIPT_SUPER);
				break;
				case UNDERLINE:
					font = setAttribute(TextAttribute.UNDERLINE, TextAttribute.UNDERLINE_ON);
				break;
			}	
		}
	}

	public String getName() {
		return name;
	}

	public float getSize() {
		return size;
	}

	public Set<FontStyle> getStyles() {
		return Collections.unmodifiableSet(styles);
	}

	public Rectangle2D getStringBoxSize(String text) {
		return font.getStringBounds(text, new FontRenderContext(font.getTransform(), true, true));
	}

	public Font getFont() {
		return font;
	}

	public boolean hasStyle(FontStyle style) {
		return styles.contains(style);
	}

	private Font setAttribute(TextAttribute attribute, Object value) {
		Map attributes = font.getAttributes();

		attributes.put(attribute, value);

		return font.deriveFont(attributes);
	}
} 
