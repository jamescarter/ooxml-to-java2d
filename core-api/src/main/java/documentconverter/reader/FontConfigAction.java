package documentconverter.reader;

import java.util.Set;

import documentconverter.renderer.FontStyle;

public class FontConfigAction {
	private String name;
	private float size;
	private Set<FontStyle> styles;

	public FontConfigAction(String name, float size, Set<FontStyle> styles) {
		this.name = name;
		this.size = size;
		this.styles = styles;
	}

	public String getName() {
		return name;
	}

	public float getSize() {
		return size;
	}

	public Set<FontStyle> getStyles() {
		return styles;
	}
}
