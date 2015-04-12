/*
 * Copyright (C) 2015 James Carter
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *      http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package ooxmltojava2d.docx;

import java.awt.Font;
import java.awt.font.FontRenderContext;
import java.awt.font.TextAttribute;
import java.awt.geom.Rectangle2D;
import java.util.Collections;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.apache.commons.lang.builder.EqualsBuilder;
import org.apache.commons.lang.builder.HashCodeBuilder;

public class FontConfig {
	private Font font = new Font(Font.SERIF, Font.PLAIN, 1);
	private String name = font.getName();
	private float size;
	private Set<FontStyle> styles = new HashSet<>();

	public FontConfig() { }

	public FontConfig(FontConfig fontConfig) {
		setName(fontConfig.getName());
		setSize(fontConfig.getSize());

		for (FontStyle fs : fontConfig.getStyles()) {
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

	@Override
	public int hashCode() {
		return new HashCodeBuilder(31, 16)
			.append(name)
			.append(size)
			.append(styles)
			.toHashCode();
	}

	@Override
	public boolean equals(Object obj) {
		if (obj == null) {
			return false;
		} else if (obj == this) {
			return true;
		} else if (obj.getClass() != getClass()) {
			return false;
		}

		FontConfig fc = (FontConfig) obj;

		return new EqualsBuilder()
			.append(name, fc.name)
			.append(size, fc.size)
			.append(styles, fc.styles)
			.isEquals();
	}
}