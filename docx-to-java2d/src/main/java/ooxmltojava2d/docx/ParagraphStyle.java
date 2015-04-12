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

import java.awt.Color;
import java.awt.geom.Rectangle2D;

import org.apache.commons.lang.builder.ToStringBuilder;
import org.apache.commons.lang.builder.ToStringStyle;

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

	@Override
	public String toString() {
		return new ToStringBuilder(this, ToStringStyle.SHORT_PREFIX_STYLE)
			.append("lineSpacing", lineSpacing)
			.append("spaceBefore", spaceBefore)
			.append("spaceAfter", spaceAfter)
			.append("color", color)
			.append("fontConfig", fontConfig)
			.append("alignment", alignment)
			.toString();
	}
}
