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

package ooxml2java2d.docx.internal.content;

import java.awt.Color;
import org.apache.commons.lang.builder.ToStringBuilder;
import org.apache.commons.lang.builder.ToStringStyle;

public class Border {
	private Color color;
	private int size;
	private BorderStyle style;

	public Border(Color color, int size, BorderStyle style) {
		this.color = color;
		this.size = size;
		this.style = style;
	}

	public Color getColor() {
		return color;
	}

	public int getSize() {
		return size;
	}

	public BorderStyle getStyle() {
		return style;
	}

	@Override
	public String toString() {
		return new ToStringBuilder(this, ToStringStyle.SHORT_PREFIX_STYLE)
			.append("color", color)
			.append("size", size)
			.append("style", style)
			.toString();
	}
}
