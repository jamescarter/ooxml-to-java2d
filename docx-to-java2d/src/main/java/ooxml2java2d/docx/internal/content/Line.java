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

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.commons.lang.builder.ToStringBuilder;
import org.apache.commons.lang.builder.ToStringStyle;

public class Line implements Row {
	private int width;
	private int contentWidth;
	private int contentHeight;
	private List<Object> actions = new ArrayList<>();

	public Line(int width) {
		this.width = width;
	}

	public int getWidth() {
		return width;
	}

	public int getContentWidth() {
		return contentWidth;
	}

	public int getContentHeight() {
		return contentHeight;
	}

	public List<Object> getActions() {
		return Collections.unmodifiableList(actions);
	}

	protected void addContent(Content content) {
		if (!canFitContent(content.getWidth())) {
			throw new ContentTooBigException("Content too big for line");
		}

		addContentForced(content);
	}

	protected void addContentForced(Content content) {
		contentWidth += content.getWidth();
		contentHeight = (int) Math.max(contentHeight, content.getHeight());
		actions.add(content);		
	}

	protected void addAction(Object action) {
		actions.add(action);
	}

	public boolean canFitContent(double newContentWidth) {
		return contentWidth + newContentWidth <= width;
	}

	@Override
	public String toString() {
		return new ToStringBuilder(this, ToStringStyle.SHORT_PREFIX_STYLE)
			.append("width", width)
			.append("contentWidth", contentWidth)
			.append("contentHeight", contentHeight)
			.append("actions", actions)
			.toString();
	}
}
