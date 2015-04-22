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

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.commons.lang.builder.ToStringBuilder;
import org.apache.commons.lang.builder.ToStringStyle;

/**
 * Columns represent an area of a page with contents that do not exceed the specified width.
 */
public class Column {
	private int xOffset;
	private int width;
	private int contentWidth;
	private int contentHeight;
	private List<Object> actions = new ArrayList<>();
	private boolean isCachedOverPageFold;

	public Column(int xOffset, int width) {
		this.width = width;
		this.xOffset = xOffset;
	}

	public int getContentWidth() {
		return contentWidth;
	}

	public int getContentHeight() {
		return contentHeight;
	}

	public int getWidth() {
		return width;
	}

	public int getXOffset() {
		return xOffset;
	}

	public List<Object> getActions() {
		return Collections.unmodifiableList(actions);
	}

	public boolean hasContent() {
		return actions.size() > 0;
	}

	public boolean canFitContent(double newContentWidth) {
		return contentWidth + newContentWidth <= width;
	}

	public void addContent(double newContentWidth, double newContentHeight, Object newContent) {
		if (!canFitContent(newContentWidth)) {
			throw new ContentTooBigException("Content too big for current line");
		}

		addContentForced(newContentWidth, newContentHeight, newContent);
	}

	public void addContentForced(double newContentWidth, double newContentHeight, Object newContent) {
		contentWidth += newContentWidth;
		contentHeight = (int) Math.max(contentHeight, newContentHeight);
		actions.add(newContent);
	}

	public void addContentOffset(int offset) {
		if (!canFitContent(offset)) {
			throw new ContentTooBigException("Offset too big for current line");
		}

		contentWidth += offset;
	}

	public void addAction(Object action) {
		actions.add(action);
	}

	public boolean isCachedOverPageFold() {
		return isCachedOverPageFold;
	}

	/**
	 * Sets whether the content should be temporarily cached across pages folds
	 * @param isCachedOverPageFold if the content should be cached across page folds
	 */
	public void setCacheOverPageFold(boolean isCachedOverPageFold) {
		this.isCachedOverPageFold = isCachedOverPageFold;
	}

	public void reset() {
		contentWidth = 0;
		contentHeight = 0;
		actions.clear();
	}

	@Override
	public String toString() {
		return new ToStringBuilder(this, ToStringStyle.SHORT_PREFIX_STYLE)
			.append("xOffset", xOffset)
			.append("width", width)
			.append("contentWidth", contentWidth)
			.append("contentHeight", contentHeight)
			.append("actions", actions)
			.append("isCachedOverPageFold", isCachedOverPageFold)
			.toString();
	}
}
