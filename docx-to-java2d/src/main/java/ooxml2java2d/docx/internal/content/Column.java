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
import java.util.List;

import org.apache.commons.lang.builder.ToStringBuilder;
import org.apache.commons.lang.builder.ToStringStyle;

/**
 * Columns represent an area of a page with contents that do not exceed the specified width.
 */
public class Column {
	private int xOffset;
	private int width;
	private boolean isCachedOverPageFold;
	private List<Row> rows = new ArrayList<>();
	private Line line;

	public Column(int xOffset, int width) {
		this.xOffset = xOffset;
		this.width = width;
	}

	public int getXOffset() {
		return xOffset;
	}

	public int getWidth() {
		return width;
	}

	public int getContentHeight() {
		int height = 0;

		for (Row row : rows) {
			height += row.getContentHeight();
		}

		return height;
	}

	public Line getCurrentLine() {
		if (line == null) {
			line = new Line(width);
			rows.add(line);
		}

		return line;
	}

	public Row[] getRows() {
		return rows.toArray(new Row[rows.size()]);
	}

	/**
	 * Sets whether the content should be temporarily cached across pages folds
	 * @param isCachedOverPageFold if the content should be cached across page folds
	 */
	public void setCacheOverPageFold(boolean isCachedOverPageFold) {
		this.isCachedOverPageFold = isCachedOverPageFold;
	}

	public void addVerticalSpace(int height) {
		rows.add(new BlankRow(height));
		line = null;
	}

	public void addHorizontalSpace(int width, int verticalSpace) {
		addContent(new Content(width, 0), verticalSpace);
	}

	public void addContent(Content content, int verticalSpace) {
		Line currentLine = getCurrentLine();

		if (!currentLine.canFitContent(content.getWidth())) {
			addVerticalSpace(verticalSpace);
		}

		currentLine.addContent(content);
	}

	public void addContentForced(Content content) {
		getCurrentLine().addContentForced(content);
	}

	public void addTableRow(TableRow tableRow) {
		rows.add(tableRow);
	}

	public void addAction(Object action) {
		getCurrentLine().addAction(action);
	}

	public void removeRow(Row row) {
		rows.remove(row);

		if (rows.isEmpty()) {
			line = null;
		}
	}

	public boolean isCachedOverPageFold() {
		return isCachedOverPageFold;
	}

	public boolean isEmpty() {
		return rows.isEmpty();
	}

	@Override
	public String toString() {
		return new ToStringBuilder(this, ToStringStyle.SHORT_PREFIX_STYLE)
			.append("xOffset", xOffset)
			.append("width", width)
			.append("isCachedOverPageFold", isCachedOverPageFold)
			.append("rows", rows)
			.toString();
	}
}