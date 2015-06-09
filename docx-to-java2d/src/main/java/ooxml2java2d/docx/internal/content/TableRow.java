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

import java.util.List;

import org.apache.commons.lang.builder.ToStringBuilder;
import org.apache.commons.lang.builder.ToStringStyle;

public class TableRow implements Row {
	private int minHeight;
	private List<Column> columns;

	public TableRow(int minHeight, List<Column> columns) {
		this.minHeight = minHeight;
		this.columns = columns;
	}

	public List<Column> getColumns() {
		return columns;
	}

	@Override
	public int getContentHeight() {
		int maxHeight = minHeight;

		for (Column column : columns) {
			maxHeight = Math.max(maxHeight, column.getContentHeight());
		}

		return maxHeight;
	}

	public boolean isEmpty() {
		for (Column column : columns) {
			if (!column.isEmpty()) {
				return false;
			}
		}

		return true;
	}

	@Override
	public String toString() {
		return new ToStringBuilder(this, ToStringStyle.SHORT_PREFIX_STYLE)
			.append("contentHeight", getContentHeight())
			.append("columns", columns)
			.toString();
	}
}
