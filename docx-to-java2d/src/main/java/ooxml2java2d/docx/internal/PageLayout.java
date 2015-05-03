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

package ooxml2java2d.docx.internal;

import org.apache.commons.lang.builder.ToStringBuilder;
import org.apache.commons.lang.builder.ToStringStyle;
import org.docx4j.model.structure.HeaderFooterPolicy;

public class PageLayout {
	private int width;
	private int height;
	private int topMargin;
	private int rightMargin;
	private int bottomMargin;
	private int leftMargin;
	private int headerMargin;
	private int footerMargin;
	private HeaderFooterPolicy headerFooterPolicy;

	public PageLayout(
		int width, int height, int topMargin, int rightMargin, int bottomMargin,
		int leftMargin, int headerMargin, int footerMargin, HeaderFooterPolicy headerFooterPolicy
	) {
		this.width = width;
		this.height = height;
		this.topMargin = topMargin;
		this.rightMargin = rightMargin;
		this.bottomMargin = bottomMargin;
		this.leftMargin = leftMargin;
		this.headerMargin = headerMargin;
		this.footerMargin = footerMargin;
		this.headerFooterPolicy = headerFooterPolicy;
	}

	public int getWidth() {
		return width;
	}

	public int getHeight() {
		return height;
	}

	public int getTopMargin() {
		return topMargin;
	}

	public int getRightMargin() {
		return rightMargin;
	}

	public int getBottomMargin() {
		return bottomMargin;
	}

	public int getLeftMargin() {
		return leftMargin;
	}

	public int getHeaderMargin() {
		return headerMargin;
	}

	public int getFooterMargin() {
		return footerMargin;
	}

	public HeaderFooterPolicy getHeaderFooterPolicy() {
		return headerFooterPolicy;
	}

	@Override
	public String toString() {
		return new ToStringBuilder(this, ToStringStyle.SHORT_PREFIX_STYLE)
			.append("width", width)
			.append("height", height)
			.append("topMargin", topMargin)
			.append("rightMargin", rightMargin)
			.append("bottomMargin", bottomMargin)
			.append("leftMargin", leftMargin)
			.append("headerMargin", headerMargin)
			.append("footerMargin", footerMargin)
			.toString();
	}
}
