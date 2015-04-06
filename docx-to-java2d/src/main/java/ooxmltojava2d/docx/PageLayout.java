package ooxmltojava2d.docx;

import org.apache.commons.lang.builder.ToStringBuilder;
import org.apache.commons.lang.builder.ToStringStyle;

public class PageLayout {
	private int width;
	private int height;
	private int topMargin;
	private int rightMargin;
	private int bottomMargin;
	private int leftMargin;

	public PageLayout(int width, int height, int topMargin, int rightMargin, int bottomMargin, int leftMargin) {
		this.width = width;
		this.height = height;
		this.topMargin = topMargin;
		this.rightMargin = rightMargin;
		this.bottomMargin = bottomMargin;
		this.leftMargin = leftMargin;
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

	@Override
	public String toString() {
		return new ToStringBuilder(this, ToStringStyle.SHORT_PREFIX_STYLE)
			.append("width", width)
			.append("height", height)
			.append("topMargin", topMargin)
			.append("rightMargin", rightMargin)
			.append("bottomMargin", bottomMargin)
			.append("leftMargin", leftMargin)
			.toString();
	}
}
