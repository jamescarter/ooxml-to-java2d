package documentconverter.reader;

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

	public boolean canFitContent(double newContentWidth) {
		return contentWidth + newContentWidth <= width;
	}

	public void addContent(double newContentWidth, double newContentHeight, Object newContent) {
		if (!canFitContent(newContentWidth)) {
			throw new ContentTooBigException("Content too big for current line");
		}

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
			.toString();
	}
}
