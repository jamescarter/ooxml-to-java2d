package ooxmltojava2d.docx;

import org.apache.commons.lang.builder.ToStringBuilder;
import org.apache.commons.lang.builder.ToStringStyle;

public class DrawStringAction {
	private String text;
	private int x;
	private int y;

	public DrawStringAction(String text, int x, int y) {
		this.text = text;
		this.x = x;
		this.y = y;
	}

	public String getText() {
		return text;
	}

	public int getX() {
		return x;
	}

	public int getY() {
		return y;
	}

	@Override
	public String toString() {
		return new ToStringBuilder(this, ToStringStyle.SHORT_PREFIX_STYLE)
			.append("text", text)
			.append("x", x)
			.append("y", y)
			.toString();
	}
}
