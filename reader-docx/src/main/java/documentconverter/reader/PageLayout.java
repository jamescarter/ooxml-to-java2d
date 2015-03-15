package documentconverter.reader;

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
}
