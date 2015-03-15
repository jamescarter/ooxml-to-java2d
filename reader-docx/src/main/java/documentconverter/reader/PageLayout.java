package documentconverter.reader;

public class PageLayout {
	private int width;
	private int height;

	public PageLayout(int width, int height) {
		this.width = width;
		this.height = height;
	}

	public int getWidth() {
		return width;
	}

	public int getHeight() {
		return height;
	}
}
