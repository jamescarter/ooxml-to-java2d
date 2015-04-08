package ooxmltojava2d.docx;

public class DrawImageAction {
	private byte[] image;
	private int width;
	private int height;
	private int x;

	public DrawImageAction(byte[] image, int width, int height, int x) {
		this.image = image;
		this.width = width;
		this.height = height;
		this.x = x;
	}

	public byte[] getImage() {
		return image;
	}

	public int getX() {
		return x;
	}

	public int getWidth() {
		return width;
	}

	public int getHeight() {
		return height;
	}
}
