package documentconverter.renderer;

import java.awt.Color;
import java.awt.Graphics;
import java.awt.image.BufferedImage;

import documentconverter.renderer.Page;

public class Image implements Page {
	private static final float scale = 0.05f;
	private BufferedImage bi;
	private Graphics g;

	public Image(int pageWidth, int pageHeight) {
		bi = new BufferedImage((int) (pageWidth * scale), (int) (pageHeight * scale), BufferedImage.TYPE_INT_RGB);
		g = bi.getGraphics();
		g.setColor(Color.WHITE);
		g.fillRect(0, 0, (int) (pageWidth * scale), (int) (pageHeight * scale));
		g.setColor(Color.BLACK);
	}

	@Override
	public void drawString(String text, int x, int y) {
		g.drawString(text, (int) (x * scale), (int) (y * scale));
	}

	public BufferedImage getImage() {
		return bi;
	}
}
