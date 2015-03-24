package documentconverter.renderer;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;

import documentconverter.renderer.Page;

public class Image implements Page {
	private static final float scale = 0.05f;
	private BufferedImage bi;
	private Graphics2D g;

	public Image(int pageWidth, int pageHeight) {
		bi = new BufferedImage((int) (pageWidth * scale), (int) (pageHeight * scale), BufferedImage.TYPE_INT_RGB);
		g = (Graphics2D) bi.getGraphics();

		g.setColor(Color.WHITE);
		g.fillRect(0, 0, bi.getWidth(), bi.getHeight());
		g.setColor(Color.BLACK);

		g.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
		g.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_LCD_HRGB);
		g.setRenderingHint(RenderingHints.KEY_TEXT_LCD_CONTRAST, 250);
	}

	@Override
	public void setColor(Color color) {
		g.setColor(color);
	}

	@Override
	public void setFontConfig(FontConfig fontConfig) {
	    g.setFont(fontConfig.getFont().deriveFont(fontConfig.getSize() * scale));
	}

	@Override
	public void drawString(String text, int x, int y) {
		g.drawString(text, (int) (x * scale), (int) (y * scale));
	}

	public BufferedImage getImage() {
		return bi;
	}
}
