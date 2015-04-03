package documentconverter.renderer;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import documentconverter.renderer.Page;

public class Image implements Page {
	private Graphics2D g;

	public Image(Graphics2D g) {
		this.g = g;

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
	    g.setFont(fontConfig.getFont());
	}

	@Override
	public void drawString(String text, int x, int y) {
		g.drawString(text, x, y);
	}

	public Graphics2D getGraphics() {
		return g;
	}
}
