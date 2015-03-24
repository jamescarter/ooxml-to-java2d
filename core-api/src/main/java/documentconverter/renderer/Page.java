package documentconverter.renderer;

import java.awt.Color;

public interface Page {
	void setColor(Color color);
	void setFontConfig(FontConfig fontConfig);
	void drawString(String text, int x, int y);
}
