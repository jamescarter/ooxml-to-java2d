package documentconverter.renderer;

public interface Page {
	void setFontConfig(FontConfig fontConfig);
	void drawString(String text, int x, int y);
}
