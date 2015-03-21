package documentconverter.renderer;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class MockPage implements Page {
	private List<String> actions = new ArrayList<>();
	private int pageWidth;
	private int pageHeight;
	private FontConfig fontConfig;

	public MockPage(int pageWidth, int pageHeight) {
		this.pageWidth = pageWidth;
		this.pageHeight = pageHeight;
	}

	@Override
	public void setFontConfig(FontConfig fontConfig) {
		this.fontConfig = fontConfig;
	}

	@Override
	public void drawString(String text, int x, int y) {
		actions.add("text: "+ text + ", x: " + x + ", y: " + y);
	}

	public int getPageWidth() {
		return pageWidth;
	}

	public int getPageHeight() {
		return pageHeight;
	}

	public FontConfig getFontConfig() {
		return fontConfig;
	}

	public List<String> getActions() {
		return Collections.unmodifiableList(actions);
	}
}
