package documentconverter.renderer;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

import documentconverter.reader.FontConfigAction;

public class MockPage implements Page {
	private List<Object> actions = new ArrayList<>();
	private int pageWidth;
	private int pageHeight;

	public MockPage(int pageWidth, int pageHeight) {
		this.pageWidth = pageWidth;
		this.pageHeight = pageHeight;
	}

	@Override
	public void setFontConfig(FontConfig fontConfig) {
		actions.add(new FontConfigAction(fontConfig.getName(), fontConfig.getSize()));
	}

	@Override
	public void drawString(String text, int x, int y) {
		actions.add(new DrawStringAction(text, x, y));
	}

	public int getPageWidth() {
		return pageWidth;
	}

	public int getPageHeight() {
		return pageHeight;
	}

	public String getAction(int index) {
		return actions.get(index).toString();
	}

	public <T> List<T> getActions(Class<T> clazz) {
		return (List<T>) actions.stream().filter(o -> clazz.isInstance(o)).collect(Collectors.toList());
	}
}
