package documentconverter.renderer;

import java.awt.Color;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

public class MockPage implements Page {
	private List<Object> actions = new ArrayList<>();
	private int pageWidth;
	private int pageHeight;

	public MockPage(int pageWidth, int pageHeight) {
		this.pageWidth = pageWidth;
		this.pageHeight = pageHeight;
	}

	@Override
	public void setColor(Color color) {
		actions.add(color);
	}

	@Override
	public void setFontConfig(FontConfig fontConfig) {
		actions.add(fontConfig);
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

	public List<Object> getActions() {
		return Collections.unmodifiableList(actions);
	}

	public List<Object> getActions(Class<?> ... classes) {
		List<Object> as = new ArrayList<>();

		for (Object action : actions) {
			for (Class<?> clazz : classes) {
				if (clazz.isInstance(action)) {
					as.add(action);
					break;
				}
			}
		}

		return as;
	}

	public <T> List<T> getActions(Class<T> clazz) {
		return (List<T>) actions.stream().filter(o -> clazz.isInstance(o)).collect(Collectors.toList());
	}
}
