package ooxmltojava2d.docx;

import java.awt.Graphics2D;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class MockGraphicsBuilder implements GraphicsBuilder {
	private List<MockGraphics2D> graphics = new ArrayList<>();

	@Override
	public Graphics2D nextPage(int pageWidth, int pageHeight) {
		MockGraphics2D g = new MockGraphics2D(pageWidth, pageHeight);

		graphics.add(g);

		return g;
	}

	public List<MockGraphics2D> getPages() {
		return Collections.unmodifiableList(graphics);
	}
}
