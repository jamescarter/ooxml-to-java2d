package documentconverter.renderer;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class MockRenderer implements Renderer {
	private List<MockPage> pages = new ArrayList<>();

	@Override
	public Page addPage(int pageWidth, int pageHeight) {
		MockPage page = new MockPage(pageWidth, pageHeight);

		pages.add(page);

		return page;
	}

	public List<MockPage> getPages() {
		return Collections.unmodifiableList(pages);
	}
}
