package documentconverter.reader;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.util.List;

import org.junit.Before;
import org.junit.Test;

import documentconverter.renderer.MockPage;
import documentconverter.renderer.MockRenderer;

public class DocxReaderTest {
	private static final File TEST_FILE_LAYOUTS = new File("src/test/resources/reader/docx/layouts.docx");
	private MockRenderer renderer;

	@Before
	public void setUp() {
		renderer = new MockRenderer();
	}

	@Test
	public void testLayouts() throws ReaderException {
		new DocxReader(renderer, TEST_FILE_LAYOUTS).process();

		List<MockPage> pages = renderer.getPages();

		assertEquals(3, pages.size());

		MockPage page1 = pages.get(0);
		MockPage page2 = pages.get(1);
		MockPage page3 = pages.get(2);

		assertEquals(12240, page1.getPageWidth());
		assertEquals(15840, page1.getPageHeight());

		assertEquals(12983, page2.getPageWidth());
		assertEquals(6463, page2.getPageHeight());

		assertEquals(15840, page3.getPageWidth());
		assertEquals(12240, page3.getPageHeight());
	}
}
