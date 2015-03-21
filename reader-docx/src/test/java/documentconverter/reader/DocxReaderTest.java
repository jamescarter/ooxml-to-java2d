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
	private static final File TEST_FILE_BODY_START_NO_HEADER = new File("src/test/resources/reader/docx/body_start_no_header.docx");
	private static final File TEST_FONT_SIZE = new File("src/test/resources/reader/docx/font_size.docx");
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

	@Test
	public void testBodyStartNoHeader() throws ReaderException {
		new DocxReader(renderer, TEST_FILE_BODY_START_NO_HEADER).process();

		List<MockPage> pages = renderer.getPages();

		assertEquals("text: Hello, World!, x: 1440, y: 1440", pages.get(0).getActions().get(0));
		assertEquals("text: Hello, World!, x: 1701, y: 1134", pages.get(1).getActions().get(0));
	}

	@Test
	public void testFontSize() throws ReaderException {
		new DocxReader(renderer, TEST_FONT_SIZE).process();

		List<String> actions = renderer.getPages().get(0).getActions();

		assertEquals("text: The, x: 1440, y: 1440", actions.get(0));
		assertEquals("text: quick, x: 2079, y: 1440", actions.get(2));
		assertEquals("text: brown, x: 3052, y: 1440", actions.get(4));
		assertEquals("text: fox, x: 4332, y: 1440", actions.get(6));
		assertEquals("text: jumps, x: 5105, y: 1440", actions.get(8));
		assertEquals("text: over, x: 6674, y: 1440", actions.get(10));
		assertEquals("text: the, x: 7978, y: 1440", actions.get(12));
		assertEquals("text: lazy, x: 9044, y: 1440", actions.get(14));
		assertEquals("text: dog, x: 10468, y: 1440", actions.get(16));
		assertEquals("text: ., x: 11714, y: 1440", actions.get(17));
	}
}
