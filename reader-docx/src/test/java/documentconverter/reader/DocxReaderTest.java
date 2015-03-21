package documentconverter.reader;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.util.List;

import org.junit.Before;
import org.junit.Test;

import documentconverter.renderer.DrawStringAction;
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

		DrawStringAction page1Action = pages.get(0).getActions(DrawStringAction.class).get(0);
		DrawStringAction page2Action = pages.get(1).getActions(DrawStringAction.class).get(0);

		assertEquals("DrawStringAction[text=Hello, World!,x=1440,y=1440]", page1Action.toString());
		assertEquals("DrawStringAction[text=Hello, World!,x=1701,y=1134]", page2Action.toString());
	}

	@Test
	public void testFontSize() throws ReaderException {
		new DocxReader(renderer, TEST_FONT_SIZE).process();

		List<DrawStringAction> actions = renderer.getPages().get(0).getActions(DrawStringAction.class);

		assertEquals("The", actions.get(0).getText());
		assertEquals("quick", actions.get(2).getText());
		assertEquals("brown", actions.get(4).getText());
		assertEquals("fox", actions.get(6).getText());
		assertEquals("jumps", actions.get(8).getText());
		assertEquals("over", actions.get(10).getText());
		assertEquals("the", actions.get(12).getText());
		assertEquals("lazy", actions.get(14).getText());
		assertEquals("dog", actions.get(16).getText());
		assertEquals(".", actions.get(17).getText());

		assertEquals(1440, actions.get(0).getX());
		assertEquals(2079, actions.get(2).getX());
		assertEquals(3052, actions.get(4).getX());
		assertEquals(4332, actions.get(6).getX());
		assertEquals(5105, actions.get(8).getX());
		assertEquals(6674, actions.get(10).getX());
		assertEquals(7978, actions.get(12).getX());
		assertEquals(9044, actions.get(14).getX());
		assertEquals(10468, actions.get(16).getX());
		assertEquals(11714, actions.get(17).getX());
	}
}
