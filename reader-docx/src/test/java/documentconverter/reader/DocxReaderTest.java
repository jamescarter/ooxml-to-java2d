package documentconverter.reader;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.awt.Color;
import java.io.File;
import java.util.List;
import java.util.Set;
import org.junit.Before;
import org.junit.Test;

import documentconverter.renderer.ColorAction;
import documentconverter.renderer.DrawStringAction;
import documentconverter.renderer.FontStyle;
import documentconverter.renderer.MockPage;
import documentconverter.renderer.MockRenderer;

public class DocxReaderTest {
	private static final File TEST_FILE_LAYOUTS = new File("src/test/resources/reader/docx/layouts.docx");
	private static final File TEST_FILE_BODY_START_NO_HEADER = new File("src/test/resources/reader/docx/body_start_no_header.docx");
	private static final File TEST_FONT_SIZE = new File("src/test/resources/reader/docx/font_size.docx");
	private static final File TEST_FONT_STYLE = new File("src/test/resources/reader/docx/font_style.docx");
	private static final File TEST_LINE_HEIGHT = new File("src/test/resources/reader/docx/line_height.docx");
	private static final File TEST_TEXT_COLOR = new File("src/test/resources/reader/docx/text_color.docx");
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

		assertEquals(16320, page1.getPageWidth());
		assertEquals(21120, page1.getPageHeight());

		assertEquals(17310, page2.getPageWidth());
		assertEquals(8617, page2.getPageHeight());

		assertEquals(21120, page3.getPageWidth());
		assertEquals(16320, page3.getPageHeight());
	}

	@Test
	public void testBodyStartNoHeader() throws ReaderException {
		new DocxReader(renderer, TEST_FILE_BODY_START_NO_HEADER).process();

		List<MockPage> pages = renderer.getPages();

		DrawStringAction page1Action = pages.get(0).getActions(DrawStringAction.class).get(0);
		DrawStringAction page2Action = pages.get(1).getActions(DrawStringAction.class).get(0);

		assertEquals("DrawStringAction[text=Hello, World!,x=1920,y=2277]", page1Action.toString());
		assertEquals("DrawStringAction[text=Hello, World!,x=1512,y=2625]", page2Action.toString());
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

		int x = actions.get(0).getX();

		assertEquals(x, actions.get(0).getX());
		assertEquals(x += 632, actions.get(2).getX());
		assertEquals(x += 951, actions.get(4).getX());
		assertEquals(x += 1243, actions.get(6).getX());
		assertEquals(x += 721, actions.get(8).getX());
		assertEquals(x += 1502, actions.get(10).getX());
		assertEquals(x += 1222, actions.get(12).getX());
		assertEquals(x += 970, actions.get(14).getX());
		assertEquals(x += 1313, actions.get(16).getX());
		assertEquals(x += 1246, actions.get(17).getX());
	}

	@Test
	public void testFontStyle() throws ReaderException {
		new DocxReader(renderer, TEST_FONT_STYLE).process();

		List<FontConfigAction> actions = renderer.getPages().get(0).getActions(FontConfigAction.class);

		assertSet(actions.get(2).getStyles(), FontStyle.BOLD);
		assertSet(actions.get(4).getStyles(), FontStyle.ITALIC);
		assertSet(actions.get(6).getStyles(), FontStyle.UNDERLINE);
		assertSet(actions.get(8).getStyles(), FontStyle.STRIKETHROUGH);
		assertSet(actions.get(10).getStyles(), FontStyle.SUPERSCRIPT);
		assertSet(actions.get(13).getStyles(), FontStyle.BOLD, FontStyle.ITALIC, FontStyle.UNDERLINE, FontStyle.SUPERSCRIPT);
	}

	@Test
	public void testLineHeight() throws ReaderException {
		new DocxReader(renderer, TEST_LINE_HEIGHT).process();

		List<DrawStringAction> actions = renderer.getPages().get(0).getActions(DrawStringAction.class);

		assertEquals("This is the first paragraph.", actions.get(0).getText());
		assertEquals("This is the second paragraph.", actions.get(1).getText());
		assertEquals("This is the third paragraph.", actions.get(2).getText());
		assertEquals("This is the fourth paragraph.", actions.get(3).getText());

		// TODO: Line spacing between paragraphs is not perfect, needs further checking
		int y = actions.get(0).getY();

		assertEquals(y += 633, actions.get(1).getY());
		assertEquals(y += 878, actions.get(2).getY());
		assertEquals(y += 1124, actions.get(3).getY());
	}

	@Test
	public void testTextColor() throws ReaderException {
		new DocxReader(renderer, TEST_TEXT_COLOR).process();

		List<Object> actions = renderer.getPages().get(0).getActions(ColorAction.class, DrawStringAction.class);

		assertEquals(Color.RED, ((ColorAction)actions.get(0)).getColor());
		assertEquals("Red", ((DrawStringAction)actions.get(1)).getText());

		assertEquals(Color.GREEN, ((ColorAction)actions.get(4)).getColor());
		assertEquals("green", ((DrawStringAction)actions.get(5)).getText());

		assertEquals(Color.BLUE, ((ColorAction)actions.get(8)).getColor());
		assertEquals("blue", ((DrawStringAction)actions.get(9)).getText());
	}

	private void assertSet(Set<FontStyle> actualStyles, FontStyle ... expectedStyles) {
		for (FontStyle expected : expectedStyles) {
			assertTrue(actualStyles.contains(expected));
		}

		assertEquals(actualStyles.size(), expectedStyles.length);
	}
}
