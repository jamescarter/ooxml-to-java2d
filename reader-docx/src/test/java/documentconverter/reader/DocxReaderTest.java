package documentconverter.reader;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.awt.Color;
import java.io.File;
import java.util.List;
import java.util.Set;

import org.junit.Before;
import org.junit.Test;

import documentconverter.renderer.DrawStringAction;
import documentconverter.renderer.FontConfig;
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
	private static final File TEST_HEADER_STYLE = new File("src/test/resources/reader/docx/header_style.docx");
	private static final File TEST_WORD_WRAP = new File("src/test/resources/reader/docx/word_wrap.docx");
	private static final File TEST_WORD_WRAP_CONTINUOUS = new File("src/test/resources/reader/docx/word_wrap_continuous.docx");
	private static final File TEST_TABBED = new File("src/test/resources/reader/docx/tabbed.docx");
	private static final File TEST_ALIGNMENT = new File("src/test/resources/reader/docx/alignment.docx");
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

		assertEquals("DrawStringAction[text=Hello, World!,x=1440,y=1708]", page1Action.toString());
		assertEquals("DrawStringAction[text=Hello, World!,x=1134,y=1969]", page2Action.toString());
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
		assertEquals(x += 474, actions.get(2).getX());
		assertEquals(x += 714, actions.get(4).getX());
		assertEquals(x += 932, actions.get(6).getX());
		assertEquals(x += 541, actions.get(8).getX());
		assertEquals(x += 1127, actions.get(10).getX());
		assertEquals(x += 916, actions.get(12).getX());
		assertEquals(x += 728, actions.get(14).getX());
		assertEquals(x += 985, actions.get(16).getX());
		assertEquals(x += 934, actions.get(17).getX());
	}

	@Test
	public void testFontStyle() throws ReaderException {
		new DocxReader(renderer, TEST_FONT_STYLE).process();

		List<FontConfig> actions = renderer.getPages().get(0).getActions(FontConfig.class);

		assertSet(actions.get(0).getStyles(), FontStyle.BOLD);
		assertSet(actions.get(2).getStyles(), FontStyle.ITALIC);
		assertSet(actions.get(4).getStyles(), FontStyle.UNDERLINE);
		assertSet(actions.get(6).getStyles(), FontStyle.STRIKETHROUGH);
		assertSet(actions.get(8).getStyles(), FontStyle.SUPERSCRIPT);
		assertSet(actions.get(10).getStyles(), FontStyle.BOLD, FontStyle.ITALIC, FontStyle.UNDERLINE, FontStyle.SUPERSCRIPT);
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

		assertEquals(y += 475, actions.get(1).getY());
		assertEquals(y += 659, actions.get(2).getY());
		assertEquals(y += 843, actions.get(3).getY());
	}

	@Test
	public void testTextColor() throws ReaderException {
		new DocxReader(renderer, TEST_TEXT_COLOR).process();

		List<Object> actions = renderer.getPages().get(0).getActions(Color.class, DrawStringAction.class);

		assertEquals(Color.RED, ((Color)actions.get(0)));
		assertEquals("Red", ((DrawStringAction)actions.get(1)).getText());

		assertEquals(Color.GREEN, ((Color)actions.get(4)));
		assertEquals("green", ((DrawStringAction)actions.get(5)).getText());

		assertEquals(Color.BLUE, ((Color)actions.get(8)));
		assertEquals("blue", ((DrawStringAction)actions.get(9)).getText());
	}

	@Test
	public void testHeaderStyle() throws ReaderException {
		new DocxReader(renderer, TEST_HEADER_STYLE).process();

		List<Object> actions = renderer.getPages().get(0).getActions(FontConfig.class, DrawStringAction.class);

		FontConfig fc1 = (FontConfig) (FontConfig) actions.get(0);
		FontConfig fc2 = (FontConfig) (FontConfig) actions.get(2);
		FontConfig fc3 = (FontConfig) (FontConfig) actions.get(4);
		FontConfig fc4 = (FontConfig) (FontConfig) actions.get(6);
		FontConfig fc5 = (FontConfig) (FontConfig) actions.get(8);

		assertEquals(560, (int) fc1.getSize());
		assertTrue(fc1.hasStyle(FontStyle.BOLD));
		assertEquals("Arial", fc1.getName());
		assertEquals("Title", ((DrawStringAction) actions.get(1)).getText());

		assertEquals(360, (int) fc2.getSize());
		assertEquals("Arial", fc2.getName());
		assertEquals("Subtitle", ((DrawStringAction) actions.get(3)).getText());

		assertEquals(360, (int) fc3.getSize());
		assertEquals("Arial", fc3.getName());
		assertTrue(fc3.hasStyle(FontStyle.BOLD));
		assertEquals("Header 1", ((DrawStringAction) actions.get(5)).getText());

		assertEquals(320, (int) fc4.getSize());
		assertEquals("Arial", fc4.getName());
		assertTrue(fc4.hasStyle(FontStyle.BOLD));
		assertEquals("Header 2", ((DrawStringAction) actions.get(7)).getText());

		assertEquals(280, (int) fc5.getSize());
		assertEquals("Arial", fc5.getName());
		assertTrue(fc5.hasStyle(FontStyle.BOLD));
		assertEquals("Header 3", ((DrawStringAction) actions.get(9)).getText());
	}

	@Test
	public void testWordWrap() throws ReaderException {
		new DocxReader(renderer, TEST_WORD_WRAP).process();

		List<DrawStringAction> actions = renderer.getPages().get(0).getActions(DrawStringAction.class);

		DrawStringAction line1 = actions.get(0);
		assertEquals("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam aliquet", line1.getText());

		DrawStringAction line2 = actions.get(1);
		assertEquals("vehicula magna, sed maximus tellus imperdiet sit amet. Maecenas eu", line2.getText());
		assertTrue(line2.getY() > line1.getY());

		DrawStringAction line3 = actions.get(2);
		assertEquals("maximus dolor. Phasellus tempor, enim non mattis porta, neque est", line3.getText());
		assertTrue(line3.getY() > line2.getY());

		DrawStringAction line4 = actions.get(3);
		assertEquals("elementum sapien, vel blandit elit turpis at orci. Sed in dolor nulla. Nunc", line4.getText());
		assertTrue(line4.getY() > line3.getY());

		DrawStringAction line5 = actions.get(4);
		assertEquals("aliquet enim eu orci finibus tincidunt. Fusce consequat blandit tellus, vel", line5.getText());
		assertTrue(line5.getY() > line4.getY());

		DrawStringAction line6 = actions.get(5);
		assertEquals("auctor dui dictum sollicitudin. Maecenas feugiat, augue vitae egestas", line6.getText());
		assertTrue(line6.getY() > line5.getY());

		DrawStringAction line7 = actions.get(6);
		assertEquals("iaculis, neque nunc sodales felis, non tempor lorem nisi nec arcu.", line7.getText());
		assertTrue(line7.getY() > line6.getY());
	}

	@Test
	public void testWordWrapContinuous() throws ReaderException {
		new DocxReader(renderer, TEST_WORD_WRAP_CONTINUOUS).process();

		List<DrawStringAction> actions = renderer.getPages().get(0).getActions(DrawStringAction.class);

		DrawStringAction line1 = actions.get(0);
		assertEquals("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", line1.getText());

		DrawStringAction line2 = actions.get(1);
		assertEquals("BBBBBBBBBBBBB CCCCCCCCCCCCC", line2.getText());
		assertTrue(line2.getY() > line1.getY());

		DrawStringAction line3 = actions.get(2);
		assertEquals("DDDDDDDDDDDD", line3.getText());
		assertTrue(line3.getY() > line2.getY());
	}

	@Test
	public void testTabbed() throws ReaderException {
		new DocxReader(renderer, TEST_TABBED).process();

		List<DrawStringAction> actions = renderer.getPages().get(0).getActions(DrawStringAction.class);

		DrawStringAction c1 = actions.get(0);
		assertEquals("Column 1", c1.getText());

		DrawStringAction c2 = actions.get(1);
		assertEquals("Column 2", c2.getText());
		assertTrue(c2.getX() > c1.getX());

		DrawStringAction c3 = actions.get(2);
		assertEquals("Column 3", c3.getText());
		assertTrue(c3.getX() > c2.getX());

		DrawStringAction r1c1 = actions.get(3);
		assertEquals("R1C1", r1c1.getText());
		assertEquals(c1.getX(), r1c1.getX());

		DrawStringAction r1c2 = actions.get(4);
		assertEquals("R1C2", r1c2.getText());
		assertEquals(c2.getX(), r1c2.getX());
		assertTrue(r1c2.getX() > r1c1.getX());

		DrawStringAction r1c3 = actions.get(5);
		assertEquals("R1C3", r1c3.getText());
		assertEquals(c3.getX(), r1c3.getX());
		assertTrue(r1c3.getX() > r1c2.getX());

		DrawStringAction r2c2 = actions.get(6);
		assertEquals("R2C2", r2c2.getText());
		assertEquals(c2.getX(), r2c2.getX());

		DrawStringAction r2c3 = actions.get(7);
		assertEquals("R2C3", r2c3.getText());
		assertEquals(c3.getX(), r2c3.getX());
		assertTrue(r2c3.getX() > r2c2.getX());
	}

	@Test
	public void testAlignment() throws ReaderException {
		new DocxReader(renderer, TEST_ALIGNMENT).process();

		List<DrawStringAction> actions = renderer.getPages().get(0).getActions(DrawStringAction.class);

		DrawStringAction left = actions.get(0);
		assertEquals("Left aligned text", left.getText());
		assertEquals(1409, left.getY());

		DrawStringAction right = actions.get(1);
		assertEquals("Right aligned text", right.getText());
		assertEquals(1684, right.getY());

		DrawStringAction center = actions.get(2);
		assertEquals("Center aligned text", center.getText());
		assertEquals(1959, center.getY());
	}

	private void assertSet(Set<FontStyle> actualStyles, FontStyle ... expectedStyles) {
		for (FontStyle expected : expectedStyles) {
			assertTrue("Expected " + expected, actualStyles.contains(expected));
		}

		assertEquals(actualStyles.size(), expectedStyles.length);
	}
}