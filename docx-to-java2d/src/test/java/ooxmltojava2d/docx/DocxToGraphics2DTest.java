/*
 * Copyright (C) 2015 James Carter
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *      http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package ooxmltojava2d.docx;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.awt.Color;
import java.awt.Font;
import java.awt.font.TextAttribute;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Set;

import ooxmltojava2d.docx.DocxToGraphics2D;

import org.junit.Before;
import org.junit.Test;

public class DocxToGraphics2DTest {
	private static final File TEST_FILE_LAYOUTS = new File("src/test/resources/docx/layouts.docx");
	private static final File TEST_FILE_BODY_START_NO_HEADER = new File("src/test/resources/docx/body_start_no_header.docx");
	private static final File TEST_FONT_SIZE = new File("src/test/resources/docx/font_size.docx");
	private static final File TEST_FONT_STYLE = new File("src/test/resources/docx/font_style.docx");
	private static final File TEST_LINE_HEIGHT = new File("src/test/resources/docx/line_height.docx");
	private static final File TEST_TEXT_COLOR = new File("src/test/resources/docx/text_color.docx");
	private static final File TEST_HEADER_STYLE = new File("src/test/resources/docx/header_style.docx");
	private static final File TEST_WORD_WRAP = new File("src/test/resources/docx/word_wrap.docx");
	private static final File TEST_WORD_WRAP_CONTINUOUS = new File("src/test/resources/docx/word_wrap_continuous.docx");
	private static final File TEST_TABBED = new File("src/test/resources/docx/tabbed.docx");
	private static final File TEST_ALIGNMENT = new File("src/test/resources/docx/alignment.docx");
	private static final File TEST_TABLE_SIMPLE = new File("src/test/resources/docx/table_simple.docx");
	private static final File TEST_IMAGE_INLINE = new File("src/test/resources/docx/image_inline.docx");
	private static final File TEST_EMPTY_PARAGRAPH = new File("src/test/resources/docx/empty_paragraph.docx");
	private static final File TEST_PAGE_BREAK = new File("src/test/resources/docx/page_break.docx");
	private static final File TEST_PAGE_BREAK_OVERFLOW = new File("src/test/resources/docx/page_break_overflow.docx");
	private MockGraphicsBuilder builder;

	@Before
	public void setUp() {
		builder = new MockGraphicsBuilder();
	}

	@Test
	public void testLayouts() throws IOException {
		new DocxToGraphics2D(builder, TEST_FILE_LAYOUTS).process();

		List<MockGraphics2D> pages = builder.getPages();

		assertEquals(3, pages.size());

		MockGraphics2D page1 = pages.get(0);
		MockGraphics2D page2 = pages.get(1);
		MockGraphics2D page3 = pages.get(2);

		assertEquals(12240, page1.getWidth());
		assertEquals(15840, page1.getHeight());

		assertEquals(12983, page2.getWidth());
		assertEquals(6463, page2.getHeight());

		assertEquals(15840, page3.getWidth());
		assertEquals(12240, page3.getHeight());
	}

	@Test
	public void testBodyStartNoHeader() throws IOException {
		new DocxToGraphics2D(builder, TEST_FILE_BODY_START_NO_HEADER).process();

		List<MockGraphics2D> pages = builder.getPages();

		DrawStringAction page1Action = pages.get(0).getActions(DrawStringAction.class).get(0);
		DrawStringAction page2Action = pages.get(1).getActions(DrawStringAction.class).get(0);

		assertEquals("DrawStringAction[text=Hello, World!,x=1440,y=1708]", page1Action.toString());
		assertEquals("DrawStringAction[text=Hello, World!,x=1134,y=1969]", page2Action.toString());
	}

	@Test
	public void testFontSize() throws IOException {
		new DocxToGraphics2D(builder, TEST_FONT_SIZE).process();

		List<DrawStringAction> actions = builder.getPages().get(0).getActions(DrawStringAction.class);

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
	public void testFontStyle() throws IOException {
		new DocxToGraphics2D(builder, TEST_FONT_STYLE).process();

		List<Font> actions = builder.getPages().get(0).getActions(Font.class);

		assertEquals(actions.get(0).getStyle(), Font.BOLD);
		assertEquals(actions.get(2).getStyle(), Font.ITALIC);
		assertFontAttributes(actions.get(4), TextAttribute.UNDERLINE);
		assertFontAttributes(actions.get(6), TextAttribute.STRIKETHROUGH);
		assertFontAttributes(actions.get(8), TextAttribute.SUPERSCRIPT);
		assertFontAttributes(actions.get(10), TextAttribute.UNDERLINE, TextAttribute.SUPERSCRIPT);
	}

	@Test
	public void testLineHeight() throws IOException {
		new DocxToGraphics2D(builder, TEST_LINE_HEIGHT).process();

		List<DrawStringAction> actions = builder.getPages().get(0).getActions(DrawStringAction.class);

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
	public void testTextColor() throws IOException {
		new DocxToGraphics2D(builder, TEST_TEXT_COLOR).process();

		List<Object> actions = builder.getPages().get(0).getActions(Color.class, DrawStringAction.class);

		assertEquals(Color.BLACK, ((Color)actions.get(0)));

		assertEquals(Color.RED, ((Color)actions.get(1)));
		assertEquals("Red", ((DrawStringAction)actions.get(2)).getText());

		assertEquals(Color.GREEN, ((Color)actions.get(5)));
		assertEquals("green", ((DrawStringAction)actions.get(6)).getText());

		assertEquals(Color.BLUE, ((Color)actions.get(9)));
		assertEquals("blue", ((DrawStringAction)actions.get(10)).getText());
	}

	@Test
	public void testHeaderStyle() throws IOException {
		new DocxToGraphics2D(builder, TEST_HEADER_STYLE).process();

		List<Object> actions = builder.getPages().get(0).getActions(Font.class, Color.class, DrawStringAction.class);

		Color h3c = (Color) actions.get(11);

		Font f1 = (Font) actions.get(1);		
		Font f2 = (Font) actions.get(4);
		Font f3 = (Font) actions.get(6);
		Font f4 = (Font) actions.get(8);
		Font f5 = (Font) actions.get(10);

		assertEquals(560, (int) f1.getSize());
		assertEquals(f1.getStyle(), Font.BOLD);
		assertEquals("Arial", f1.getName());
		assertEquals("Title", ((DrawStringAction) actions.get(3)).getText());

		assertEquals(360, (int) f2.getSize());
		assertEquals("Arial", f2.getName());
		assertEquals("Subtitle", ((DrawStringAction) actions.get(5)).getText());

		assertEquals(360, (int) f3.getSize());
		assertEquals("Arial", f3.getName());
		assertEquals(f3.getStyle(), Font.BOLD);
		assertEquals("Header 1", ((DrawStringAction) actions.get(7)).getText());

		assertEquals(320, (int) f4.getSize());
		assertEquals("Arial", f4.getName());
		assertEquals(f4.getStyle(), Font.BOLD);
		assertEquals("Header 2", ((DrawStringAction) actions.get(9)).getText());

		assertEquals(280, (int) f5.getSize());
		assertEquals("Arial", f5.getName());
		assertEquals(f5.getStyle(), Font.BOLD);
		assertEquals("Header 3", ((DrawStringAction) actions.get(12)).getText());
		assertEquals(new Color(128, 128, 128), h3c);
	}

	@Test
	public void testWordWrap() throws IOException {
		new DocxToGraphics2D(builder, TEST_WORD_WRAP).process();

		List<DrawStringAction> actions = builder.getPages().get(0).getActions(DrawStringAction.class);

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
	public void testWordWrapContinuous() throws IOException {
		new DocxToGraphics2D(builder, TEST_WORD_WRAP_CONTINUOUS).process();

		List<DrawStringAction> actions = builder.getPages().get(0).getActions(DrawStringAction.class);

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
	public void testTabbed() throws IOException {
		new DocxToGraphics2D(builder, TEST_TABBED).process();

		List<DrawStringAction> actions = builder.getPages().get(0).getActions(DrawStringAction.class);

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
	public void testAlignment() throws IOException {
		new DocxToGraphics2D(builder, TEST_ALIGNMENT).process();

		List<DrawStringAction> actions = builder.getPages().get(0).getActions(DrawStringAction.class);

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

	@Test
	public void testTableSimple() throws IOException {
		new DocxToGraphics2D(builder, TEST_TABLE_SIMPLE).process();

		List<DrawStringAction> actions = builder.getPages().get(0).getActions(DrawStringAction.class);

		DrawStringAction r1c1 = actions.get(0);
		assertEquals("Header 1", r1c1.getText());
		assertEquals(1134, r1c1.getX());
		assertEquals(1409, r1c1.getY());

		DrawStringAction r1c2 = actions.get(1);
		assertEquals("Header 2", r1c2.getText());
		assertEquals(r1c1.getX(), r1c1.getX());
		assertEquals(r1c1.getY(), r1c2.getY());

		DrawStringAction r2c1 = actions.get(2);
		assertEquals("Row 1 H1", r2c1.getText());
		assertEquals(r1c1.getX(), r2c1.getX());
		assertEquals(1684, r2c1.getY());

		DrawStringAction r2c2 = actions.get(3);
		assertEquals("Row 1 H2", r2c2.getText());
		assertEquals(r1c1.getX(), r2c1.getX());
		assertEquals(r2c1.getY(), r2c2.getY());
	}

	@Test
	public void testImageInline() throws IOException {
		new DocxToGraphics2D(builder, TEST_IMAGE_INLINE).process();

		List<Object> actions = builder.getPages().get(0).getActions(DrawStringAction.class, DrawImageAction.class);

		DrawStringAction leftText = (DrawStringAction) actions.get(0);
		assertEquals("Here is an image ", leftText.getText());

		DrawImageAction image = (DrawImageAction) actions.get(1);
		assertEquals(713, image.getWidth());
		assertEquals(540, image.getHeight());
		assertTrue(image.getX() > leftText.getX());

		DrawStringAction rightText = (DrawStringAction) actions.get(02);
		assertEquals(" with some text on the other side.", rightText.getText());
		assertTrue(rightText.getX() > image.getX());
	}

	@Test
	public void testEmptyParagraph() throws IOException {
		new DocxToGraphics2D(builder, TEST_EMPTY_PARAGRAPH).process();

		List<DrawStringAction> actions = builder.getPages().get(0).getActions(DrawStringAction.class);

		DrawStringAction ds1 = actions.get(0);
		DrawStringAction ds2 = actions.get(1);
		DrawStringAction ds3 = actions.get(2);

		assertEquals("Paragraph 1", ds1.getText());
		int y = ds1.getY();

		assertEquals("Paragraph 3", ds2.getText());
		assertEquals(y += 550, ds2.getY());

		assertEquals("Paragraph 6", ds3.getText());
		assertEquals(y += 825, ds3.getY());
	}

	/*
	 * Test that a new page is created when an explicit page break is found
	 */
	@Test
	public void testPageBreak() throws IOException {
		new DocxToGraphics2D(builder, TEST_PAGE_BREAK).process();

		List<MockGraphics2D> pages = builder.getPages();

		assertEquals(2, pages.size());
		assertEquals("Page 1", pages.get(0).getActions(DrawStringAction.class).get(0).getText());
		assertEquals("Page 2", pages.get(1).getActions(DrawStringAction.class).get(0).getText());
	}

	/*
	 * Test that a new page is created when the content is too large for the current page
	 */
	@Test
	public void testPageBreakOverflow() throws IOException {
		new DocxToGraphics2D(builder, TEST_PAGE_BREAK_OVERFLOW).process();

		List<MockGraphics2D> pages = builder.getPages();

		// Check that the second page has the content we're expecting
		assertEquals(2, pages.size());
		assertEquals("O", pages.get(1).getActions(DrawStringAction.class).get(0).getText());
		assertEquals("P", pages.get(1).getActions(DrawStringAction.class).get(1).getText());
	}

	private void assertFontAttributes(Font font, Object ... expectedStyles) {
		Set<TextAttribute> actualStyles = font.getAttributes().keySet();

		for (Object style : expectedStyles) {
			assertTrue("Expected " + style, actualStyles.contains(style));
		}
	}
}