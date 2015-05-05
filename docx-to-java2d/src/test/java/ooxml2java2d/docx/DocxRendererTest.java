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

package ooxml2java2d.docx;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.awt.Color;
import java.awt.Font;
import java.awt.font.TextAttribute;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Set;

import ooxml2java2d.docx.internal.DrawImageAction;
import ooxml2java2d.docx.internal.DrawStringAction;

import org.junit.Before;
import org.junit.Test;

public class DocxRendererTest {
	private static final File TEST_FILE_LAYOUTS = new File("src/test/resources/docx/layouts.docx");
	private static final File TEST_FILE_BODY_START_NO_HEADER = new File("src/test/resources/docx/body_start_no_header.docx");
	private static final File TEST_HEADER = new File("src/test/resources/docx/header.docx");
	private static final File TEST_HEADER_FIRST_EVEN_ODD = new File("src/test/resources/docx/header_first_even_odd.docx");
	private static final File TEST_FOOTER_FIRST_EVEN_ODD = new File("src/test/resources/docx/footer_first_even_odd.docx");
	private static final File TEST_FONT_SIZE = new File("src/test/resources/docx/font_size.docx");
	private static final File TEST_FONT_STYLE = new File("src/test/resources/docx/font_style.docx");
	private static final File TEST_LINE_HEIGHT = new File("src/test/resources/docx/line_height.docx");
	private static final File TEST_TEXT_COLOR = new File("src/test/resources/docx/text_color.docx");
	private static final File TEST_HEADER_STYLE = new File("src/test/resources/docx/header_style.docx");
	private static final File TEST_WORD_WRAP = new File("src/test/resources/docx/word_wrap.docx");
	private static final File TEST_WORD_WRAP_CONTINUOUS = new File("src/test/resources/docx/word_wrap_continuous.docx");
	private static final File TEST_TABBED = new File("src/test/resources/docx/tabbed.docx");
	private static final File TEST_TABBED2 = new File("src/test/resources/docx/tabbed2.docx");
	private static final File TEST_ALIGNMENT = new File("src/test/resources/docx/alignment.docx");
	private static final File TEST_TABLE_SIMPLE = new File("src/test/resources/docx/table_simple.docx");
	private static final File TEST_TABLE_MERGE_HORIZONTAL = new File("src/test/resources/docx/table_merge_horizontal.docx");
	private static final File TEST_IMAGE_INLINE = new File("src/test/resources/docx/image_inline.docx");
	private static final File TEST_IMAGE_ANCHOR = new File("src/test/resources/docx/image_anchor.docx");
	private static final File TEST_IMAGE_ANCHOR_PAGE_WRAP_BG = new File("src/test/resources/docx/image_anchor_page_wrap_background.docx");
	private static final File TEST_PAGE_BREAK = new File("src/test/resources/docx/page_break.docx");
	private static final File TEST_PAGE_BREAK_OVERFLOW = new File("src/test/resources/docx/page_break_overflow.docx");
	private static final File TEST_PAGE_BREAK_TABLE_OVERFLOW = new File("src/test/resources/docx/page_break_table_overflow.docx");
	private static final File TEST_PAGE_BREAK_TABLE_OVERFLOW2 = new File("src/test/resources/docx/page_break_table_overflow2.docx");
	private static final File TEST_PAGE_BREAK_TABLE_NESTED = new File("src/test/resources/docx/page_break_table_nested.docx");
	private static final File TEST_EMPTY_PARAGRAPH = new File("src/test/resources/docx/empty_paragraph.docx");
	private static final File TEST_PARAGRAPH_BREAK = new File("src/test/resources/docx/paragraph_break.docx");
	private static final File TEST_PARAGRAPH_SPACING = new File("src/test/resources/docx/paragraph_spacing.docx");
	private static final File TEST_LIST_BULLET = new File("src/test/resources/docx/list_bullet.docx");
	private MockGraphicsBuilder builder;

	@Before
	public void setUp() {
		builder = new MockGraphicsBuilder();
	}

	@Test
	public void testLayouts() throws IOException {
		new DocxRenderer(TEST_FILE_LAYOUTS).render(builder);

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
		new DocxRenderer(TEST_FILE_BODY_START_NO_HEADER).render(builder);

		List<MockGraphics2D> pages = builder.getPages();

		DrawStringAction page1Action = pages.get(0).getActions(DrawStringAction.class).get(0);
		DrawStringAction page2Action = pages.get(1).getActions(DrawStringAction.class).get(0);

		assertEquals("DrawStringAction[text=Hello, World!,x=1440,y=1708]", page1Action.toString());
		assertEquals("DrawStringAction[text=Hello, World!,x=1134,y=1969]", page2Action.toString());
	}

	@Test
	public void testHeader() throws IOException {
		new DocxRenderer(TEST_HEADER).render(builder);

		List<DrawStringAction> actions = builder.getPages().get(0).getActions(DrawStringAction.class);
		DrawStringAction a = actions.get(0);
		DrawStringAction b = actions.get(1);
		DrawStringAction c = actions.get(2);
		DrawStringAction d = actions.get(3);
		DrawStringAction e = actions.get(4);
		DrawStringAction f = actions.get(5);

		assertEquals("A", a.getText());
		assertEquals(1409, a.getY());

		assertEquals("B", b.getText());
		assertEquals(1684, b.getY());

		assertEquals("C", c.getText());
		assertEquals(1959, c.getY());

		assertEquals("D", d.getText());
		assertEquals(2520, d.getY());

		assertEquals("E", e.getText());
		assertEquals(2795, e.getY());

		assertEquals("F", f.getText());
		assertEquals(3070, f.getY());
	}

	@Test
	public void testHeaderFirstEvenOdd() throws IOException {
		new DocxRenderer(TEST_HEADER_FIRST_EVEN_ODD).render(builder);

		List<MockGraphics2D> pages = builder.getPages();
		DrawStringAction p1Header = pages.get(0).getActions(DrawStringAction.class).get(0);
		DrawStringAction p2Header = pages.get(1).getActions(DrawStringAction.class).get(0);
		DrawStringAction p3Header = pages.get(2).getActions(DrawStringAction.class).get(0);

		assertEquals("First", p1Header.getText());
		assertEquals("Even", p2Header.getText());
		assertEquals("Odd", p3Header.getText());
	}

	@Test
	public void testFooterFirstEvenOdd() throws IOException {
		new DocxRenderer(TEST_FOOTER_FIRST_EVEN_ODD).render(builder);

		List<MockGraphics2D> pages = builder.getPages();
		DrawStringAction p1Footer = pages.get(0).getActions(DrawStringAction.class).get(0);
		DrawStringAction p2Footer = pages.get(1).getActions(DrawStringAction.class).get(0);
		DrawStringAction p3Footer = pages.get(2).getActions(DrawStringAction.class).get(0);

		assertEquals("First", p1Footer.getText());
		assertEquals("Even", p2Footer.getText());
		assertEquals("Odd", p3Footer.getText());
	}

	@Test
	public void testFontSize() throws IOException {
		new DocxRenderer(TEST_FONT_SIZE).render(builder);

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
		new DocxRenderer(TEST_FONT_STYLE).render(builder);

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
		new DocxRenderer(TEST_LINE_HEIGHT).render(builder);

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
		new DocxRenderer(TEST_TEXT_COLOR).render(builder);

		List<Object> actions = builder.getPages().get(0).getActions(Color.class, DrawStringAction.class);

		assertEquals(Color.BLACK, (Color) actions.get(0));

		assertEquals(Color.RED, (Color) actions.get(1));
		assertEquals("Red", ((DrawStringAction) actions.get(2)).getText());

		assertEquals(Color.GREEN, (Color) actions.get(5));
		assertEquals("green", ((DrawStringAction) actions.get(6)).getText());

		assertEquals(Color.BLUE, (Color) actions.get(9));
		assertEquals("blue", ((DrawStringAction) actions.get(10)).getText());
	}

	@Test
	public void testHeaderStyle() throws IOException {
		new DocxRenderer(TEST_HEADER_STYLE).render(builder);

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
		new DocxRenderer(TEST_WORD_WRAP).render(builder);

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
		new DocxRenderer(TEST_WORD_WRAP_CONTINUOUS).render(builder);

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
		new DocxRenderer(TEST_TABBED).render(builder);

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
	public void testTabbed2() throws IOException {
		new DocxRenderer(TEST_TABBED2).render(builder);

		List<DrawStringAction> actions = builder.getPages().get(0).getActions(DrawStringAction.class);

		DrawStringAction c1 = actions.get(0);
		assertEquals("Column 1", c1.getText());

		DrawStringAction c2 = actions.get(1);
		assertEquals("Column 2", c2.getText());
		assertTrue(c2.getX() > c1.getX());

		// Test all rows in Column B are aligned
		DrawStringAction r1c2 = actions.get(3);
		assertEquals("B", r1c2.getText());
		assertEquals(c2.getX(), r1c2.getX());

		DrawStringAction r2c2 = actions.get(5);
		assertEquals("B", r2c2.getText());
		assertEquals(c2.getX(), r2c2.getX());

		DrawStringAction r3c2 = actions.get(7);
		assertEquals("B", r3c2.getText());
		assertEquals(c2.getX(), r3c2.getX());

		DrawStringAction r4c2 = actions.get(9);
		assertEquals("B", r4c2.getText());
		assertEquals(c2.getX(), r4c2.getX());
	}

	@Test
	public void testAlignment() throws IOException {
		new DocxRenderer(TEST_ALIGNMENT).render(builder);

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
		new DocxRenderer(TEST_TABLE_SIMPLE).render(builder);

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
	public void testTableMergeHorizontal() throws IOException {
		new DocxRenderer(TEST_TABLE_MERGE_HORIZONTAL).render(builder);

		List<DrawStringAction> actions = builder.getPages().get(0).getActions(DrawStringAction.class);

		// row 1
		DrawStringAction a1 = actions.get(0);
		assertEquals("A1", a1.getText());

		DrawStringAction b1c1 = actions.get(1);
		assertEquals("B1C1_Merge", b1c1.getText());
		assertTrue(b1c1.getX() > a1.getX());

		DrawStringAction d1 = actions.get(2);
		assertEquals("D1", d1.getText());
		assertTrue(d1.getX() > b1c1.getX());

		DrawStringAction e1 = actions.get(3);
		assertEquals("E1", e1.getText());
		assertTrue(e1.getX() > d1.getX());

		// row 2
		DrawStringAction a2b2 = actions.get(4);
		assertEquals("A2B2_Merge", a2b2.getText());
		assertEquals(a1.getX(), a2b2.getX());

		DrawStringAction c2d2 = actions.get(5);
		assertEquals("C2D2_Merge", c2d2.getText());
		assertTrue(c2d2.getX() > a2b2.getX());
		assertTrue(c2d2.getX() > b1c1.getX());

		DrawStringAction e2 = actions.get(6);
		assertEquals("E2", e2.getText());
		assertEquals(e1.getX(), e2.getX());
		assertTrue(e2.getX() > c2d2.getX());
	}

	@Test
	public void testImageInline() throws IOException {
		new DocxRenderer(TEST_IMAGE_INLINE).render(builder);

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
	public void testImageAnchor() throws IOException {
		new DocxRenderer(TEST_IMAGE_ANCHOR).render(builder);

		List<DrawImageAction> actions = builder.getPages().get(0).getActions(DrawImageAction.class);
		DrawImageAction di = actions.get(0);
		assertEquals(713, di.getWidth());
		assertEquals(540, di.getHeight());
	}

	@Test
	public void testImageAnchorPageWrapBackground() throws IOException {
		new DocxRenderer(TEST_IMAGE_ANCHOR_PAGE_WRAP_BG).render(builder);

		List<DrawImageAction> actions = builder.getPages().get(0).getActions(DrawImageAction.class);
		DrawImageAction di = actions.get(0);
		assertEquals(2219, di.getX());
	}

	/*
	 * Test that a new page is created when an explicit page break is found
	 */
	@Test
	public void testPageBreak() throws IOException {
		new DocxRenderer(TEST_PAGE_BREAK).render(builder);

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
		new DocxRenderer(TEST_PAGE_BREAK_OVERFLOW).render(builder);

		List<MockGraphics2D> pages = builder.getPages();

		// Check that the second page has the content we're expecting
		assertEquals(2, pages.size());
		assertEquals("O", pages.get(1).getActions(DrawStringAction.class).get(0).getText());
		assertEquals("P", pages.get(1).getActions(DrawStringAction.class).get(1).getText());
	}

	@Test
	public void testPageBreakTableOverflow() throws IOException {
		new DocxRenderer(TEST_PAGE_BREAK_TABLE_OVERFLOW).render(builder);

		List<MockGraphics2D> pages = builder.getPages();

		// Check that the second page has the content we're expecting
		assertEquals(2, pages.size());

		List<DrawStringAction> actions = pages.get(1).getActions(DrawStringAction.class);

		// We're only expecting a single character
		assertEquals(1, actions.size());
		assertEquals("P", actions.get(0).getText());
	}

	@Test
	public void testPageBreakTableOverflow2() throws IOException {
		new DocxRenderer(TEST_PAGE_BREAK_TABLE_OVERFLOW2).render(builder);

		List<MockGraphics2D> pages = builder.getPages();

		// Check that the second page has the content we're expecting
		assertEquals(2, pages.size());

		List<DrawStringAction> actions = pages.get(1).getActions(DrawStringAction.class);

		int baseY = actions.get(0).getY();

		for (int i = 1; i < actions.size(); i++) {
			assertEquals(baseY, actions.get(i).getY());
		}
	}

	@Test
	public void testPageBreakTableNested() throws IOException {
		new DocxRenderer(TEST_PAGE_BREAK_TABLE_NESTED).render(builder);

		List<MockGraphics2D> pages = builder.getPages();

		// Check that the second page has the content we're expecting
		assertEquals(2, pages.size());

		List<DrawStringAction> actions = pages.get(1).getActions(DrawStringAction.class);

		DrawStringAction c = actions.get(0);
		DrawStringAction d = actions.get(0);
		DrawStringAction e = actions.get(0);
		DrawStringAction g = actions.get(0);
		DrawStringAction f = actions.get(0);
		DrawStringAction h = actions.get(0);

		// check row 2 lines up horizontally
		assertEquals(c.getY(), e.getY());
		assertEquals(c.getY(), g.getY());

		// check row 3 lines up horizontally
		assertEquals(d.getY(), f.getY());
		assertEquals(d.getY(), h.getY());

		// check columns line up vertically
		assertEquals(c.getX(), d.getX());
		assertEquals(e.getX(), f.getX());
		assertEquals(g.getX(), h.getX());
	}

	@Test
	public void testEmptyParagraph() throws IOException {
		new DocxRenderer(TEST_EMPTY_PARAGRAPH).render(builder);

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
	 * Test that a paragraph break moves the following content to the next line
	 */
	@Test
	public void testParagraphBreak() throws IOException {
		new DocxRenderer(TEST_PARAGRAPH_BREAK).render(builder);

		List<DrawStringAction> actions = builder.getPages().get(0).getActions(DrawStringAction.class);

		DrawStringAction hello = actions.get(0);
		DrawStringAction world = actions.get(1);

		assertEquals("Hello", hello.getText());
		assertEquals("World", world.getText());
		assertTrue(hello.getY() < world.getY());
	}

	/*
	 * Test that paragraph spacing is used by checking "B" is pushed onto the second page
	 */
	@Test
	public void testParagraphSpacing() throws IOException {
		new DocxRenderer(TEST_PARAGRAPH_SPACING).render(builder);

		List<DrawStringAction> actions = builder.getPages().get(1).getActions(DrawStringAction.class);

		assertEquals("B", actions.get(0).getText());
	}

	/*
	 * Tests multilevel list of bullets is properly indented
	 */
	@Test
	public void testBulletList() throws IOException {
		new DocxRenderer(TEST_LIST_BULLET).render(builder);

		List<DrawStringAction> actions = builder.getPages().get(0).getActions(DrawStringAction.class);

		DrawStringAction b1 = actions.get(1);
		DrawStringAction b2 = actions.get(3);
		DrawStringAction b3 = actions.get(5);
		DrawStringAction b4 = actions.get(7);
		DrawStringAction b5 = actions.get(9);

		DrawStringAction t1 = actions.get(2);
		DrawStringAction t2 = actions.get(4);
		DrawStringAction t3 = actions.get(6);
		DrawStringAction t4 = actions.get(8);
		DrawStringAction t5 = actions.get(10);

		assertEquals("•", b1.getText());
		assertEquals(b1.getText(), b2.getText());
		assertEquals(b1.getText(), b3.getText());
		assertEquals(b1.getText(), b4.getText());
		assertEquals(b1.getText(), b5.getText());

		assertEquals("Level 1 #1", t1.getText());
		assertEquals("Level 2 #1", t2.getText());
		assertEquals("Level 3 #1", t3.getText());
		assertEquals("Level 2 #2", t4.getText());
		assertEquals("Level 1 #2", t5.getText());

		assertEquals(1938, t1.getX());
		assertEquals(2298, t2.getX());
		assertEquals(2658, t3.getX());
		assertEquals(t2.getX(), t4.getX());
		assertEquals(t1.getX(), t5.getX());
	}

	private void assertFontAttributes(Font font, Object ... expectedStyles) {
		Set<TextAttribute> actualStyles = font.getAttributes().keySet();

		for (Object style : expectedStyles) {
			assertTrue("Expected " + style, actualStyles.contains(style));
		}
	}
}