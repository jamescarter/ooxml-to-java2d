package ooxmltojava2d.docx;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import ooxmltojava2d.docx.Column;
import ooxmltojava2d.docx.ContentTooBigException;

import org.junit.Before;
import org.junit.Test;

public class ColumnTest {
	private Column column;

	@Before
	public void setUp() {
		column = new Column(0, 100);
	}

	@Test
	public void testCanFitContent() {
		assertTrue(column.canFitContent(10));
		assertTrue(column.canFitContent(50));
		assertTrue(column.canFitContent(99));
		assertTrue(column.canFitContent(100));
		assertFalse(column.canFitContent(101));
		assertFalse(column.canFitContent(110));
	}

	@Test
	public void testAddContent() {
		column.addContent(10, 10, new DrawStringAction("Text", 0, 0));
		column.addContent(90, 10, new DrawStringAction("Text", 0, 0));
	}

	@Test(expected = ContentTooBigException.class)
	public void testContentTooBig() {
		column.addContent(101, 10, new DrawStringAction("Text", 0, 0));
	}

	@Test
	public void testContentWidthAndReset() {
		column.addContentOffset(50);
		assertEquals(50, column.getContentWidth());

		column.reset();
		assertEquals(0, column.getContentWidth());
	}
}
