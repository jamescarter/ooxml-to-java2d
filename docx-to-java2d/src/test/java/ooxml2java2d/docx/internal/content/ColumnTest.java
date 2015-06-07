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

package ooxml2java2d.docx.internal.content;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

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
		assertTrue(column.getCurrentLine().canFitContent(10));
		assertTrue(column.getCurrentLine().canFitContent(50));
		assertTrue(column.getCurrentLine().canFitContent(99));
		assertTrue(column.getCurrentLine().canFitContent(100));
		assertFalse(column.getCurrentLine().canFitContent(101));
		assertFalse(column.getCurrentLine().canFitContent(110));
	}

	@Test
	public void testAddContent() {
		column.addContent(new StringContent(10, 10, "Text"), 0);
		column.addContent(new StringContent(90, 10, "Text"), 0);
	}

	@Test
	public void testAddContentExceptionHandling() {
		assertEquals(0, column.getRows().length);

		column.addContent(new StringContent(80, 10, "Text"), 0);
		assertEquals(1, column.getRows().length);

		// too big for current line, so should be added onto a new line
		column.addContent(new StringContent(80, 10, "Text"), 0);
		assertEquals(2, column.getRows().length);

		// too big to fit on current or a new line, so should throw an exception
		try {
			column.addContent(new StringContent(120, 10, "Text"), 0);
			fail();
		} catch (ContentTooBigException ctbe) {
			// Verify that a new line was not created
			assertEquals(2, column.getRows().length);
		}
	}

	@Test
	public void testAddVerticalSpace() {
		column.addVerticalSpace(0);
		assertTrue(column.isEmpty());
		column.addVerticalSpace(100);
		assertFalse(column.isEmpty());
	}
	
	@Test(expected = ContentTooBigException.class)
	public void testContentTooBig() {
		column.addContent(new StringContent(101, 10, "Text"), 0);
	}

	@Test
	public void testContentWidthAndReset() {
		column.addHorizontalSpace(50, 0);
		assertEquals(50, column.getCurrentLine().getContentWidth());

		column.removeRow(column.getCurrentLine());
		assertEquals(0, column.getCurrentLine().getContentWidth());
	}
}
