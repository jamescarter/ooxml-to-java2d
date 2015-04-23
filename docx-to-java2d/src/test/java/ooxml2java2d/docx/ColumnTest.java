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
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import ooxml2java2d.docx.internal.Column;
import ooxml2java2d.docx.internal.ContentTooBigException;
import ooxml2java2d.docx.internal.DrawStringAction;

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
