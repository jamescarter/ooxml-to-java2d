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

import java.awt.Graphics2D;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import ooxml2java2d.GraphicsBuilder;

public class MockGraphicsBuilder implements GraphicsBuilder {
	private List<MockGraphics2D> graphics = new ArrayList<>();

	@Override
	public Graphics2D nextPage(int pageWidth, int pageHeight) {
		MockGraphics2D g = new MockGraphics2D(pageWidth, pageHeight);

		graphics.add(g);

		return g;
	}

	public List<MockGraphics2D> getPages() {
		return Collections.unmodifiableList(graphics);
	}
}
