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

package ooxml2java2d;

import java.awt.Graphics2D;

/**
 * Builds {@link Graphics2D} objects sequentially for each page of the document being rendered. 
 */
public interface GraphicsBuilder {
	/**
	 * Returns the next {@link Graphics2D} representing a page of the given dimensions.
	 * 
	 * This is guaranteed to only be called once the previous page has been fully rendered.
	 * 
	 * @param pageWidth The page width
	 * @param pageHeight The page height
	 * @return The {@link Graphics2D} object of the given dimensions
	 */
	Graphics2D nextPage(int pageWidth, int pageHeight);
}
