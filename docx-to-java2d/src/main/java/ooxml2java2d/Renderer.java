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

/**
 * Converts a document to Java2D commands.
 *
 * Renderer implementations must only call {@link GraphicsBuilder#nextPage(int, int)} when the current
 * page has been completely rendered and must not render to a previous page. 
 */
public interface Renderer {
	/**
	 * Starts the document rendering process
	 * 
	 * @param builder The builder that will provide the {@link java.awt.Graphics2D} objects to render pages to.
	 */
	void render(GraphicsBuilder builder);
}
