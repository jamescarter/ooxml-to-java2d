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

package ooxml2java2d.example.docx.image;

import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.imageio.ImageIO;

import ooxml2java2d.GraphicsBuilder;
import ooxml2java2d.docx.DocxRenderer;

public class DocxToImage {
	public static void main(String[] args) throws IOException {
		if (args.length != 1) {
			System.out.println("Expected a single docx file as an input");
			System.exit(1);
		}

		BufferedImageBuilder builder = new BufferedImageBuilder();
		new DocxRenderer(new File(args[0])).render(builder);
		builder.writeToDisk();
	}

	private static class BufferedImageBuilder implements GraphicsBuilder {
		private static final float scale = 0.05f; // twips-to-72dpi
		private List<BufferedImage> images = new ArrayList<>();

		@Override
		public Graphics2D nextPage(int pageWidth, int pageHeight) {
			int width = (int) (pageWidth * scale);
			int height = (int) (pageHeight * scale);
			BufferedImage bi = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
			Graphics2D g = (Graphics2D) bi.getGraphics();
			g.scale(scale, scale);
			images.add(bi);
			return g;
		}

		public void writeToDisk() throws IOException {
			for (int i = 0; i < images.size(); i++) {
				BufferedImage bi = images.get(i);

				ImageIO.write(bi, "PNG", new File("page" + (i + 1) + ".png"));
			}
		}
	}
}
