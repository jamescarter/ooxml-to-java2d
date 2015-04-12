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

package ooxmltojava2d.example.docx.svg;

import java.awt.Graphics2D;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

import org.jfree.graphics2d.svg.SVGGraphics2D;

import ooxmltojava2d.docx.DocxToGraphics2D;
import ooxmltojava2d.docx.GraphicsBuilder;

public class DocxToSvg {
	public static void main(String[] args) throws IOException {
		if (args.length != 1) {
			System.out.println("Expected a single docx file as an input");
			System.exit(1);
		}

		SvgBuilder builder = new SvgBuilder();
		new DocxToGraphics2D(builder, new File(args[0])).process();
		builder.writeToDisk();
	}

	private static class SvgBuilder implements GraphicsBuilder {
		private List<SVGGraphics2D> svgs = new ArrayList<>();

		@Override
		public Graphics2D nextPage(int pageWidth, int pageHeight) {
			SVGGraphics2D g = new SVGGraphics2D(pageWidth, pageHeight);
			svgs.add(g);
			return g;
		}

		public void writeToDisk() throws IOException {
			for (int i=0; i<svgs.size(); i++) {
				SVGGraphics2D g = svgs.get(i);

				try (PrintStream out = new PrintStream(new FileOutputStream("page" + (i + 1) + ".svg"))) {
					out.print(g.getSVGDocument());
				}
			}
		}
	}
}
