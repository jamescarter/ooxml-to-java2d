package ooxmltojava2d.example.docx.pdf;

import java.awt.Dimension;
import java.awt.Graphics2D;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.freehep.graphicsio.pdf.PDFGraphics2D;

import ooxmltojava2d.docx.DocxToGraphics2D;
import ooxmltojava2d.docx.GraphicsBuilder;

public class DocxToPdfFreeHep {
	public static void main(String[] args) throws IOException {
		if (args.length != 1) {
			System.out.println("Expected a single docx file as an input");
			System.exit(1);
		}

		PdfBuilder builder = new PdfBuilder(new File("output.pdf"));
		new DocxToGraphics2D(builder, new File(args[0])).process();
		builder.finish();
	}

	private static class PdfBuilder implements GraphicsBuilder {
		private File pdf;
		private PDFGraphics2D g;

		public PdfBuilder(File pdf) throws FileNotFoundException {
			this.pdf = pdf;
		}

		@Override
		public Graphics2D nextPage(int pageWidth, int pageHeight) {
			try {
				if (g == null) {
					g = new PDFGraphics2D(pdf, new Dimension(pageWidth, pageHeight));
					g.startExport();
				} else {
					g.closePage();
					g.openPage(new Dimension(pageWidth, pageHeight), null);
				}
			} catch (IOException ioe) {
				ioe.printStackTrace();
			}

			return g;
		}

		public void finish() {
			g.endExport();
		}
	}
}
