package ooxmltojava2d.example.docx.image;

import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.imageio.ImageIO;

import ooxmltojava2d.docx.DocxToGraphics2D;
import ooxmltojava2d.docx.GraphicsBuilder;

public class DocxToImage {
	public static void main(String[] args) throws IOException {
		if (args.length != 1) {
			System.out.println("Expected a single docx file as an input");
			System.exit(1);
		}

		BufferedImageBuilder builder = new BufferedImageBuilder();
		new DocxToGraphics2D(builder, new File(args[0])).process();
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
			for (int i=0; i<images.size(); i++) {
				BufferedImage bi = images.get(i);

				ImageIO.write(bi, "PNG", new File("page" + (i + 1) +".png"));
			}
		}
	}
}
