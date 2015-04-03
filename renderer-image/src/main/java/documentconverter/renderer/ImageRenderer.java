package documentconverter.renderer;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import documentconverter.renderer.Page;
import documentconverter.renderer.Renderer;

public class ImageRenderer implements Renderer {
	private static final float scale = 0.05f;
	private List<BufferedImage> images = new ArrayList<>();

	@Override
	public Page addPage(int pageWidth, int pageHeight) {
		BufferedImage bi = new BufferedImage((int) (pageWidth * scale), (int) (pageHeight * scale), BufferedImage.TYPE_INT_RGB);
		Graphics2D g = (Graphics2D) bi.getGraphics();

		g.setBackground(Color.WHITE);
		g.clearRect(0, 0, bi.getWidth(), bi.getHeight());
		g.scale(scale, scale);

		Image page = new Image(g);

		images.add(bi);

		return page;
	}

	public List<BufferedImage> getImages() {
		return Collections.unmodifiableList(images);
	}
}
