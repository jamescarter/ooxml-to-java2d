package documentconverter.renderer;

import java.awt.image.BufferedImage;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import documentconverter.renderer.Page;
import documentconverter.renderer.Renderer;

public class ImageRenderer implements Renderer {
	private List<BufferedImage> images = new ArrayList<>();

	@Override
	public Page addPage(int pageWidth, int pageHeight) {
		Image page = new Image(pageWidth, pageHeight);

		images.add(page.getImage());

		return page;
	}

	public List<BufferedImage> getImages() {
		return Collections.unmodifiableList(images);
	}
}
