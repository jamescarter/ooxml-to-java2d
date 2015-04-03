package documentconverter.renderer;

import java.awt.Color;
import java.util.ArrayList;
import java.util.List;

import org.jfree.graphics2d.svg.SVGGraphics2D;

public class SvgRenderer implements Renderer {
	private List<SVGGraphics2D> images = new ArrayList<>();

	@Override
	public Page addPage(int pageWidth, int pageHeight) {
		SVGGraphics2D g = new SVGGraphics2D(pageWidth, pageHeight);
		g.setBackground(Color.WHITE);
		g.clearRect(0, 0, pageWidth, pageHeight);

		images.add(g);

		return new Image(g);
	}

	public List<String> getSvgs() {
		List<String> svgs = new ArrayList<>();

		for (SVGGraphics2D g : images) {
			svgs.add(g.getSVGDocument());
		}

		return svgs;
	}
}
