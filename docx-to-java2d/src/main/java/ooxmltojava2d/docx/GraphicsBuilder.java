package ooxmltojava2d.docx;

import java.awt.Graphics2D;

public interface GraphicsBuilder {
	Graphics2D nextPage(int pageWidth, int pageHeight);
}
