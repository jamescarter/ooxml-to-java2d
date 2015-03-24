package documentconverter.reader;

import java.awt.Color;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.List;

import org.apache.commons.lang.StringUtils;
import org.apache.log4j.Logger;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase.PStyle;
import org.docx4j.wml.PPrBase.Spacing;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.SectPr.PgMar;
import org.docx4j.wml.SectPr.PgSz;
import org.docx4j.wml.Style;
import org.docx4j.wml.Text;

import javax.xml.bind.JAXBElement;

import documentconverter.renderer.ColorAction;
import documentconverter.renderer.FontConfig;
import documentconverter.renderer.Page;
import documentconverter.renderer.Renderer;

public class DocxReader implements Reader {
	private static final Logger LOG = Logger.getLogger(DocxReader.class);
	private Renderer renderer;
	private File docx;
	private MainDocumentPart main;
	private Deque<PageLayout> layouts = new ArrayDeque<>();
	private PageLayout layout;
	private Page page;
	private int xOffset;
	private int yOffset;
	private FontConfig fontConfig = new FontConfig();
	private int lineSpacing;
	private int paragraphSpacingAfter;
	private int paragraphSpacingBefore;
	private Color color = Color.BLACK;

	/**
	 * The actions that represent the current line. These are stored until the
	 * line height can be calculated.
	 */
	private List<Object> actions = new ArrayList<>();

	/**
	 * The current maximum height for the current line being processed.
	 * This is effectively the height of the objects stored in the actions list.
	 */
	private double lineHeight;

	public DocxReader(Renderer renderer, File docx) {
		this.renderer = renderer;
		this.docx = docx;
	}

	@Override
	public void process() throws ReaderException {
		WordprocessingMLPackage word;

		try {
			word = WordprocessingMLPackage.load(docx);
		} catch (Docx4JException e) {
			throw new ReaderException("Error loading document", e);
		}

		main = word.getMainDocumentPart();

		setPageLayouts(word);
		createPageFromNextLayout();
		iterateContentParts(main);
	}

	private void setPageLayouts(WordprocessingMLPackage word) {
		for (SectionWrapper section : word.getDocumentModel().getSections()) {
			layouts.add(createPageLayout(section.getSectPr()));
		}
	}

	private PageLayout createPageLayout(SectPr sectPr) {
		PgSz size = sectPr.getPgSz();
		PgMar margin = sectPr.getPgMar();

		return new PageLayout(
			getSize(size.getW().intValue()),
			getSize(size.getH().intValue()),
			getSize(margin.getTop().intValue()),
			getSize(margin.getRight().intValue()),
			getSize(margin.getBottom().intValue()),
			getSize(margin.getLeft().intValue())
		);
	}

	private void createPageFromNextLayout() {
		layout = layouts.removeFirst();
		createPageFromLayout();
	}

	private void createPageFromLayout() {
		page = renderer.addPage(layout.getWidth(), layout.getHeight());
		xOffset = layout.getLeftMargin();
		yOffset = layout.getTopMargin();
	}

	public void iterateContentParts(ContentAccessor ca) {
		for (Object obj : ca.getContent()) {
			if (obj instanceof P) {
				processParagraph((P) obj);
			} else if (obj instanceof R) {
				processTextRun((R) obj);
			} else if (obj instanceof JAXBElement) {
				JAXBElement<?> element = (JAXBElement<?>) obj;

				if (element.getDeclaredType().equals(Text.class)) {
					processText((Text) element.getValue());
				} else {
					LOG.trace("Unhandled JAXBElement class " + obj.getClass());
				}
			} else {
				LOG.trace("Unhandled document class " + obj.getClass());
			}
		}
	}

	private void processParagraph(P p) {
		PPr properties = p.getPPr();
		PStyle pstyle = properties.getPStyle();

		if (pstyle != null) {
			Style style = main.getStyleDefinitionsPart().getStyleById(pstyle.getVal());

			setStyle(style);
		}

		iterateContentParts(p);
		renderActionsForLine();

		if (properties.getSectPr() != null) {
			// The presence of SectPr indicates the next part should be started on a new page with a different layout
			createPageFromNextLayout();
		}
	}

	private void processTextRun(R r) {
		if (r.getRPr() != null) {
			setStyle(r.getRPr());
		}

		iterateContentParts(r);
	}

	private void processText(Text text) {
		actions.add(new DrawStringAction(text.getValue(), xOffset));
		Rectangle2D bounds = fontConfig.getStringBoxSize(text.getValue());
		xOffset += bounds.getWidth();
		lineHeight = Math.max(lineHeight, bounds.getHeight());
	}

	private void setStyle(Style style) {
		if (style.getPPr() != null) {
			Spacing spacing = style.getPPr().getSpacing();

			if (spacing != null) {
				lineSpacing = getSize(spacing.getLine().intValue());
				paragraphSpacingAfter = getSize(spacing.getAfter().intValue());
				paragraphSpacingBefore = getSize(spacing.getBefore().intValue());
			}
		}

		setStyle(style.getRPr());
	}

	private void setStyle(RPr runProperties) {
		boolean fontConfigChanged = false;

		// font
		if (runProperties.getRFonts() != null) {
			fontConfig.setName(runProperties.getRFonts().getAscii());
			fontConfigChanged = true;
		}

		// font size
		if (runProperties.getSz() != null) {
			float sizePt = runProperties.getSz().getVal().floatValue() / 2;

			// scale by a factor of 20 for trips units
			fontConfig.setSize(getSize(sizePt * 20));
			fontConfigChanged = true;
		}

		// text color
		Color newColor = Color.BLACK;

		if (runProperties.getColor() != null) {
			String strColor = runProperties.getColor().getVal();

			if (strColor.equals("auto")) {
				newColor = Color.BLACK;
			} else {
				String hex = StringUtils.leftPad(strColor, 6, '0');

				newColor = new Color(
					Integer.valueOf(hex.substring(0, 2), 16),
					Integer.valueOf(hex.substring(2, 4), 16),
					Integer.valueOf(hex.substring(4, 6), 16)
				);
			}
		}

		// Only add to actions if there were changes to avoid redundant commands
		if (fontConfigChanged) {
			actions.add(new FontConfigAction(fontConfig.getName(), fontConfig.getSize()));
		}

		if (newColor != color) {
			actions.add(new ColorAction(newColor));
			this.color = newColor;
		}
	}

	private void renderActionsForLine() {
		xOffset = layout.getLeftMargin();
		yOffset += lineHeight + paragraphSpacingBefore;

		for (Object obj : actions) {
			if (obj instanceof DrawStringAction) {
				DrawStringAction ds = (DrawStringAction) obj;
				page.drawString(ds.getText(), ds.getX(), yOffset);
			} else if (obj instanceof ColorAction) {
				page.setColor(((ColorAction) obj).getColor());
			} else if (obj instanceof FontConfigAction) {
				FontConfigAction fc = (FontConfigAction) obj;
				page.setFontConfig(new FontConfig(fc.getName(), fc.getSize()));
			}
		}

		yOffset += paragraphSpacingAfter;
		lineHeight = 0;
		actions.clear();
	}

	// Convert from 72 dpi to 96
	private float getSize(float f) {
		return f * 96 / 72;
	}

	private int getSize(int i) {
		return i * 96 / 72;
	}
}
