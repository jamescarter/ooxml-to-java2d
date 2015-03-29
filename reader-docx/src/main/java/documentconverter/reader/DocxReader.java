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
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase.PStyle;
import org.docx4j.wml.PPrBase.Spacing;
import org.docx4j.wml.DocDefaults;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STVerticalAlignRun;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.SectPr.PgMar;
import org.docx4j.wml.SectPr.PgSz;
import org.docx4j.wml.Style;
import org.docx4j.wml.Text;
import org.docx4j.wml.UnderlineEnumeration;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import documentconverter.renderer.FontConfig;
import documentconverter.renderer.FontStyle;
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
	private ParagraphStyle defaultParaStyle;
	private ParagraphStyle paraStyle;
	private ParagraphStyle runStyle;

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

		setDefaultStyles();
		setPageLayouts(word);
		createPageFromNextLayout();
		iterateContentParts(main);
	}

	private void setDefaultStyles() throws ReaderException {
		List<Object> docs;

		try {
			docs = (List<Object>) main.getStyleDefinitionsPart().getJAXBNodesViaXPath("/w:styles/w:docDefaults", false);
		} catch (XPathBinderAssociationIsPartialException | JAXBException e) {
			throw new ReaderException("Error retrieving default styles", e);
		}

		DocDefaults docDef = (DocDefaults) docs.get(0);
		defaultParaStyle = getStyle(new ParagraphStyle(), docDef.getRPrDefault().getRPr());
	}

	private void setPageLayouts(WordprocessingMLPackage word) {
		for (SectionWrapper section : word.getDocumentModel().getSections()) {
			// TODO: support header and footers - section.getHeaderFooterPolicy()
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

		if (properties != null && properties.getPStyle() != null) {
			PStyle pstyle = properties.getPStyle();

			paraStyle = getStyleById(defaultParaStyle, pstyle.getVal());
		} else {
			paraStyle = defaultParaStyle;
		}

		iterateContentParts(p);
		renderActionsForLine();

		if (properties!= null && properties.getSectPr() != null) {
			// The presence of SectPr indicates the next part should be started on a new page with a different layout
			createPageFromNextLayout();
		}
	}

	private void processTextRun(R r) {
		ParagraphStyle newRunStyle = getStyle(paraStyle, r.getRPr());

		if (runStyle == null || !newRunStyle.getFontConfig().equals(runStyle.getFontConfig())) {
			actions.add(newRunStyle.getFontConfig());
		}

		if (runStyle == null || !newRunStyle.getColor().equals(runStyle.getColor())) {
			actions.add(new Color(newRunStyle.getColor().getRGB()));
		}

		runStyle = newRunStyle;

		iterateContentParts(r);
	}

	private void processText(Text text) {
		actions.add(new DrawStringAction(text.getValue(), xOffset));
		Rectangle2D bounds = runStyle.getStringBoxSize(text.getValue());
		xOffset += bounds.getWidth();
		lineHeight = Math.max(lineHeight, bounds.getHeight());
	}

	private ParagraphStyle getStyleById(ParagraphStyle baseStyle, String styleId) {
		return getStyle(baseStyle, main.getStyleDefinitionsPart().getStyleById(styleId));
	}

	private ParagraphStyle getStyle(ParagraphStyle baseStyle, Style style) {
		ParagraphStyle newStyle;

		if (style.getBasedOn() == null) {
			newStyle = new ParagraphStyle(baseStyle);
		} else {
			newStyle = getStyleById(baseStyle, style.getBasedOn().getVal());
		}

		if (style.getPPr() != null) {
			Spacing spacing = style.getPPr().getSpacing();

			if (spacing != null && spacing.getLine() != null) {
				newStyle.setLineSpacing(getSize(spacing.getLine().intValue()));
				newStyle.setSpaceBefore(getSize(spacing.getBefore().intValue()));
				newStyle.setSpaceAfter(getSize(spacing.getAfter().intValue()));
			}
		}

		return getStyle(newStyle, style.getRPr());
	}

	private ParagraphStyle getStyle(ParagraphStyle baseStyle, RPr runProperties) {
		if (runProperties == null) {
			return baseStyle;
		}

		ParagraphStyle newStyle = new ParagraphStyle(baseStyle);

		// font
		if (runProperties.getRFonts() != null) {
			newStyle.setFontName(runProperties.getRFonts().getAscii());
		}

		// font size
		if (runProperties.getSz() != null) {
			float sizePt = runProperties.getSz().getVal().floatValue() / 2;

			// scale by a factor of 20 for trips units
			newStyle.setFontSize(getSize(sizePt * 20));
		}

		// style
		// TODO: support other style types (subscript, outline, shadow, ...)
		if (runProperties.getB() != null) {
			newStyle.enableFontStyle(FontStyle.BOLD);
		}

		if (runProperties.getI() != null) {
			newStyle.enableFontStyle(FontStyle.ITALIC);
		}

		if (runProperties.getStrike() != null) {
			newStyle.enableFontStyle(FontStyle.STRIKETHROUGH);
		}

		if (runProperties.getU() != null && runProperties.getU().getVal().equals(UnderlineEnumeration.SINGLE)) {
			// TODO: support the other underline types
			newStyle.enableFontStyle(FontStyle.UNDERLINE);
		}

		if (runProperties.getVertAlign() != null && runProperties.getVertAlign().getVal().equals(STVerticalAlignRun.SUPERSCRIPT)) {
			newStyle.enableFontStyle(FontStyle.SUPERSCRIPT);
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

		newStyle.setColor(newColor);

		return newStyle;
	}

	private void renderActionsForLine() {
		xOffset = layout.getLeftMargin();
		yOffset += lineHeight + paraStyle.getSpaceBefore();

		for (Object obj : actions) {
			if (obj instanceof DrawStringAction) {
				DrawStringAction dsa = (DrawStringAction) obj;
				page.drawString(dsa.getText(), dsa.getX(), yOffset);
			} else if (obj instanceof Color) {
				page.setColor(((Color) obj));
			} else if (obj instanceof FontConfig) {
				FontConfig fc = (FontConfig) obj;
				page.setFontConfig(fc);
			}
		}

		yOffset += paraStyle.getSpaceAfter();
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
