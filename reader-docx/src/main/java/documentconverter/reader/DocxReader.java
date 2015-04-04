package documentconverter.reader;

import java.awt.Color;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.Iterator;
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
import org.docx4j.wml.R.Tab;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STVerticalAlignRun;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.SectPr.PgMar;
import org.docx4j.wml.SectPr.PgSz;
import org.docx4j.wml.Style;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblGridCol;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.docx4j.wml.UnderlineEnumeration;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import documentconverter.renderer.FontConfig;
import documentconverter.renderer.FontStyle;
import documentconverter.renderer.Page;
import documentconverter.renderer.Renderer;

public class DocxReader implements Reader {
	private static final Logger LOG = Logger.getLogger(DocxReader.class);
	private static final int TAB_WIDTH = 950;
	private Renderer renderer;
	private File docx;
	private MainDocumentPart main;
	private Deque<PageLayout> layouts = new ArrayDeque<>();
	private PageLayout layout;
	private Page page;
	private int yOffset;
	private ParagraphStyle defaultParaStyle;
	private ParagraphStyle paraStyle;
	private ParagraphStyle runStyle;

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
		iterateContentParts(main, new Column(layout.getLeftMargin(), layout.getWidth() - layout.getLeftMargin() - layout.getRightMargin()));
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
			size.getW().intValue(),
			size.getH().intValue(),
			margin.getTop().intValue(),
			margin.getRight().intValue(),
			margin.getBottom().intValue(),
			margin.getLeft().intValue()
		);
	}

	private void createPageFromNextLayout() {
		layout = layouts.removeFirst();
		createPageFromLayout();
	}

	private void createPageFromLayout() {
		page = renderer.addPage(layout.getWidth(), layout.getHeight());
		yOffset = layout.getTopMargin();
	}

	public void iterateContentParts(ContentAccessor ca, Column column) {
		for (Object obj : ca.getContent()) {
			if (obj instanceof P) {
				if (processParagraph((P) obj, column)) {
					column = new Column(layout.getLeftMargin(), layout.getWidth() - layout.getLeftMargin() - layout.getRightMargin());
				}
			} else if (obj instanceof R) {
				processTextRun((R) obj, column);
			} else if (obj instanceof JAXBElement) {
				JAXBElement<?> element = (JAXBElement<?>) obj;

				if (element.getDeclaredType().equals(Text.class)) {
					processText(((Text) element.getValue()).getValue(), column);
				} else if (element.getDeclaredType().equals(Tab.class)) {
					processTab((Tab) element.getValue(), column);
				} else if (element.getDeclaredType().equals(Tbl.class)) {
					processTable((Tbl) element.getValue(), column);
				} else {
					LOG.trace("Unhandled JAXBElement object " + element.getDeclaredType());
				}
			} else {
				LOG.trace("Unhandled document object " + obj.getClass());
			}
		}
	}

	// Returns true if a new page was created
	private boolean processParagraph(P p, Column column) {
		PPr properties = p.getPPr();
		ParagraphStyle newParaStyle = new ParagraphStyle(defaultParaStyle);

		if (properties != null) {
			if (properties.getPStyle() != null) {
				PStyle pstyle = properties.getPStyle();

				newParaStyle = getStyleById(defaultParaStyle, pstyle.getVal());
			}

			if (properties.getJc() != null) {
				switch(properties.getJc().getVal()) {
					case RIGHT:
						newParaStyle.setAlignment(Alignment.RIGHT);
					break;
					case CENTER:
						newParaStyle.setAlignment(Alignment.CENTER);
					break;
					default:
						// default to LEFT aligned
				}
			}
		}

		yOffset += newParaStyle.getSpaceBefore();
		paraStyle = newParaStyle;

		iterateContentParts(p, column);
		renderActionsForLine(column);

		yOffset += paraStyle.getSpaceAfter();

		if (properties!= null && properties.getSectPr() != null) {
			// The presence of SectPr indicates the next part should be started on a new page with a different layout
			createPageFromNextLayout();
			return true;
		}

		return false;
	}

	private void processTextRun(R r, Column column) {
		ParagraphStyle newRunStyle = getStyle(paraStyle, r.getRPr());

		if (runStyle == null || !newRunStyle.getFontConfig().equals(runStyle.getFontConfig())) {
			column.addAction(newRunStyle.getFontConfig());
		}

		if (runStyle == null || !newRunStyle.getColor().equals(runStyle.getColor())) {
			column.addAction(new Color(newRunStyle.getColor().getRGB()));
		}

		runStyle = newRunStyle;

		iterateContentParts(r, column);
	}

	private void processText(String text, Column column) {
		Rectangle2D bounds = runStyle.getStringBoxSize(text);

		if (column.canFitContent(bounds.getWidth())) {
			column.addContent(bounds.getWidth(), bounds.getHeight(), new DrawStringAction(text, column.getContentWidth()));
		} else {
			// Text needs wrapping, work out how to fit it into multiple lines
			// TODO: optimize word-wrapping routine
			String[] words = text.split(" ");
			String newText = "";

			for (int i=0; i<words.length; i++) {
				if (newText.isEmpty()) {
					bounds = runStyle.getStringBoxSize(words[i]);
				} else {
					bounds = runStyle.getStringBoxSize(newText + " " + words[i]);
				}
				
				// Check if adding the word will push it over the page content width
				if (!column.canFitContent(bounds.getWidth())) {
					// If this is the first word, break it up
					if (i == 0) {
						char[] chars = text.toCharArray();

						for (int k=0; k<chars.length; k++) {
							bounds = runStyle.getStringBoxSize(newText + chars[k]);

							if (column.canFitContent(bounds.getWidth())) {
								newText += chars[k];
							} else {
								break;
							}
						}
					} else {
						break;
					}
				} else if (newText.isEmpty()) {
					newText = words[i];
				} else {
					newText += " " + words[i];
				}
			}

			bounds = runStyle.getStringBoxSize(newText);
			column.addContent(bounds.getWidth(), bounds.getHeight(), new DrawStringAction(newText, column.getContentWidth()));

			renderActionsForLine(column);
			processText(text.substring(newText.length()).trim(), column);
		}
	}

	private void processTab(Tab tab, Column column) {
		// TODO: check if this should cause a new line to be created
		column.addContentOffset(TAB_WIDTH - (column.getContentWidth() % TAB_WIDTH));
	}

	private void processTable(Tbl table, Column column) {
		List<Integer> columnWidths = new ArrayList<>();

		for (TblGridCol tableColumn : table.getTblGrid().getGridCol()) {
			columnWidths.add(tableColumn.getW().intValue());
		}

		for (Object tblObj : table.getContent()) {
			if (tblObj instanceof Tr) {
				Tr tableRow = (Tr) tblObj;
				List<Integer> columnYOffsets = new ArrayList<>();
				int xOffset = column.getXOffset();
				int col = 0;

				// Set all columns y-offsets to the same position so they're all in-line
				for (int i=0; i<columnWidths.size(); i++) {
					columnYOffsets.add(yOffset);
				}

				for (Object rowObj : tableRow.getContent()) {
					if (rowObj instanceof JAXBElement) {
						JAXBElement<?> element = (JAXBElement<?>) rowObj;

						if (element.getDeclaredType().equals(Tc.class)) {
							Tc tableCell = (Tc) element.getValue();
							Column cellColumn = new Column(xOffset, columnWidths.get(col));

							// restore this columns previous y-offset before processing
							yOffset = columnYOffsets.get(col);

							iterateContentParts(tableCell, cellColumn);

							// remember where this columns content got to
							columnYOffsets.set(col, yOffset);
							xOffset += cellColumn.getWidth();
							++col;
						}
					} else {
						LOG.trace("Unhandled row object " + rowObj.getClass());
					}
				}

				// Set the next content that's output to start after the last row 
				yOffset = columnYOffsets.stream().max(Integer::compare).get();
			}
		}
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
				newStyle.setLineSpacing(spacing.getLine().intValue());
				newStyle.setSpaceBefore(spacing.getBefore().intValue());
				newStyle.setSpaceAfter(spacing.getAfter().intValue());
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
			newStyle.setFontSize(sizePt * 20);
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

	private void renderActionsForLine(Column column) {
		yOffset += column.getContentHeight();
		int alignmentOffset = column.getXOffset();

		switch(paraStyle.getAlignment()) {
			case RIGHT:
				alignmentOffset += column.getWidth() - column.getContentWidth();
			break;
			case CENTER:
				alignmentOffset += (column.getWidth() - column.getContentWidth()) / 2;
			break;
			default:
				// default to LEFT aligned
		}

		for (Object obj : column.getActions()) {
			if (obj instanceof DrawStringAction) {
				DrawStringAction dsa = (DrawStringAction) obj;
				page.drawString(dsa.getText(), alignmentOffset + dsa.getX(), yOffset);
			} else if (obj instanceof Color) {
				page.setColor(((Color) obj));
			} else if (obj instanceof FontConfig) {
				FontConfig fc = (FontConfig) obj;
				page.setFontConfig(fc);
			}
		}

		column.reset();
	}
}
