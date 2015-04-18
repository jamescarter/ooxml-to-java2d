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

package ooxmltojava2d.docx;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.List;

import org.apache.commons.lang.StringUtils;
import org.apache.log4j.Logger;
import org.docx4j.dml.CTPositiveSize2D;
import org.docx4j.dml.Graphic;
import org.docx4j.dml.wordprocessingDrawing.Anchor;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.wml.Br;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Drawing;
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

import javax.imageio.ImageIO;
import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

public class DocxToGraphics2D {
	private static final Logger LOG = Logger.getLogger(DocxToGraphics2D.class);
	private static final int TAB_WIDTH = 950;
	private static final int EMU_DIVISOR = 635; // divide emu by this to convert to dxa
	private GraphicsBuilder builder;
	private Graphics2D g;
	private File docx;
	private WordprocessingMLPackage word;
	private MainDocumentPart main;
	private Deque<PageLayout> layouts = new ArrayDeque<>();
	private PageLayout layout;
	private int page = 1;
	private int yOffset;
	private ParagraphStyle defaultParaStyle;
	private ParagraphStyle paraStyle;
	private ParagraphStyle runStyle;
	private RelationshipsPart relPart;

	public DocxToGraphics2D(GraphicsBuilder builder, File docx) {
		this.builder = builder;
		this.docx = docx;
	}

	public void process() throws IOException {
		try {
			word = WordprocessingMLPackage.load(docx);
		} catch (Docx4JException e) {
			throw new IOException("Error loading document", e);
		}

		main = word.getMainDocumentPart();

		setDefaultStyles();
		setPageLayouts(word);
		createPageFromNextLayout();
		iterateContentParts(main, new Column(layout.getLeftMargin(), layout.getWidth() - layout.getLeftMargin() - layout.getRightMargin()));
	}

	private void setDefaultStyles() throws IOException {
		List<Object> docs;

		try {
			docs = (List<Object>) main.getStyleDefinitionsPart().getJAXBNodesViaXPath("/w:styles/w:docDefaults", false);
		} catch (XPathBinderAssociationIsPartialException | JAXBException e) {
			throw new IOException("Error retrieving default styles", e);
		}

		DocDefaults docDef = (DocDefaults) docs.get(0);
		defaultParaStyle = getStyle(new ParagraphStyle(), docDef.getRPrDefault().getRPr());

		if (main.getStyleDefinitionsPart().getDefaultParagraphStyle() != null) {
			defaultParaStyle = getStyle(defaultParaStyle, main.getStyleDefinitionsPart().getDefaultParagraphStyle());
		}
	}

	private void setPageLayouts(WordprocessingMLPackage word) {
		for (SectionWrapper sw : word.getDocumentModel().getSections()) {
			layouts.add(createPageLayout(sw));
		}
	}

	private PageLayout createPageLayout(SectionWrapper sw) {
		SectPr sectPr = sw.getSectPr();
		PgSz size = sectPr.getPgSz();
		PgMar margin = sectPr.getPgMar();

		return new PageLayout(
			size.getW().intValue(),
			size.getH().intValue(),
			margin.getTop().intValue(),
			margin.getRight().intValue(),
			margin.getBottom().intValue(),
			margin.getLeft().intValue(),
			margin.getHeader().intValue(),
			sw.getHeaderFooterPolicy()
		);
	}

	private void createPageFromNextLayout() {
		layout = layouts.removeFirst();
		createPageFromLayout();
	}

	private void createPageFromLayout() {
		g = builder.nextPage(layout.getWidth(), layout.getHeight());
		g.setBackground(Color.WHITE);
		g.clearRect(0, 0, layout.getWidth(), layout.getHeight());
		g.setColor(Color.BLACK);	
		g.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
		g.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_LCD_HRGB);
		g.setRenderingHint(RenderingHints.KEY_TEXT_LCD_CONTRAST, 250);

		HeaderPart header = null;

		if (page == 1) {
			header = layout.getHeaderFooterPolicy().getFirstHeader();
		}

		if (header == null) {
			header = layout.getHeaderFooterPolicy().getHeader(page);
		}

		if (header == null) {
			header = layout.getHeaderFooterPolicy().getDefaultHeader();
		}

		if (header != null) {
			yOffset = layout.getHeaderMargin();
			relPart = header.getRelationshipsPart();

			iterateContentParts(header, new Column(layout.getLeftMargin(), layout.getWidth()));
		}

		++page;
		relPart = main.getRelationshipsPart();
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
			} else if (obj instanceof Br) {
				processBreak((Br) obj, column);
			} else if (obj instanceof JAXBElement) {
				JAXBElement<?> element = (JAXBElement<?>) obj;

				if (element.getDeclaredType().equals(Text.class)) {
					processText(((Text) element.getValue()).getValue(), column);
				} else if (element.getDeclaredType().equals(Tab.class)) {
					processTab((Tab) element.getValue(), column);
				} else if (element.getDeclaredType().equals(Tbl.class)) {
					processTable((Tbl) element.getValue(), column);
				} else if (element.getDeclaredType().equals(Drawing.class)) {
					processDrawing((Drawing) element.getValue(), column);
				} else {
					LOG.debug("Unhandled JAXBElement object " + element.getDeclaredType());
				}
			} else {
				LOG.debug("Unhandled document object " + obj.getClass());
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

		if (p.getContent().size() == 0) {
			yOffset += paraStyle.getStringBoxSize("").getHeight();
		} else {
			iterateContentParts(p, column);
		}

		renderActionsForLine(column);

		yOffset += paraStyle.getSpaceAfter();

		if (properties!= null && properties.getSectPr() != null) {
			// The presence of SectPr indicates the next part should be started on a new page with a different layout
			createPageFromNextLayout();
			return true;
		}

		return false;
	}

	private void processTextRun(R run, Column column) {
		ParagraphStyle newRunStyle = getStyle(paraStyle, run.getRPr());

		if (runStyle == null || !newRunStyle.getFontConfig().equals(runStyle.getFontConfig())) {
			column.addAction(newRunStyle.getFontConfig());
		}

		if (runStyle == null || !newRunStyle.getColor().equals(runStyle.getColor())) {
			column.addAction(new Color(newRunStyle.getColor().getRGB()));
		}

		runStyle = newRunStyle;

		if (run.getRPr() != null && run.getContent().size() == 0) {
			yOffset += paraStyle.getStringBoxSize("").getHeight();
		} else {
			iterateContentParts(run, column);
		}
	}

	private void processBreak(Br br, Column column) {
		if (br.getType() == null) {
			// paragraph break
			renderActionsForLine(column);
		} else {
			switch (br.getType()) {
				case PAGE:
					renderActionsForLine(column);
					createPageFromLayout();
				break;
				default:
					LOG.debug("Unhandled break type " + br.getType());
			}
		}
	}

	private void processText(String text, Column column) {
		Rectangle2D bounds = runStyle.getStringBoxSize(text);

		if (column.canFitContent(bounds.getWidth())) {
			column.addContent(bounds.getWidth(), bounds.getHeight(), new DrawStringAction(text, column.getContentWidth(), 0));
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
			column.addContent(bounds.getWidth(), bounds.getHeight(), new DrawStringAction(newText, column.getContentWidth(), 0));

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
							int width = columnWidths.get(col);

							// Horizontal cell merge
							if (tableCell.getTcPr().getGridSpan() != null) {
								int mergeCols = tableCell.getTcPr().getGridSpan().getVal().intValue();
								width = columnWidths.get(col);

								for (int i=0; i<mergeCols-1; i++) {
									width += columnWidths.get(++col);
								}
							}

							Column cellColumn = new Column(xOffset, width);

							// restore this columns previous y-offset before processing
							yOffset = columnYOffsets.get(col);

							iterateContentParts(tableCell, cellColumn);

							// remember where this columns content got to
							columnYOffsets.set(col, yOffset);
							xOffset += cellColumn.getWidth();
							++col;
						}
					} else {
						LOG.debug("Unhandled row object " + rowObj.getClass());
					}
				}

				// Set the next content that's output to start after the last row 
				yOffset = columnYOffsets.stream().max(Integer::compare).get();
			}
		}
	}

	private void processDrawing(Drawing drawing, Column column) {
		for (Object obj : drawing.getAnchorOrInline()) {
			if (obj instanceof Inline) {
				Inline inline = (Inline) obj;
				processGraphic(inline.getExtent(), inline.getGraphic(), column);
			} else if (obj instanceof Anchor) {
				Anchor anchor = (Anchor) obj;
				processGraphic(anchor.getExtent(), anchor.getGraphic(), column);
			} else {
				LOG.debug("Unhandled drawing object " + obj.getClass());
			}
		}
	}

	private void processGraphic(CTPositiveSize2D extent, Graphic graphic, Column column) {
		int width = (int) extent.getCx() / EMU_DIVISOR;
		int height = (int) extent.getCy() / EMU_DIVISOR;

		if (!column.canFitContent(width)) {
			renderActionsForLine(column);
		}

		String rId = graphic.getGraphicData().getPic().getBlipFill().getBlip().getEmbed();

		column.addContent(width, height, new DrawImageAction(rId, width, height, column.getContentWidth()));
	}

	private ParagraphStyle getStyleById(ParagraphStyle baseStyle, String styleId) {
		return getStyle(baseStyle, main.getStyleDefinitionsPart().getStyleById(styleId));
	}

	private ParagraphStyle getStyle(ParagraphStyle baseStyle, Style style) {
		if (style == null) {
			return baseStyle;
		}

		ParagraphStyle newStyle;

		if (style.getBasedOn() == null) {
			newStyle = new ParagraphStyle(baseStyle);
		} else {
			newStyle = getStyleById(baseStyle, style.getBasedOn().getVal());
		}

		if (style.getPPr() != null) {
			Spacing spacing = style.getPPr().getSpacing();

			if (spacing != null) {
				if (spacing.getLine() != null) {
					newStyle.setLineSpacing(spacing.getLine().intValue());
				}

				if (spacing.getBefore() != null) {
					newStyle.setSpaceBefore(spacing.getBefore().intValue());
				}

				if (spacing.getAfter() != null) {
					newStyle.setSpaceAfter(spacing.getAfter().intValue());
				}
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
			if (runProperties.getRFonts().getAscii() != null) {
				newStyle.setFontName(runProperties.getRFonts().getAscii());
			}
		}

		// font size
		if (runProperties.getSz() != null) {
			float sizePt = runProperties.getSz().getVal().floatValue() / 2;

			// scale by a factor of 20 for trips units
			newStyle.setFontSize(sizePt * 20);
		} else if (runProperties.getSzCs() != null) {
			float sizePt = runProperties.getSzCs().getVal().floatValue() / 2;

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
		if (runProperties.getColor() != null) {
			String strColor = runProperties.getColor().getVal();
			Color newColor;

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

			if (!newColor.equals(baseStyle.getColor())) {
				newStyle.setColor(newColor);
			}
		}

		return newStyle;
	}

	private void renderActionsForLine(Column column) {
		// Check if this line will fit onto the current page, otherwise create a new page
		if (yOffset + column.getContentHeight() > layout.getHeight() - layout.getBottomMargin()) {
			createPageFromLayout();
		}

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
				DrawStringAction ds = (DrawStringAction) obj;
				g.drawString(ds.getText(), alignmentOffset + ds.getX(), yOffset);
			} else if (obj instanceof Color) {
				g.setColor(((Color) obj));
			} else if (obj instanceof FontConfig) {
				FontConfig fc = (FontConfig) obj;
				g.setFont(fc.getFont());
			} else if (obj instanceof DrawImageAction) {
				DrawImageAction di = (DrawImageAction) obj;

				try {
					BufferedImage bi = getImage(di.getRelationshipId());

					g.drawImage(
						bi,
						alignmentOffset + di.getX(),
						yOffset - di.getHeight(),
						di.getWidth(),
						di.getHeight(),
						null
					);
				} catch (IOException ioe) {
					LOG.error("Error reading image", ioe);
				}
			}
		}

		column.reset();
	}

	private BufferedImage getImage(String relationshipId) throws IOException {
		BinaryPart binary = (BinaryPart) relPart.getPart(relationshipId);

		return ImageIO.read(new ByteArrayInputStream(binary.getBytes()));
	}
}
