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

package ooxml2java2d.docx;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.RenderingHints;
import java.awt.geom.Rectangle2D;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.List;

import ooxml2java2d.GraphicsBuilder;
import ooxml2java2d.Renderer;
import ooxml2java2d.docx.internal.Alignment;
import ooxml2java2d.docx.internal.FontConfig;
import ooxml2java2d.docx.internal.FontStyle;
import ooxml2java2d.docx.internal.PageLayout;
import ooxml2java2d.docx.internal.ParagraphStyle;
import ooxml2java2d.docx.internal.content.BlankRow;
import ooxml2java2d.docx.internal.content.Column;
import ooxml2java2d.docx.internal.content.Content;
import ooxml2java2d.docx.internal.content.ImageContent;
import ooxml2java2d.docx.internal.content.Line;
import ooxml2java2d.docx.internal.content.Row;
import ooxml2java2d.docx.internal.content.StringContent;
import ooxml2java2d.docx.internal.content.TableRow;

import org.apache.commons.lang.StringUtils;
import org.docx4j.dml.CTPositiveSize2D;
import org.docx4j.dml.GraphicData;
import org.docx4j.dml.wordprocessingDrawing.Anchor;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.model.listnumbering.AbstractListNumberingDefinition;
import org.docx4j.model.listnumbering.ListLevel;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.wml.Br;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.Lvl;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase.NumPr;
import org.docx4j.wml.PPrBase.PStyle;
import org.docx4j.wml.PPrBase.Spacing;
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
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import javax.xml.bind.JAXBElement;
import javax.xml.namespace.QName;

/**
 * Converts a Microsoft Word (DOCX) document to Java2D commands.
 */
public class DocxRenderer implements Renderer {
	private static final Logger LOG = LoggerFactory.getLogger(DocxRenderer.class);
	private static final QName QNAME_TEXT = new QName(Namespaces.NS_WORD12, "t");
	private static final int TAB_WIDTH = 712;
	private static final int EMU_DIVISOR = 635; // divide emu by this to convert to dxa
	private static final String BULLET = new Character((char) 0x2022).toString();
	private GraphicsBuilder builder;
	private Graphics2D g;
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
	private int tableRowNesting = 0;

	public DocxRenderer(File docx) throws IOException {
		try {
			word = WordprocessingMLPackage.load(docx);
		} catch (Docx4JException e) {
			throw new IOException("Error loading document", e);
		}
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public void render(GraphicsBuilder builder) {
		this.builder = builder;
		this.main = word.getMainDocumentPart();

		setDefaultStyles();
		setPageLayouts(word);
		createPageFromNextLayout();
		iterateContentParts(main, new Column(layout.getLeftMargin(), layout.getWidth() - layout.getLeftMargin() - layout.getRightMargin()));
	}

	private void setDefaultStyles() {
		defaultParaStyle = getRunStyle(
			new ParagraphStyle(),
			main.getStyleTree().getParagraphStylesTree().get("DocDefaults").getData().getStyle().getRPr()
		);

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
			getValue(margin.getTop()),
			getValue(margin.getRight()),
			getValue(margin.getBottom()),
			getValue(margin.getLeft()),
			getValue(margin.getHeader()),
			getValue(margin.getFooter()),
			sw.getHeaderFooterPolicy()
		);
	}

	private void createPageFromNextLayout() {
		layout = layouts.removeFirst();
		createPageFromLayout();
	}

	private void createPageFromLayout() {
		Graphics2D g2 = builder.nextPage(layout.getWidth(), layout.getHeight());

		g2.setBackground(Color.WHITE);
		g2.clearRect(0, 0, layout.getWidth(), layout.getHeight());
		g2.setColor(Color.BLACK);	
		g2.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
		g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_LCD_HRGB);
		g2.setRenderingHint(RenderingHints.KEY_TEXT_LCD_CONTRAST, 250);

		// Use the settings from the previous 'page'
		if (g != null) {
			g2.setFont(g.getFont());
			g2.setColor(g.getColor());
		}

		g = g2;

		HeaderPart header = null;
		FooterPart footer = null;

		// header
		if (page == 1) {
			header = layout.getHeaderFooterPolicy().getFirstHeader();
			footer = layout.getHeaderFooterPolicy().getFirstFooter();
		}

		if (header == null) {
			header = layout.getHeaderFooterPolicy().getHeader(page);
		}
		
		if (header == null) {
			header = layout.getHeaderFooterPolicy().getDefaultHeader();
		}

		if (header != null) {
			yOffset = layout.getHeaderMargin();
			relPart = header.getRelationshipsPart(false);

			iterateContentParts(header, new Column(layout.getLeftMargin(), layout.getWidth()));
		}

		// footer
		if (footer == null) {
			footer = layout.getHeaderFooterPolicy().getFooter(page);
		}

		if (footer == null) {
			footer = layout.getHeaderFooterPolicy().getDefaultFooter();
		}

		if (footer != null) {
			relPart = footer.getRelationshipsPart(false);

			Column footerCol = new Column(layout.getLeftMargin(), layout.getWidth());

			footerCol.setBuffered(true);

			iterateContentParts(footer, footerCol);

			footerCol.setBuffered(false);

			yOffset = layout.getHeight() - layout.getFooterMargin() - footerCol.getContentHeight();

			// force footer content onto current page, otherwise it may be pushed onto a new page
			render(footerCol, true);
		}

		++page;
		relPart = main.getRelationshipsPart();
		yOffset = layout.getTopMargin();
	}

	private void iterateContentParts(ContentAccessor ca, Column column) {
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
					if (element.getName().equals(QNAME_TEXT)) {
						processText(((Text) element.getValue()).getValue(), column);
					}
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

		paraStyle = getParagraphStyle(defaultParaStyle, properties);

		column.addVerticalSpace(paraStyle.getSpaceBefore());

		if (properties != null && properties.getNumPr() != null) {
			NumPr numberingProperties = properties.getNumPr();
			String abstractNumId = numberingProperties.getNumId().getVal().toString();
			String levelReference = numberingProperties.getIlvl().getVal().toString();
			AbstractListNumberingDefinition numberingDef = main.getNumberingDefinitionsPart().getAbstractListDefinitions().get(abstractNumId);

			if (numberingDef != null) {
				ListLevel listLvl = numberingDef.getListLevels().get(levelReference);

				if (listLvl.IsBullet()) {
					Lvl lvl = listLvl.getJaxbAbstractLvl();

					paraStyle = getParagraphStyle(paraStyle, lvl.getPPr());

					Rectangle2D bounds = paraStyle.getStringBoxSize(BULLET);

					column.addHorizontalSpace(getValue(lvl.getPPr().getInd().getLeft()) - getValue(lvl.getPPr().getInd().getHanging()), paraStyle.getLineSpacing());
					column.addContent(new StringContent((int) bounds.getWidth(), (int) bounds.getHeight(), BULLET), paraStyle.getLineSpacing());
					column.addHorizontalSpace(getValue(lvl.getPPr().getInd().getHanging()), paraStyle.getLineSpacing());
				}
			}
		}

		if (properties != null && p.getContent().size() == 0) {
			column.addVerticalSpace((int) paraStyle.getStringBoxSize("").getHeight());
		} else {
			iterateContentParts(p, column);
		}

		column.addVerticalSpace(paraStyle.getSpaceAfter());

		render(column, false);

		if (properties != null && properties.getSectPr() != null) {
			// The presence of SectPr indicates the next part should be started on a new page with a different layout
			createPageFromNextLayout();
			return true;
		}

		return false;
	}

	private void processTextRun(R run, Column column) {
		ParagraphStyle newRunStyle = getRunStyle(paraStyle, run.getRPr());

		if (runStyle == null || !newRunStyle.getFontConfig().equals(runStyle.getFontConfig())) {
			column.addAction(newRunStyle.getFontConfig());
		}

		if (runStyle == null || !newRunStyle.getColor().equals(runStyle.getColor())) {
			column.addAction(new Color(newRunStyle.getColor().getRGB()));
		}

		runStyle = newRunStyle;

		if (run.getRPr() != null && run.getContent().size() == 0) {
			column.addVerticalSpace((int) paraStyle.getStringBoxSize("").getHeight());
		} else {
			iterateContentParts(run, column);
		}
	}

	private void processBreak(Br br, Column column) {
		if (br.getType() == null) {
			// paragraph break
			render(column, false);
		} else {
			switch (br.getType()) {
				case PAGE:
					render(column, false);
					createPageFromLayout();
				break;
				default:
					LOG.debug("Unhandled break type " + br.getType());
			}
		}
	}

	private void processText(String text, Column column) {
		Rectangle2D bounds = runStyle.getStringBoxSize(text);
		Line line = column.getCurrentLine();

		if (line.canFitContent(bounds.getWidth())) {
			column.addContent(new StringContent((int) bounds.getWidth(), (int) bounds.getHeight(), text), paraStyle.getLineSpacing());
		} else {
			// Text needs wrapping, work out how to fit it into multiple lines
			// TODO: optimize word-wrapping routine
			String[] words = text.split(" ");
			String newText = "";

			for (int i = 0; i < words.length; i++) {
				if (newText.isEmpty()) {
					bounds = runStyle.getStringBoxSize(words[i]);
				} else {
					bounds = runStyle.getStringBoxSize(newText + " " + words[i]);
				}

				// Check if adding the word will push it over the page content width
				if (!line.canFitContent(bounds.getWidth())) {
					// If this is the first word, break it up
					if (i == 0) {
						char[] chars = text.toCharArray();

						for (int k = 0; k < chars.length; k++) {
							bounds = runStyle.getStringBoxSize(newText + chars[k]);

							if (line.canFitContent(bounds.getWidth())) {
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

			String nextText = text.substring(newText.length()).trim();

			// Make sure we don't end up in an infinite loop unable to output the content
			if (nextText.equals(text)) {
				LOG.error("Unable to fit content, skipping: " + nextText);
			} else {
				column.addContent(new StringContent((int) bounds.getWidth(), (int) bounds.getHeight(), newText), 0);
				column.addVerticalSpace(0);

				render(column, false);
				processText(nextText, column);
			}
		}
	}

	private void processTab(Tab tab, Column column) {
		Line line = column.getCurrentLine();
		int tabWidth = TAB_WIDTH - (line.getContentWidth() % TAB_WIDTH);

		if (line.canFitContent(tabWidth)) {
			column.addContent(new Content(tabWidth, 0), paraStyle.getLineSpacing());
		} else {
			column.addContent(new Content(TAB_WIDTH, 0), paraStyle.getLineSpacing());
		}
	}

	private void processTable(Tbl table, Column column) {
		List<Integer> columnWidths = new ArrayList<>();

		for (TblGridCol tableColumn : table.getTblGrid().getGridCol()) {
			columnWidths.add(tableColumn.getW().intValue());
		}

		for (Object tblObj : table.getContent()) {
			if (tblObj instanceof Tr) {
				Tr tableRow = (Tr) tblObj;
				int xOffset = column.getXOffset();
				int col = 0;
				List<Column> cells = new ArrayList<>();

				for (Object rowObj : tableRow.getContent()) {
					if (rowObj instanceof JAXBElement) {
						JAXBElement<?> element = (JAXBElement<?>) rowObj;
						Color fill = null;

						if (element.getDeclaredType().equals(Tc.class)) {
							Tc tableCell = (Tc) element.getValue();
							int width = columnWidths.get(col);

							// Horizontal cell merge
							if (tableCell.getTcPr().getGridSpan() != null) {
								int mergeCols = tableCell.getTcPr().getGridSpan().getVal().intValue();

								for (int i = 0; i < mergeCols - 1; i++) {
									width += columnWidths.get(++col);
								}
							}

							if (tableCell.getTcPr().getShd() != null) {
								fill = getColor(tableCell.getTcPr().getShd().getFill());
							}

							Column cell = new Column(xOffset, width, fill);

							cell.setBuffered(true);

							iterateContentParts(tableCell, cell);

							cell.setBuffered(false);

							xOffset += cell.getWidth();
							++col;
							cells.add(cell);
						}
					} else {
						LOG.debug("Unhandled row object " + rowObj.getClass());
					}
				}

				column.addTableRow(new TableRow(cells));
				render(column, false);
			}
		}
	}

	private void processDrawing(Drawing drawing, Column column) {
		for (Object obj : drawing.getAnchorOrInline()) {
			if (obj instanceof Inline) {
				Inline inline = (Inline) obj;

				processGraphic(inline.getExtent(), inline.getGraphic().getGraphicData(), column);
			} else if (obj instanceof Anchor) {
				Anchor anchor = (Anchor) obj;

				if (anchor.isBehindDoc()) {
					// position image absolutely, no need to add to column
					int width = (int) anchor.getExtent().getCx() / EMU_DIVISOR;
					int height = (int) anchor.getExtent().getCy() / EMU_DIVISOR;
					int x = getValue(anchor.getPositionH().getPosOffset()) / EMU_DIVISOR;
					int y = getValue(anchor.getPositionV().getPosOffset()) / EMU_DIVISOR;

					renderImage(
						anchor.getGraphic().getGraphicData().getPic().getBlipFill().getBlip().getEmbed(),
						x,
						y,
						width,
						height
					);
				} else {
					processGraphic(anchor.getExtent(), anchor.getGraphic().getGraphicData(), column);
				}
			} else {
				LOG.debug("Unhandled drawing object " + obj.getClass());
			}
		}
	}

	private void processGraphic(CTPositiveSize2D extent, GraphicData graphicData, Column column) {
		int width = (int) extent.getCx() / EMU_DIVISOR;
		int height = (int) extent.getCy() / EMU_DIVISOR;

		// TODO: handle null pic reference
		if (graphicData.getPic() != null) {
			String rId = graphicData.getPic().getBlipFill().getBlip().getEmbed();

			// TODO: Add support for external reference
			if (!rId.isEmpty()) {
				column.addContentForced(new ImageContent(width, height, rId));
			}
		}
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
			newStyle = getParagraphStyle(newStyle, style.getPPr());
		}

		return getRunStyle(newStyle, style.getRPr());
	}

	private ParagraphStyle getParagraphStyle(ParagraphStyle baseStyle, PPr properties) {
		ParagraphStyle newStyle = new ParagraphStyle(baseStyle);

		if (properties != null) {
			if (properties.getPStyle() != null) {
				PStyle pstyle = properties.getPStyle();

				newStyle = getStyleById(baseStyle, pstyle.getVal());
			}

			Spacing spacing = properties.getSpacing();

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

			if (properties.getJc() != null) {
				switch (properties.getJc().getVal()) {
					case RIGHT:
						newStyle.setAlignment(Alignment.RIGHT);
					break;
					case CENTER:
						newStyle.setAlignment(Alignment.CENTER);
					break;
					default:
						// default to LEFT aligned
				}
			}
		}

		return newStyle;
	}

	private ParagraphStyle getRunStyle(ParagraphStyle baseStyle, RPr runProperties) {
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
			Color newColor = getColor(strColor);

			if (newColor == null) {
				newColor = Color.BLACK;
			}

			if (!newColor.equals(baseStyle.getColor())) {
				newStyle.setColor(newColor);
			}
		}

		return newStyle;
	}

	private Color getColor(String strColor) {
		if (strColor.equals("auto")) {
			return null;
		} else {
			String hex = StringUtils.leftPad(strColor, 6, '0');

			return new Color(
				Integer.valueOf(hex.substring(0, 2), 16),
				Integer.valueOf(hex.substring(2, 4), 16),
				Integer.valueOf(hex.substring(4, 6), 16)
			);
		}
	}

	private Image getImage(String relationshipId) throws IOException {
		BinaryPart binary = (BinaryPart) relPart.getPart(relationshipId);

		return ImageIO.read(new ByteArrayInputStream(binary.getBytes()));
	}

	private int getValue(Integer i) {
		return (i == null) ? 0 : i;
	}

	private int getValue(BigInteger bi) {
		return (bi == null) ? 0 : bi.intValue();
	}

	private void render(Column column, boolean forceOntoCurrentPage) {
		render(column, forceOntoCurrentPage, false, column.getContentHeight());
	}

	private void render(Column column, boolean forceOntoCurrentPage, boolean delayPageCreation, int contentHeight) {
		if (column.isBuffered()) {
			return;
		}

		//int finalYOffset = yOffset + contentHeight;
		
		if (column.getFill() != null) {
			Color origColor = g.getColor();

			g.setColor(column.getFill());
			g.fillRect(column.getXOffset(), yOffset, column.getWidth(), contentHeight);
			g.setColor(origColor);
		}

		for (Row row : column.getRows()) {
			if (row instanceof Line) {
				// Check if this line will fit onto the current page, otherwise create a new page
				if (!forceOntoCurrentPage && yOffset + row.getContentHeight() > layout.getHeight() - layout.getBottomMargin()) {
					if (column.isBuffered() || delayPageCreation) {
						return;
					} else {
						createPageFromLayout();
					}
				}

				render((Line) row, column.getXOffset());
			} else if (row instanceof BlankRow) {
				yOffset += row.getContentHeight();
			} else if (row instanceof TableRow) {
				++tableRowNesting;
				renderTableRow((TableRow) row, forceOntoCurrentPage);
				--tableRowNesting;

				// If this is the top-level table row and there's still content, create a new page and output it
				if (tableRowNesting > 0) {
					return;
				} else if (row.getContentHeight() > 0) {
					createPageFromLayout();
					renderTableRow((TableRow) row, forceOntoCurrentPage);
				}
			} else {
				LOG.debug("Unhandled row object " + row.getClass());
			}

			column.removeRow(row);
		}
	}

	private void render(Line line, int initialXOffset) {
		yOffset += line.getContentHeight();
		int xOffset = initialXOffset;

		switch (paraStyle.getAlignment()) {
			case RIGHT:
				xOffset += line.getWidth() - line.getContentWidth();
			break;
			case CENTER:
				xOffset += (line.getWidth() - line.getContentWidth()) / 2;
			break;
			default:
				// default to LEFT aligned
		}

		for (Object obj : line.getActions()) {
			if (obj instanceof Content) {
				Content content = (Content) obj;

				if (obj instanceof StringContent) {
					StringContent sc = (StringContent) obj;

					g.drawString(sc.getText(), xOffset, yOffset);
				} else if (obj instanceof ImageContent) {
					ImageContent di = (ImageContent) obj;

					renderImage(
						di.getRelationshipId(),
						xOffset,
						yOffset - di.getHeight(),
						di.getWidth(),
						di.getHeight()
					);
				} else {
					yOffset += content.getHeight();
				}

				xOffset += content.getWidth();
			} else if (obj instanceof Color) {
				g.setColor((Color) obj);
			} else if (obj instanceof FontConfig) {
				FontConfig fc = (FontConfig) obj;

				g.setFont(fc.getFont());
			} else {
				LOG.debug("Unhandled render object " + obj.getClass());
			}
		}
	}

	private void renderTableRow(TableRow row, boolean forceOntoCurrentPage) {
		int start = yOffset; // start every column from the same position
		int maxYOffset = start;
		int contentHeight = row.getContentHeight();

		// Render as much content from each column onto the current page as possible
		for (Column cell : row.getColumns()) {
			yOffset = start;
			render(cell, forceOntoCurrentPage, true, contentHeight);
			maxYOffset = Math.max(maxYOffset, yOffset);
		}

		yOffset = maxYOffset;
	}

	private void renderImage(String relationshipId, int x, int y, int width, int height) {
		try {
			Image bi = getImage(relationshipId);

			if (bi == null) {
				LOG.error("Error creating image for " + relationshipId);
			} else {
				g.drawImage(
					bi,
					x,
					y,
					width,
					height,
					null
				);
			}
		} catch (IOException ioe) {
			LOG.error("Error reading image", ioe);
		}
	}
}
