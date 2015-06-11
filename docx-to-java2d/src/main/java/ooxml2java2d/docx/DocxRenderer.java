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
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.List;

import ooxml2java2d.GraphicsBuilder;
import ooxml2java2d.Renderer;
import ooxml2java2d.docx.internal.HAlignment;
import ooxml2java2d.docx.internal.FontStyle;
import ooxml2java2d.docx.internal.GraphicsRenderer;
import ooxml2java2d.docx.internal.PageInitiationAdapter;
import ooxml2java2d.docx.internal.PageLayout;
import ooxml2java2d.docx.internal.ParagraphStyle;
import ooxml2java2d.docx.internal.VAlignment;
import ooxml2java2d.docx.internal.content.Border;
import ooxml2java2d.docx.internal.content.BorderStyle;
import ooxml2java2d.docx.internal.content.Column;
import ooxml2java2d.docx.internal.content.Content;
import ooxml2java2d.docx.internal.content.ImageContent;
import ooxml2java2d.docx.internal.content.Line;
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
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.wml.Br;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.CTHeight;
import org.docx4j.wml.CTTblCellMar;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.Lvl;
import org.docx4j.wml.P;
import org.docx4j.wml.P.Hyperlink;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase.Ind;
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
import org.docx4j.wml.TblWidth;
import org.docx4j.wml.Tc;
import org.docx4j.wml.TcMar;
import org.docx4j.wml.TcPrInner.TcBorders;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.docx4j.wml.UnderlineEnumeration;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

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
	private WordprocessingMLPackage word;
	private MainDocumentPart main;
	private PageInitiationAdapter initiation;
	private GraphicsRenderer renderer;
	private Deque<PageLayout> layouts;
	private PageLayout layout;
	private int page = 1;
	private ParagraphStyle defaultParaStyle;
	private ParagraphStyle paraStyle;
	private ParagraphStyle runStyle;
	private RelationshipsPart relationshipPart;

	public DocxRenderer(File docx) throws IOException {
		try {
			this.word = WordprocessingMLPackage.load(docx);
			this.main = word.getMainDocumentPart();
		} catch (Docx4JException e) {
			throw new IOException("Error loading document", e);
		}

		this.initiation = new PageInitiationAdapter() {
			@Override
			public void initiatePage() {
				initPage();
			}
		};
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public void render(GraphicsBuilder builder) {
		this.renderer = new GraphicsRenderer(builder, initiation);
		this.layouts = getPageLayouts(word);

		setDefaultStyles();
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

	private Deque<PageLayout> getPageLayouts(WordprocessingMLPackage word) {
		Deque<PageLayout> layouts = new ArrayDeque<>();

		for (SectionWrapper sw : word.getDocumentModel().getSections()) {
			layouts.add(createPageLayout(sw));
		}

		return layouts;
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
		renderer.nextPage(layout.getWidth(), layout.getHeight());
	}

	private void initPage() {
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
			relationshipPart = header.getRelationshipsPart(false);
			renderer.setYOffset(layout.getHeaderMargin());

			iterateContentParts(header, new Column(layout.getLeftMargin(), layout.getWidth()));

			if (renderer.getYOffset() < layout.getTopMargin()) {
				renderer.setYOffset(layout.getTopMargin());
			}
		} else {
			renderer.setYOffset(layout.getTopMargin());
		}

		int headerEndYOffset = renderer.getYOffset();

		// footer
		if (footer == null) {
			footer = layout.getHeaderFooterPolicy().getFooter(page);
		}

		if (footer == null) {
			footer = layout.getHeaderFooterPolicy().getDefaultFooter();
		}

		int footerStart;

		if (footer != null) {
			relationshipPart = footer.getRelationshipsPart(false);

			Column footerCol = new Column(layout.getLeftMargin(), layout.getWidth());

			footerCol.setBuffered(true);

			iterateContentParts(footer, footerCol);

			footerCol.setBuffered(false);
			footerStart = layout.getHeight() - layout.getFooterMargin() - footerCol.getContentHeight();

			renderer.setYOffset(footerStart);
			renderer.renderColumn(footerCol);
		} else {
			footerStart = layout.getHeight() - layout.getBottomMargin();
		}

		++page;
		relationshipPart = main.getRelationshipsPart();
		renderer.setYOffset(headerEndYOffset);
		renderer.setEndPosition(footerStart);
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
				} else if (element.getDeclaredType().equals(Hyperlink.class)) {
					processHyperlink((Hyperlink) element.getValue(), column);
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

		column.addAction(paraStyle.getHAlignment());
		column.addVerticalSpace(paraStyle.getSpaceBefore());

		Column paraContent = new Column(column.getXOffset() + paraStyle.getIndentLeft(), column.getWidth() - paraStyle.getIndentLeft() - paraStyle.getIndentRight());

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

					paraContent = new Column(column.getXOffset() + paraStyle.getIndentLeft(), column.getWidth() - paraStyle.getIndentLeft());
					paraContent.addContent(new StringContent((int) bounds.getWidth(), (int) bounds.getHeight(), BULLET), paraStyle.getLineSpacing());
					paraContent.addHorizontalSpace(paraStyle.getIndentHanging(), paraStyle.getLineSpacing());
				}
			}
		}

		if (properties != null && p.getContent().size() == 0) {
			column.addVerticalSpace((int) paraStyle.getStringBoxSize("").getHeight());
		} else {
			paraContent.setBuffered(column.isBuffered());

			iterateContentParts(p, paraContent);

			paraContent.setBuffered(false);
		}

		column.addRow(paraContent);
		column.addVerticalSpace(paraStyle.getSpaceAfter());

		renderer.renderColumn(column);

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

		column.addAction(paraStyle.getHAlignment());

		if (run.getRPr() != null && run.getContent().size() == 0) {
			column.addVerticalSpace((int) paraStyle.getStringBoxSize("").getHeight());
		} else {
			iterateContentParts(run, column);
		}
	}

	private void processBreak(Br br, Column column) {
		if (br.getType() == null) {
			// paragraph break
			renderer.renderColumn(column);
		} else {
			switch (br.getType()) {
				case PAGE:
					renderer.renderColumn(column);
					renderer.nextPage(layout.getWidth(), layout.getHeight());
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

				renderer.renderColumn(column);
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

		CTTblCellMar tableMargins = table.getTblPr().getTblCellMar();
		int topMargin = 0;
		int rightMargin = 0;
		int bottomMargin = 0;
		int leftMargin = 0;

		if (tableMargins != null) {
			topMargin = getValue(tableMargins.getTop());
			rightMargin = getValue(tableMargins.getRight());
			bottomMargin = getValue(tableMargins.getBottom());
			leftMargin = getValue(tableMargins.getLeft());
		}

		for (Object tblObj : table.getContent()) {
			if (tblObj instanceof Tr) {
				Tr tableRow = (Tr) tblObj;
				int xOffset = column.getXOffset();
				int col = 0;
				List<Column> cells = new ArrayList<>();
				int minHeight = getMinRowHeight(tableRow);

				for (Object rowObj : tableRow.getContent()) {
					if (rowObj instanceof JAXBElement) {
						JAXBElement<?> element = (JAXBElement<?>) rowObj;

						if (element.getDeclaredType().equals(Tc.class)) {
							Tc tableCell = (Tc) element.getValue();
							int width = columnWidths.get(col);
							VAlignment vAlignment = VAlignment.TOP;
							Color fill = null;
							Border top = null;
							Border right = null;
							Border bottom = null;
							Border left = null;

							TcMar cellMargins = tableCell.getTcPr().getTcMar();

							if (cellMargins != null) {
								topMargin = getValue(cellMargins.getTop(), topMargin);
								rightMargin = getValue(cellMargins.getRight(), rightMargin);
								bottomMargin = getValue(cellMargins.getBottom(), bottomMargin);
								leftMargin = getValue(cellMargins.getLeft(), leftMargin);
							}

							if (tableCell.getTcPr().getVAlign() != null) {
								switch (tableCell.getTcPr().getVAlign().getVal()) {
									case BOTTOM:
										vAlignment = VAlignment.BOTTOM;
									break;
									case CENTER:
										vAlignment = VAlignment.CENTER;
									break;
									default:
										// default to TOP vertical alignment
								}
							}

							// Horizontal cell merge
							if (tableCell.getTcPr().getGridSpan() != null) {
								int mergeCols = tableCell.getTcPr().getGridSpan().getVal().intValue();

								for (int i = 0; i < mergeCols - 1; i++) {
									width += columnWidths.get(++col);
								}
							}

							if (tableCell.getTcPr().getShd() != null) {
								fill = getColor(tableCell.getTcPr().getShd().getFill(), null);
							}

							if (tableCell.getTcPr().getTcBorders() != null) {
								TcBorders borders = tableCell.getTcPr().getTcBorders();

								top = getBorder(borders.getTop());
								right = getBorder(borders.getRight());
								bottom = getBorder(borders.getBottom());
								left = getBorder(borders.getLeft());
							}

							Column cell = new Column(xOffset, width, vAlignment, fill, top, right, bottom, left);
							cell.addVerticalSpace(topMargin);

							// Add actual cell contents to an inner column to account for margins
							Column cellContent = new Column(xOffset + leftMargin, width - leftMargin - rightMargin);

							cellContent.setBuffered(true);

							iterateContentParts(tableCell, cellContent);

							cellContent.setBuffered(false);
							cell.addRow(cellContent);
							cell.addVerticalSpace(bottomMargin);
							xOffset += cell.getWidth();
							++col;
							cells.add(cell);
						}
					} else {
						LOG.debug("Unhandled row object " + rowObj.getClass());
					}
				}

				column.addRow(new TableRow(minHeight, cells));
				renderer.renderColumn(column);
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

					renderer.renderImage(
						new ImageContent(width, height, relationshipPart, anchor.getGraphic().getGraphicData().getPic().getBlipFill().getBlip().getEmbed()),
						x,
						y
					);
				} else {
					processGraphic(anchor.getExtent(), anchor.getGraphic().getGraphicData(), column);
				}
			} else {
				LOG.debug("Unhandled drawing object " + obj.getClass());
			}
		}
	}

	private void processHyperlink(Hyperlink link, Column column) {
		iterateContentParts(link, column);
	}

	private void processGraphic(CTPositiveSize2D extent, GraphicData graphicData, Column column) {
		int width = (int) extent.getCx() / EMU_DIVISOR;
		int height = (int) extent.getCy() / EMU_DIVISOR;

		// TODO: handle null pic reference
		if (graphicData.getPic() != null) {
			String rId = graphicData.getPic().getBlipFill().getBlip().getEmbed();

			// TODO: Add support for external reference
			if (!rId.isEmpty()) {
				column.addContentForced(new ImageContent(width, height, relationshipPart, rId));
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
					newStyle.setLineSpacing(getValue(spacing.getLine()));
				}

				if (spacing.getBefore() != null) {
					newStyle.setSpaceBefore(getValue(spacing.getBefore()));
				}

				if (spacing.getAfter() != null) {
					newStyle.setSpaceAfter(getValue(spacing.getAfter()));
				}
			}

			Ind indent = properties.getInd();

			if (indent != null) {
				newStyle.setIndentLeft(getValue(indent.getLeft()));
				newStyle.setIndentRight(getValue(indent.getRight()));
				newStyle.setIndentHanging(getValue(indent.getHanging()));
			}

			if (properties.getJc() != null) {
				switch (properties.getJc().getVal()) {
					case RIGHT:
						newStyle.setHAlignment(HAlignment.RIGHT);
					break;
					case CENTER:
						newStyle.setHAlignment(HAlignment.CENTER);
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
			Color newColor = getColor(strColor, Color.BLACK);

			if (!newColor.equals(baseStyle.getColor())) {
				newStyle.setColor(newColor);
			}
		}

		return newStyle;
	}

	private Color getColor(String strColor, Color defaultColor) {
		if (strColor == null || strColor.equals("auto")) {
			return defaultColor;
		} else {
			String hex = StringUtils.leftPad(strColor, 6, '0');

			return new Color(
				Integer.valueOf(hex.substring(0, 2), 16),
				Integer.valueOf(hex.substring(2, 4), 16),
				Integer.valueOf(hex.substring(4, 6), 16)
			);
		}
	}

	private Border getBorder(CTBorder ctBorder) {
		Border border = null;

		if (ctBorder != null) {
			// TODO: support other border styles
			int size = getValue(ctBorder.getSz());

			if (size > 0) {
				border = new Border(getColor(ctBorder.getColor(), Color.BLACK), size / 8 * 20, BorderStyle.SINGLE);
			}
		}

		return border;
	}

	private int getMinRowHeight(Tr tableRow) {
		if (tableRow.getTrPr() != null) {
			for (JAXBElement<?> element : tableRow.getTrPr().getCnfStyleOrDivIdOrGridBefore()) {
				if (element.getValue() instanceof CTHeight) {
					CTHeight height = (CTHeight) element.getValue();

					return getValue(height.getVal());
				}
			}
		}

		return 0;
	}

	private int getValue(Integer i) {
		return (i == null) ? 0 : i;
	}

	private int getValue(BigInteger bi) {
		return getValue(bi, 0);
	}

	private int getValue(BigInteger bi, int defaultValue) {
		return (bi == null) ? defaultValue : bi.intValue();
	}

	private int getValue(TblWidth width) {
		return getValue(width, 0);
	}

	private int getValue(TblWidth width, int defaultValue) {
		return (width == null) ? defaultValue : getValue(width.getW(), defaultValue);
	}
}
