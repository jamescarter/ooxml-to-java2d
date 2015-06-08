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

package ooxml2java2d.docx.internal;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.RenderingHints;
import java.io.IOException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import ooxml2java2d.GraphicsBuilder;
import ooxml2java2d.docx.internal.content.BlankRow;
import ooxml2java2d.docx.internal.content.Border;
import ooxml2java2d.docx.internal.content.Column;
import ooxml2java2d.docx.internal.content.Content;
import ooxml2java2d.docx.internal.content.ImageContent;
import ooxml2java2d.docx.internal.content.Line;
import ooxml2java2d.docx.internal.content.Row;
import ooxml2java2d.docx.internal.content.StringContent;
import ooxml2java2d.docx.internal.content.TableRow;

public class GraphicsRenderer {
	private static final Logger LOG = LoggerFactory.getLogger(GraphicsRenderer.class);
	private GraphicsBuilder builder;
	private PageInitiationAdapter initiation;
	private Graphics2D g2;
	private HAlignment hAlignment = HAlignment.LEFT;
	private int yOffset;
	private int pageWidth;
	private int pageHeight;
	private int endPosition;
	private int tableRowNesting = 0;

	public GraphicsRenderer(GraphicsBuilder builder, PageInitiationAdapter initiation) {
		this.builder = builder;
		this.initiation = initiation;
	}

	public int getYOffset() {
		return yOffset;
	}

	public void setYOffset(int yOffset) {
		this.yOffset = yOffset;
	}

	public void setEndPosition(int endPosition) {
		this.endPosition = endPosition;
	}

	public void nextPage(int pageWidth, int pageHeight) {
		this.yOffset = 0;
		this.pageWidth = pageWidth;
		this.pageHeight = pageHeight;
		this.endPosition = pageHeight;

		Graphics2D nextG2 = builder.nextPage(pageWidth, pageHeight);

		nextG2.setBackground(Color.WHITE);
		nextG2.clearRect(0, 0, pageWidth, pageHeight);
		nextG2.setColor(Color.BLACK);
		nextG2.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
		nextG2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_LCD_HRGB);
		nextG2.setRenderingHint(RenderingHints.KEY_TEXT_LCD_CONTRAST, 250);

		// Use the settings from the previous 'page'
		if (g2 != null) {
			nextG2.setFont(g2.getFont());
			nextG2.setColor(g2.getColor());
		}

		g2 = nextG2;
		initiation.initiatePage();
	}

	public void renderColumn(Column column) {
		renderColumn(column, false, column.getContentHeight());
	}

	public void renderImage(ImageContent ic, int x, int y) {
		try {
			Image image = ic.getImage();

			if (image == null) {
				LOG.error("Error creating image for " + ic.getRelationshipId());
			} else {
				g2.drawImage(
					image,
					x,
					y,
					ic.getWidth(),
					ic.getHeight(),
					null
				);
			}
		} catch (IOException ioe) {
			LOG.error("Error reading image", ioe);
		}
	}

	private void renderColumn(Column column, boolean delayPageCreation, int contentHeight) {
		if (column.isBuffered()) {
			return;
		}

		Color origColor = g2.getColor();

		if (column.getFill() != null) {
			g2.setColor(column.getFill());
			g2.fillRect(column.getXOffset(), yOffset, column.getWidth(), contentHeight);
		}

		if (column.getTopBorder() != null) {
			Border top = column.getTopBorder();
			g2.setColor(top.getColor());
			g2.setStroke(new BasicStroke((float) top.getSize()));
			g2.drawLine(column.getXOffset(), yOffset, column.getXOffset() + column.getWidth(), yOffset);
		}

		if (column.getRightBorder() != null) {
			Border right = column.getRightBorder();
			g2.setColor(right.getColor());
			g2.setStroke(new BasicStroke((float) right.getSize()));
			g2.drawLine(column.getXOffset() + column.getWidth(), yOffset, column.getXOffset() + column.getWidth(), yOffset + contentHeight);
		}

		if (column.getBottomBorder() != null) {
			Border bottom = column.getBottomBorder();
			g2.setColor(bottom.getColor());
			g2.setStroke(new BasicStroke((float) bottom.getSize()));
			g2.drawLine(column.getXOffset(), yOffset + contentHeight, column.getXOffset() + column.getWidth(), yOffset + contentHeight);
		}

		if (column.getLeftBorder() != null) {
			Border left = column.getLeftBorder();
			g2.setColor(left.getColor());
			g2.setStroke(new BasicStroke((float) left.getSize()));
			g2.drawLine(column.getXOffset(), yOffset, column.getXOffset(), yOffset + contentHeight);
		}

		g2.setColor(origColor);

		for (Row row : column.getRows()) {
			if (row instanceof Line) {
				// Check if this line will fit onto the current page, otherwise create a new page
				if (yOffset + row.getContentHeight() > endPosition) {
					if (delayPageCreation) {
						return;
					} else {
						nextPage();
					}
				}

				renderLine((Line) row, column.getXOffset());
			} else if (row instanceof BlankRow) {
				yOffset += row.getContentHeight();
			} else if (row instanceof TableRow) {
				++tableRowNesting;
				renderTableRow((TableRow) row);
				--tableRowNesting;

				// If this is the top-level table row and there's still content, create a new page and output it
				if (tableRowNesting > 0) {
					return;
				} else if (row.getContentHeight() > 0) {
					nextPage();
					renderTableRow((TableRow) row);
				}
			} else {
				LOG.debug("Unhandled row object " + row.getClass());
			}

			column.removeRow(row);
		}
	}

	private void renderLine(Line line, int initialXOffset) {
		yOffset += line.getContentHeight();
		int xOffset = initialXOffset;

		switch (hAlignment) {
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

					g2.drawString(sc.getText(), xOffset, yOffset);
				} else if (obj instanceof ImageContent) {
					ImageContent di = (ImageContent) obj;

					renderImage(
						di,
						xOffset,
						yOffset - di.getHeight()
					);
				} else {
					yOffset += content.getHeight();
				}

				xOffset += content.getWidth();
			} else if (obj instanceof HAlignment) {
				hAlignment = (HAlignment) obj;
			} else if (obj instanceof Color) {
				g2.setColor((Color) obj);
			} else if (obj instanceof FontConfig) {
				FontConfig fc = (FontConfig) obj;

				g2.setFont(fc.getFont());
			} else {
				LOG.debug("Unhandled render object " + obj.getClass());
			}
		}
	}

	private void renderTableRow(TableRow row) {
		int start = yOffset; // start every column from the same position
		int maxYOffset = start;
		int contentHeight = row.getContentHeight();

		// Render as much content from each column onto the current page as possible
		for (Column cell : row.getColumns()) {
			yOffset = start;
			renderColumn(cell, true, contentHeight);
			maxYOffset = Math.max(maxYOffset, yOffset);
		}

		yOffset = maxYOffset;
	}

	private void nextPage() {
		nextPage(pageWidth, pageHeight);
	}
}
