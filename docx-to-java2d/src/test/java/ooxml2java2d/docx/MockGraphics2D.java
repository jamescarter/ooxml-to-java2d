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
import java.awt.Composite;
import java.awt.Font;
import java.awt.FontMetrics;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.GraphicsConfiguration;
import java.awt.Image;
import java.awt.Paint;
import java.awt.Rectangle;
import java.awt.RenderingHints;
import java.awt.RenderingHints.Key;
import java.awt.Shape;
import java.awt.Stroke;
import java.awt.font.FontRenderContext;
import java.awt.font.GlyphVector;
import java.awt.geom.AffineTransform;
import java.awt.image.BufferedImage;
import java.awt.image.BufferedImageOp;
import java.awt.image.ImageObserver;
import java.awt.image.RenderedImage;
import java.awt.image.renderable.RenderableImage;
import java.text.AttributedCharacterIterator;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class MockGraphics2D extends Graphics2D {
	private List<Object> actions = new ArrayList<>();
	private int width;
	private int height;

	public MockGraphics2D(int width, int height) {
		this.width = width;
		this.height = height;
	}

	public int getWidth() {
		return width;
	}

	public int getHeight() {
		return height;
	}

	public List<Object> getActions(Class<?> ... classes) {
		List<Object> as = new ArrayList<>();

		for (Object action : actions) {
			for (Class<?> clazz : classes) {
				if (clazz.isInstance(action)) {
					as.add(action);
					break;
				}
			}
		}

		return as;
	}

	public <T> List<T> getActions(Class<T> clazz) {
		List<T> list = new ArrayList<>();

		for (Object action : actions) {
			if (clazz.isInstance(action)) {
				list.add((T) action);
			}
		}

		return list;
	}

	@Override
	public void addRenderingHints(Map<?, ?> hints) {

	}

	@Override
	public void clip(Shape s) {

	}

	@Override
	public void draw(Shape s) {

	}

	@Override
	public void drawGlyphVector(GlyphVector g, float x, float y) {

	}

	@Override
	public boolean drawImage(Image img, AffineTransform xform, ImageObserver obs) {
		return false;
	}

	@Override
	public void drawImage(BufferedImage img, BufferedImageOp op, int x, int y) {

	}

	@Override
	public void drawRenderableImage(RenderableImage img, AffineTransform xform) {

	}

	@Override
	public void drawRenderedImage(RenderedImage img, AffineTransform xform) {

	}

	@Override
	public void drawString(String str, int x, int y) {
		actions.add(new DrawStringAction(str, x, y));
	}

	@Override
	public void drawString(String str, float x, float y) {

	}

	@Override
	public void drawString(AttributedCharacterIterator iterator, int x, int y) {

	}

	@Override
	public void drawString(AttributedCharacterIterator iterator, float x, float y) {

	}

	@Override
	public void fill(Shape s) {

	}

	@Override
	public Color getBackground() {
		return null;
	}

	@Override
	public Composite getComposite() {
		return null;
	}

	@Override
	public GraphicsConfiguration getDeviceConfiguration() {
		return null;
	}

	@Override
	public FontRenderContext getFontRenderContext() {
		return null;
	}

	@Override
	public Paint getPaint() {
		return null;
	}

	@Override
	public Object getRenderingHint(Key hintKey) {
		return null;
	}

	@Override
	public RenderingHints getRenderingHints() {
		return null;
	}

	@Override
	public Stroke getStroke() {
		return null;
	}

	@Override
	public AffineTransform getTransform() {
		return null;
	}

	@Override
	public boolean hit(Rectangle rect, Shape s, boolean onStroke) {
		return false;
	}

	@Override
	public void rotate(double theta) {

	}

	@Override
	public void rotate(double theta, double x, double y) {

	}

	@Override
	public void scale(double sx, double sy) {

	}

	@Override
	public void setBackground(Color color) {

	}

	@Override
	public void setComposite(Composite comp) {

	}

	@Override
	public void setPaint(Paint paint) {

	}

	@Override
	public void setRenderingHint(Key hintKey, Object hintValue) {

	}

	@Override
	public void setRenderingHints(Map<?, ?> hints) {

	}

	@Override
	public void setStroke(Stroke s) {

	}

	@Override
	public void setTransform(AffineTransform Tx) {

	}

	@Override
	public void shear(double shx, double shy) {

	}

	@Override
	public void transform(AffineTransform Tx) {

	}

	@Override
	public void translate(int x, int y) {

	}

	@Override
	public void translate(double tx, double ty) {

	}

	@Override
	public void clearRect(int x, int y, int width, int height) {

	}

	@Override
	public void clipRect(int x, int y, int width, int height) {

	}

	@Override
	public void copyArea(int x, int y, int width, int height, int dx, int dy) {

	}

	@Override
	public Graphics create() {
		return null;
	}

	@Override
	public void dispose() {

	}

	@Override
	public void drawArc(int x, int y, int width, int height, int startAngle, int arcAngle) {

	}

	@Override
	public boolean drawImage(Image img, int x, int y, ImageObserver observer) {
		return false;
	}

	@Override
	public boolean drawImage(Image img, int x, int y, Color bgcolor, ImageObserver observer) {
		return false;
	}

	@Override
	public boolean drawImage(Image img, int x, int y, int width, int height, ImageObserver observer) {
		actions.add(new DrawImageAction(x, y, width, height));

		return false;
	}

	@Override
	public boolean drawImage(Image img, int x, int y, int width, int height, Color bgcolor, ImageObserver observer) {
		return false;
	}

	@Override
	public boolean drawImage(Image img, int dx1, int dy1, int dx2, int dy2, int sx1, int sy1, int sx2, int sy2, ImageObserver observer) {
		return false;
	}

	@Override
	public boolean drawImage(Image img, int dx1, int dy1, int dx2, int dy2, int sx1, int sy1, int sx2, int sy2, Color bgcolor, ImageObserver observer) {
		return false;
	}

	@Override
	public void drawLine(int x1, int y1, int x2, int y2) {

	}

	@Override
	public void drawOval(int x, int y, int width, int height) {

	}

	@Override
	public void drawPolygon(int[] xPoints, int[] yPoints, int nPoints) {

	}

	@Override
	public void drawPolyline(int[] xPoints, int[] yPoints, int nPoints) {

	}

	@Override
	public void drawRoundRect(int x, int y, int width, int height, int arcWidth, int arcHeight) {

	}

	@Override
	public void fillArc(int x, int y, int width, int height, int startAngle, int arcAngle) {

	}

	@Override
	public void fillOval(int x, int y, int width, int height) {

	}

	@Override
	public void fillPolygon(int[] xPoints, int[] yPoints, int nPoints) {

	}

	@Override
	public void fillRect(int x, int y, int width, int height) {

	}

	@Override
	public void fillRoundRect(int x, int y, int width, int height, int arcWidth, int arcHeight) {

	}

	@Override
	public Shape getClip() {

		return null;
	}

	@Override
	public Rectangle getClipBounds() {

		return null;
	}

	@Override
	public Color getColor() {

		return null;
	}

	@Override
	public Font getFont() {

		return null;
	}

	@Override
	public FontMetrics getFontMetrics(Font f) {

		return null;
	}

	@Override
	public void setClip(Shape clip) {

	}

	@Override
	public void setClip(int x, int y, int width, int height) {

	}

	@Override
	public void setColor(Color color) {
		actions.add(color);
	}

	@Override
	public void setFont(Font font) {
		actions.add(font);
	}

	@Override
	public void setPaintMode() {

	}

	@Override
	public void setXORMode(Color color) {

	}
}
