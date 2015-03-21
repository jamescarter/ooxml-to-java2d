package documentconverter.reader;

import java.io.File;
import java.util.ArrayDeque;
import java.util.Deque;

import org.apache.log4j.Logger;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.PPrBase.PStyle;
import org.docx4j.wml.SectPr.PgMar;
import org.docx4j.wml.SectPr.PgSz;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;

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
	private FontConfig font = new FontConfig();

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
		xOffset = layout.getTopMargin();
		yOffset = layout.getLeftMargin();
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

			setStyle(style.getRPr());
		}

		iterateContentParts(p);

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
		page.drawString(text.getValue(), xOffset, yOffset);
		xOffset += font.getWidth(text.getValue());
	}

	private void setStyle(RPr runProperties) {
		if (runProperties.getRFonts() != null) {
			font.setFontName(runProperties.getRFonts().getAscii());
		}

		if (runProperties.getSz() != null) {
			font.setSize(runProperties.getSz().getVal().intValue() / 2);
			page.setFontConfig(font);
		}
	}
}
