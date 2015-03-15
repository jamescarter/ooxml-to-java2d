package documentconverter.reader;

import java.io.File;
import java.util.ArrayDeque;
import java.util.Deque;
import java.util.List;

import org.apache.log4j.Logger;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.SectPr.PgSz;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import documentconverter.renderer.Page;
import documentconverter.renderer.Renderer;

public class DocxReader implements Reader {
	private static final Logger LOG = Logger.getLogger(DocxReader.class);
	private Renderer renderer;
	private File docx;
	private Deque<PageLayout> layouts = new ArrayDeque<>();
	private PageLayout layout;
	private Page page;

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

		MainDocumentPart main = word.getMainDocumentPart();

		setPageLayouts(main);
		createPageFromNextLayout();
		iterateContentParts(main);
	}

	private void setPageLayouts(MainDocumentPart main) throws ReaderException {
		List<Object> sections;

		try {
			// Page settings are stored in the last paragraph of each page
			// Retrieve all the sections and create the pages
			sections = main.getJAXBNodesViaXPath("//w:sectPr", false);
		} catch (XPathBinderAssociationIsPartialException | JAXBException e) {
			throw new ReaderException("Error retrieving sections", e);
		}

		for (Object obj : sections) {
			SectPr sectPr = (SectPr) obj;

			layouts.add(createPageLayout(sectPr));
		}
	}

	private PageLayout createPageLayout(SectPr sectPr) {
		PgSz size = sectPr.getPgSz();

		return new PageLayout(size.getW().intValue(), size.getH().intValue());
	}

	private void createPageFromNextLayout() {
		layout = layouts.removeFirst();
		createPageFromLayout();
	}

	private void createPageFromLayout() {
		page = renderer.addPage(layout.getWidth(), layout.getHeight());
	}

	public void iterateContentParts(ContentAccessor ca) {
        for (Object obj: ca.getContent()) {
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
		iterateContentParts(p);

		PPr ppr = p.getPPr();

		if (ppr.getSectPr() != null) {
			// The presence of SectPr indicates the next part should be started on a new page with a different layout
			createPageFromNextLayout();
		}		
	}

	private void processTextRun(R r) {
		iterateContentParts(r);
	}

	private void processText(Text value) {
		page.drawString(value.getValue(), 10, 100);
	}
}
