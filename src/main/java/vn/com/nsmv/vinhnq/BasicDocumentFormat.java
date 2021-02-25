package vn.com.nsmv.vinhnq;

import com.artofsolving.jodconverter.DocumentFormatRegistry;
import com.artofsolving.jodconverter.DocumentFamily;
import com.artofsolving.jodconverter.DocumentFormat;

public class BasicDocumentFormat {
	public static DocumentFormat getFormatByFileExtension(String fileExtension) {
		if (null == fileExtension) {
			return null;
		}
		switch (fileExtension.toLowerCase()) {
		case "pdf":
			return pdf();
		case "swf":
			return swf();
		case "xhtml":
			return xhtml();
		case "html":
			return html();
		case "odt":
			return odt();
		case "sxw":
			return sxw();
		case "doc":
			return doc();
		case "rtf":
			return rtf();
		case "wpd":
			return wpd();
		case "txt":
			return txt();
		case "wikitext":
			return wikitext();
		case "ods":
			return ods();
		case "sxc":
			return sxc();
		case "xls":
			return xls();
		case "xlsx":
			return xlsx();
		case "csv": 
			return csv();
		case "tsv":
			return tsv();
		case "odp":
			return odp();
		case "sxi":
			return sxi();
		case "ppt":
			return ppt();
		case "odg":
			return odg();
		case "svg":
			return svg();
		default:
			return null;
		}

	}

	public static DocumentFormat pdf() {
		final DocumentFormat pdf = new DocumentFormat("Portable Document Format", "application/pdf", "pdf");
		pdf.setExportFilter(DocumentFamily.DRAWING, "draw_pdf_Export");
		pdf.setExportFilter(DocumentFamily.PRESENTATION, "impress_pdf_Export");
		pdf.setExportFilter(DocumentFamily.SPREADSHEET, "calc_pdf_Export");
		pdf.setExportFilter(DocumentFamily.TEXT, "writer_pdf_Export");
		return pdf;
	}

	public static DocumentFormat swf() {
		final DocumentFormat swf = new DocumentFormat("Macromedia Flash", "application/x-shockwave-flash", "swf");
		swf.setExportFilter(DocumentFamily.DRAWING, "draw_flash_Export");
		swf.setExportFilter(DocumentFamily.PRESENTATION, "impress_flash_Export");
		return swf;
	}

	public static DocumentFormat xhtml() {
		final DocumentFormat xhtml = new DocumentFormat("XHTML", "application/xhtml+xml", "xhtml");
		xhtml.setExportFilter(DocumentFamily.PRESENTATION, "XHTML Impress File");
		xhtml.setExportFilter(DocumentFamily.SPREADSHEET, "XHTML Calc File");
		xhtml.setExportFilter(DocumentFamily.TEXT, "XHTML Writer File");
		return xhtml;
	}

	// HTML is treated as Text when supplied as input, but as an output it is also
	// available for exporting Spreadsheet and Presentation formats
	public static DocumentFormat html() {
		final DocumentFormat html = new DocumentFormat("HTML", DocumentFamily.TEXT, "text/html", "html");
		html.setExportFilter(DocumentFamily.PRESENTATION, "impress_html_Export");
		html.setExportFilter(DocumentFamily.SPREADSHEET, "HTML (StarCalc)");
		html.setExportFilter(DocumentFamily.TEXT, "HTML (StarWriter)");
		return html;
	}

	public static DocumentFormat odt() {
		final DocumentFormat odt = new DocumentFormat("OpenDocument Text", DocumentFamily.TEXT,
				"application/vnd.oasis.opendocument.text", "odt");
		odt.setExportFilter(DocumentFamily.TEXT, "writer8");
		return odt;
	}

	public static DocumentFormat sxw() {
		final DocumentFormat sxw = new DocumentFormat("OpenOffice.org 1.0 Text Document", DocumentFamily.TEXT,
				"application/vnd.sun.xml.writer", "sxw");
		sxw.setExportFilter(DocumentFamily.TEXT, "StarOffice XML (Writer)");
		return sxw;
	}

	public static DocumentFormat doc() {
		final DocumentFormat doc = new DocumentFormat("Microsoft Word", DocumentFamily.TEXT, "application/msword",
				"doc");
		doc.setExportFilter(DocumentFamily.TEXT, "MS Word 97");
		return doc;
	}

	public static DocumentFormat docx() {
		final DocumentFormat docx = new DocumentFormat("Microsoft Word 2007 XML", DocumentFamily.TEXT,
				"application/vnd.openxmlformats-officedocument.wordprocessingml.document", "docx");
		return docx;
	}

	public static DocumentFormat rtf() {
		final DocumentFormat rtf = new DocumentFormat("Rich Text Format", DocumentFamily.TEXT, "text/rtf", "rtf");
		rtf.setExportFilter(DocumentFamily.TEXT, "Rich Text Format");
		return rtf;
	}

	public static DocumentFormat wpd() {
		final DocumentFormat wpd = new DocumentFormat("WordPerfect", DocumentFamily.TEXT, "application/wordperfect",
				"wpd");
		return wpd;
	}

	public static DocumentFormat txt() {
		final DocumentFormat txt = new DocumentFormat("Plain Text", DocumentFamily.TEXT, "text/plain", "txt");
		// default to "Text (encoded)" UTF8,CRLF to prevent OOo from trying to display
		// the "ASCII Filter Options" dialog
		txt.setImportOption("FilterName", "Text (encoded)");
		txt.setImportOption("FilterOptions", "UTF8,CRLF");
		txt.setExportFilter(DocumentFamily.TEXT, "Text (encoded)");
		txt.setExportOption(DocumentFamily.TEXT, "FilterOptions", "UTF8,CRLF");
		return txt;
	}

	public static DocumentFormat wikitext() {
		final DocumentFormat wikitext = new DocumentFormat("MediaWiki wikitext", "text/x-wiki", "wiki");
		wikitext.setExportFilter(DocumentFamily.TEXT, "MediaWiki");
		return wikitext;
	}

	public static DocumentFormat ods() {
		final DocumentFormat ods = new DocumentFormat("OpenDocument Spreadsheet", DocumentFamily.SPREADSHEET,
				"application/vnd.oasis.opendocument.spreadsheet", "ods");
		ods.setExportFilter(DocumentFamily.SPREADSHEET, "calc8");
		return ods;
	}

	public static DocumentFormat sxc() {
		final DocumentFormat sxc = new DocumentFormat("OpenOffice.org 1.0 Spreadsheet", DocumentFamily.SPREADSHEET,
				"application/vnd.sun.xml.calc", "sxc");
		sxc.setExportFilter(DocumentFamily.SPREADSHEET, "StarOffice XML (Calc)");
		return sxc;
	}

	public static DocumentFormat xls() {
		final DocumentFormat xls = new DocumentFormat("Microsoft Excel", DocumentFamily.SPREADSHEET,
				"application/vnd.ms-excel", "xls");
		xls.setExportFilter(DocumentFamily.SPREADSHEET, "MS Excel 97");
		return xls;
	}

	public static DocumentFormat xlsx() {
		final DocumentFormat xlsx = new DocumentFormat("Microsoft Excel 2007 XML", DocumentFamily.SPREADSHEET,
				"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx");
		return xlsx;
	}

	public static DocumentFormat csv() {
		final DocumentFormat csv = new DocumentFormat("CSV", DocumentFamily.SPREADSHEET, "text/csv", "csv");
		csv.setImportOption("FilterName", "Text - txt - csv (StarCalc)");
		csv.setImportOption("FilterOptions", "44,34,0"); // Field Separator: ','; Text Delimiter: '"'
		csv.setExportFilter(DocumentFamily.SPREADSHEET, "Text - txt - csv (StarCalc)");
		csv.setExportOption(DocumentFamily.SPREADSHEET, "FilterOptions", "44,34,0");
		return csv;
	}

	public static DocumentFormat tsv() {
		final DocumentFormat tsv = new DocumentFormat("Tab-separated Values", DocumentFamily.SPREADSHEET,
				"text/tab-separated-values", "tsv");
		tsv.setImportOption("FilterName", "Text - txt - csv (StarCalc)");
		tsv.setImportOption("FilterOptions", "9,34,0"); // Field Separator: '\t'; Text Delimiter: '"'
		tsv.setExportFilter(DocumentFamily.SPREADSHEET, "Text - txt - csv (StarCalc)");
		tsv.setExportOption(DocumentFamily.SPREADSHEET, "FilterOptions", "9,34,0");
		return tsv;
	}

	public static DocumentFormat odp() {
		final DocumentFormat odp = new DocumentFormat("OpenDocument Presentation", DocumentFamily.PRESENTATION,
				"application/vnd.oasis.opendocument.presentation", "odp");
		odp.setExportFilter(DocumentFamily.PRESENTATION, "impress8");
		return odp;
	}

	public static DocumentFormat sxi() {
		final DocumentFormat sxi = new DocumentFormat("OpenOffice.org 1.0 Presentation", DocumentFamily.PRESENTATION,
				"application/vnd.sun.xml.impress", "sxi");
		sxi.setExportFilter(DocumentFamily.PRESENTATION, "StarOffice XML (Impress)");
		return sxi;
	}

	public static DocumentFormat ppt() {
		final DocumentFormat ppt = new DocumentFormat("Microsoft PowerPoint", DocumentFamily.PRESENTATION,
				"application/vnd.ms-powerpoint", "ppt");
		ppt.setExportFilter(DocumentFamily.PRESENTATION, "MS PowerPoint 97");
		return ppt;
	}

	public static DocumentFormat pptx() {
		final DocumentFormat pptx = new DocumentFormat("Microsoft PowerPoint 2007 XML", DocumentFamily.PRESENTATION,
				"application/vnd.openxmlformats-officedocument.presentationml.presentation", "pptx");
		return pptx;
	}

	public static DocumentFormat odg() {
		final DocumentFormat odg = new DocumentFormat("OpenDocument Drawing", DocumentFamily.DRAWING,
				"application/vnd.oasis.opendocument.graphics", "odg");
		odg.setExportFilter(DocumentFamily.DRAWING, "draw8");
		return odg;
	}

	public static DocumentFormat svg() {
		final DocumentFormat svg = new DocumentFormat("Scalable Vector Graphics", "image/svg+xml", "svg");
		svg.setExportFilter(DocumentFamily.DRAWING, "draw_svg_Export");
		return svg;
	}
}
