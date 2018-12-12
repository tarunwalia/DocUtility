package com.mywork.services;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.HeaderStories;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;;

/**
 * @author tarun
 *
 */
public class DocUtility {
	private WordExtractor wordExtractor;
	private HWPFDocument doc;
	private HWPFDocument targetDoc;
	private XWPFDocument xDoc;
	private XWPFDocument targetxDoc;
	private int pageCount;
	private int targetPageCount;

	/**
	 * DocUtility constructor is used to initialize the doc/docx object
	 * 
	 * @param filePath
	 * @throws IOException
	 * @throws InvalidFormatException
	 */
	public DocUtility(String filePath) throws IOException, InvalidFormatException {
		String fileExt = getFileExtension(filePath);
		File file = new File(filePath);
		FileInputStream fis = new FileInputStream(file);
		if ("doc".equalsIgnoreCase(fileExt)) {
			doc = new HWPFDocument(fis);
			getPageCount();
		} else if ("docx".equalsIgnoreCase(fileExt)) {
			xDoc = new XWPFDocument(OPCPackage.open(fis));
			getPageCount();
		} else {
			throw new FileNotFoundException("File extension is not as expected. It should be doc or docx.");
		}
	}

	/**
	 * this method will return the count of number of pages of input doc file
	 * 
	 * @return int
	 */
	public int getPageCount() {
		pageCount = (null != xDoc) ? xDoc.getProperties().getExtendedProperties().getUnderlyingProperties().getPages()
				: doc.getSummaryInformation().getPageCount();
		return pageCount;
		/*
		 * if (null != xDoc) { int pageC =
		 * xDoc.getProperties().getExtendedProperties().getUnderlyingProperties().
		 * getPages(); return pageC; } else { pageCount =
		 * doc.getSummaryInformation().getPageCount(); return pageCount; }
		 */
	}

	/**
	 * this method will return the doc text as string
	 * 
	 * @return String
	 */
	public String getText() {
		if (null != xDoc) {
			String text = new XWPFWordExtractor(xDoc).getText();
			return text;
		} else {
			wordExtractor = new WordExtractor(doc);
			String text = wordExtractor.getText();
			return text;
		}
	}

	/**
	 * this method will return the table elements as list for input file
	 * 
	 * @param tableNumber
	 * @return list
	 */
	public List<List<String>> getTableData(int tableNumber) {
		Iterator<IBodyElement> docIterator = xDoc.getBodyElementsIterator();
		String text = "";
		List<List<String>> list = new ArrayList<List<String>>();
		while (docIterator.hasNext()) {
			IBodyElement ele = docIterator.next();
			if ("TABLE".equalsIgnoreCase(ele.getElementType().name())) {
				List<XWPFTable> tableList = ele.getBody().getTables();
				XWPFTable xTable = tableList.get(tableNumber);
				for (int i = 0; i < xTable.getRows().size(); i++) {
					List<String> l = new ArrayList<String>();
					for (int j = 0; j < xTable.getRow(i).getTableCells().size(); j++) {
						text = xTable.getRow(i).getCell(j).getText().trim();
						l.add(text);
					}
					list.add(l);
				}
			}
		}
		return list;
	}

	/**
	 * this method is used to compare the source doc to target doc file
	 * 
	 * @param targetDoc
	 * @return boolean
	 * @throws IOException
	 * @throws InvalidFormatException
	 */
	public boolean compareDoc(String targetDoc) throws IOException, InvalidFormatException {
		boolean isDocMatching = false;
		String sourceDocText = getText();
		String targetDocText = getTargetDocText(targetDoc);
		if (pageCount == targetPageCount) {
			if (sourceDocText.equals(targetDocText)) {
				isDocMatching = true;
				return isDocMatching;
			}
		}
		return isDocMatching;
	}

	/**
	 * this method will search the input string in doc file
	 * 
	 * @param searchString
	 * @return boolean
	 */
	public boolean searchText(String searchString) {
		boolean textAvailable = false;
		String docText = getText();
		if (docText.toLowerCase().contains(searchString.toLowerCase())) {
			textAvailable = true;
			return textAvailable;
		}
		return textAvailable;
	}

	/*
	 * public void replaceText(String source, String target, String outputFilePath)
	 * throws IOException { for (XWPFParagraph p : xDoc.getParagraphs()) {
	 * List<XWPFRun> runs = p.getRuns(); if (runs != null) { for (XWPFRun r : runs)
	 * { String text = r.getText(0); if (text != null && text.contains(source)) {
	 * text = text.replace(source, target); r.setText(text, 0); } } } }
	 * 
	 * for (XWPFTable tbl : xDoc.getTables()) { for (XWPFTableRow row :
	 * tbl.getRows()) { for (XWPFTableCell cell : row.getTableCells()) { for
	 * (XWPFParagraph p : cell.getParagraphs()) { for (XWPFRun r : p.getRuns()) {
	 * String text = r.getText(0); if (text != null && text.contains("2011-2015")) {
	 * text = text.replace("2011-2015", "2009-2015"); r.setText(text, 0); } } } } }
	 * } xDoc.write(new FileOutputStream(outputFilePath)); }
	 */

	/**
	 * @param targetDocPath
	 * @return
	 * @throws IOException
	 * @throws InvalidFormatException
	 */
	private String getTargetDocText(String targetDocPath) throws IOException, InvalidFormatException {
		String fileExt = getFileExtension(targetDocPath);
		File file = new File(targetDocPath);
		FileInputStream fis = new FileInputStream(file);
		if ("doc".equalsIgnoreCase(fileExt)) {
			targetDoc = new HWPFDocument(fis);
			targetPageCount = targetDoc.getSummaryInformation().getPageCount();
			wordExtractor = new WordExtractor(targetDoc);
			String text = wordExtractor.getText();
			return text;
		} else if ("docx".equalsIgnoreCase(fileExt)) {
			targetxDoc = new XWPFDocument(OPCPackage.open(fis));
			targetPageCount = targetxDoc.getProperties().getExtendedProperties().getUnderlyingProperties().getPages();
			String text = new XWPFWordExtractor(targetxDoc).getText();
			return text;
		} else {
			throw new FileNotFoundException("File extension is not as expected. It should be doc or docx.");
		}
	}

	/**
	 * this method will return the header text of input file
	 * 
	 * @return String
	 */
	public String getHeaderText() {
		if (null != xDoc) {
			XWPFHeaderFooterPolicy xfPolicy = new XWPFHeaderFooterPolicy(xDoc);
			XWPFHeader xfHeader = xfPolicy.getDefaultHeader();
			if (null != xfHeader) {
				return xfHeader.getText();
			}
		} else {
			HeaderStories hStories = new HeaderStories(doc);
			String header = hStories.getHeader(this.pageCount);
			return header;
		}
		return null;
	}

	/**
	 * this method will return the footer text of input file
	 * 
	 * @return String
	 */
	public String getFooterText() {
		if (null != xDoc) {
			XWPFHeaderFooterPolicy xfPolicy = new XWPFHeaderFooterPolicy(xDoc);
			XWPFFooter xfFooter = xfPolicy.getDefaultFooter();
			if (null != xfFooter) {
				return xfFooter.getText();
			}
		} else {
			HeaderStories hStories = new HeaderStories(doc);
			String footer = hStories.getFooter(this.pageCount);
			return footer;
		}
		return null;
	}

	/**
	 * this method is used to get the file extension
	 * 
	 * @param filePath
	 * @return
	 */
	private String getFileExtension(String filePath) {
		String fileExtesion = filePath.substring(filePath.length() - 4);
		fileExtesion = fileExtesion.contains(".") ? filePath.substring(filePath.length() - 3) : fileExtesion;
		return fileExtesion;
	}

}
