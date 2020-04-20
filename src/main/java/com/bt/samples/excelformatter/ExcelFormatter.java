package com.bt.samples.excelformatter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFormatter {

	private String inputFilePath;

	private String outputFilePath;

	public ExcelFormatter(String inputFilePath, String outputFilePath) {
		this.inputFilePath = inputFilePath;
		this.outputFilePath = outputFilePath;
	}

	public void format() throws IOException, EncryptedDocumentException, InvalidFormatException {

		InputStream inputStream = null;
		OutputStream outputStream = null;
		File inputFile = null;
		File outputFile = null;
		try {
			inputFile = new File(inputFilePath);
			outputFile = new File(outputFilePath);

			inputStream = new FileInputStream(inputFile);
			Workbook workBook = WorkbookFactory.create(inputStream);

			formatSummaryPage(workBook);
			formatOtherPages(workBook);

			outputStream = new FileOutputStream(outputFile);
			workBook.write(outputStream);
			workBook.close();

		} finally {
			if (null != inputStream)
				inputStream.close();

			if (null != outputStream)
				outputStream.close();
		}
	}

	private void formatSummaryPage(Workbook workBook) {
		Sheet summarySheet = workBook.getSheet("SUMMARY");

		alignCellContentsToLeftInSheet(summarySheet, workBook);
		setAddressLinksToOtherSheets(summarySheet, workBook);
		autoSizeColumns(summarySheet, summarySheet.getRow(summarySheet.getFirstRowNum()), 0);
		formatSummaryHeaders(summarySheet, workBook);
	}

	private void formatSummaryHeader2(Sheet summarySheet, Workbook workBook) {
		Row header2 = summarySheet.getRow(8);
		Font summaryHeaderFont = workBook.createFont();
		summaryHeaderFont.setBold(true);
		summaryHeaderFont.setColor(IndexedColors.WHITE.getIndex());

		CellStyle summaryHeaderStyle2 = workBook.createCellStyle();
		summaryHeaderStyle2.setFont(summaryHeaderFont);
		summaryHeaderStyle2.setAlignment(HorizontalAlignment.LEFT);
		summaryHeaderStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		summaryHeaderStyle2.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());

		header2.forEach(cell -> cell.setCellStyle(summaryHeaderStyle2));
	}

	private void setAddressLinksToOtherSheets(Sheet summarySheet, Workbook workBook) {
		CreationHelper createHelper = workBook.getCreationHelper();

		Font addressLinkFont = workBook.createFont();
		addressLinkFont.setUnderline(Font.U_SINGLE);
		addressLinkFont.setColor(IndexedColors.DARK_BLUE.getIndex());

		CellStyle addressLinkStyle = workBook.createCellStyle();
		addressLinkStyle.setAlignment(HorizontalAlignment.LEFT);
		addressLinkStyle.setFont(addressLinkFont);

		if (summarySheet.getLastRowNum() > 0) {
			for (int i = 9; i <= summarySheet.getLastRowNum(); i++) {
				Row row = summarySheet.getRow(i);
				if (row.getLastCellNum() > 0) {

					if (workBook.getSheet(row.getCell(0).getStringCellValue().toString()) != null) {
						Hyperlink link = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
						link.setAddress(
								"'" + workBook.getSheet(row.getCell(0).getStringCellValue().toString()).getSheetName()
										+ "'!A1");
						link.setLabel(workBook.getSheet(row.getCell(0).getStringCellValue().toString()).getSheetName());
						row.getCell(1).setCellStyle(addressLinkStyle);
						row.getCell(1).setHyperlink(link);
					}
				}
			}

		}
	}

	private void formatSummaryHeaders(Sheet summarySheet, Workbook workBook) {
		formatSummaryHeader1(summarySheet, workBook);
		formatSummaryHeader2(summarySheet, workBook);
	}

	private void formatSummaryHeader1(Sheet summarySheet, Workbook workBook) {
		Row header1 = summarySheet.getRow(0);

		Font summaryHeaderFont = workBook.createFont();
		summaryHeaderFont.setBold(true);
		summaryHeaderFont.setColor(IndexedColors.WHITE.getIndex());

		CellStyle summaryHeaderStyle1 = workBook.createCellStyle();
		summaryHeaderStyle1.setFont(summaryHeaderFont);
		summaryHeaderStyle1.setAlignment(HorizontalAlignment.CENTER);
		summaryHeaderStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		summaryHeaderStyle1.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());

		header1.forEach(cell -> cell.setCellStyle(summaryHeaderStyle1));
	}

	private void alignCellContentsToLeftInSheet(Sheet sheet, Workbook workBook) {
		if (null != sheet) {
			CellStyle alignLeft = workBook.createCellStyle();
			alignLeft.setAlignment(HorizontalAlignment.LEFT);
			if (sheet.getPhysicalNumberOfRows() > 0) {
				for (Row row : sheet) {
					row.setRowStyle(alignLeft);
					for (Cell cell : row) {
						cell.setCellStyle(alignLeft);
						checkForDateCell(workBook, cell);
					}
				}
			}
		}

	}

	private void checkForDateCell(Workbook workBook, Cell cell) {
		try {
			if (cell.getDateCellValue() != null) {
				CellStyle dateStyle = workBook.createCellStyle();
				dateStyle.setAlignment(HorizontalAlignment.LEFT);
				CreationHelper createHelper = workBook.getCreationHelper();
				dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/mm/yyyy"));
				cell.setCellStyle(dateStyle);
			}
		} catch (Exception e) {

		}
	}

	public void formatOtherPages(Workbook workBook) {
		formatCells(workBook);
	}

	private void formatCells(Workbook workBook) {
		Font headerFont = workBook.createFont();
		headerFont.setBold(true);
		headerFont.setColor(IndexedColors.BLACK.getIndex());

		CellStyle headerStyle = workBook.createCellStyle();
		headerStyle.setFont(headerFont);
		headerStyle.setWrapText(true);
		headerStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);

		int numberOfSheets = workBook.getNumberOfSheets();
		for (int i = 1; i < numberOfSheets; i++) {
			Sheet currentSheet = workBook.getSheetAt(i);
			if (currentSheet.getPhysicalNumberOfRows() > 0) {
				Row row = currentSheet.getRow(currentSheet.getFirstRowNum());
				row.setRowStyle(headerStyle);
				for (Cell cell : row) {
					cell.setCellStyle(headerStyle);
				}
				formatDataCells(workBook, currentSheet);
				autoSizeColumns(currentSheet, row, 3);
			}
		}
	}

	private void formatDataCells(Workbook workBook, Sheet currentSheet) {
		CellStyle dataRowStyle = workBook.createCellStyle();
		dataRowStyle.setWrapText(true);
		dataRowStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);

		for (int j = 1; j <= currentSheet.getLastRowNum(); j++) {
			Row dataRow = currentSheet.getRow(j);
			dataRow.setRowStyle(dataRowStyle);
			formatCellValues(workBook, currentSheet, dataRow, 0);
		}
	}

	private void autoSizeColumns(Sheet sheet, Row row, int limit) {
		if (row.getPhysicalNumberOfCells() > 0) {
			limit = limit != 0 ? limit : row.getPhysicalNumberOfCells();
			for (int j = 0; j < limit; j++) {
				sheet.autoSizeColumn(row.getCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK).getColumnIndex());
			}
		}
	}

	private void formatCellValues(Workbook workBook, Sheet sheet, Row row, int limit) {
		DataFormatter objDefaultFormat = new DataFormatter();
		FormulaEvaluator objFormulaEvaluator = workBook.getCreationHelper().createFormulaEvaluator();

		CellStyle nCellStyle = workBook.createCellStyle();
		nCellStyle.setWrapText(true);
		nCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		nCellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());

		CellStyle dataCellStyle = workBook.createCellStyle();
		dataCellStyle.setWrapText(true);
		dataCellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);

		if (row.getPhysicalNumberOfCells() > 0) {
			limit = limit != 0 ? limit : row.getLastCellNum();
			for (int j = 0; j <= limit; j++) {
				objFormulaEvaluator.evaluate(row.getCell(j));
				String cellValueStr = objDefaultFormat.formatCellValue(row.getCell(j), objFormulaEvaluator);
				if ("N".equals(cellValueStr)) {
					row.getCell(j).setCellStyle(nCellStyle);
				} else if (row.getCell(j) != null) {
					row.getCell(j).setCellStyle(dataCellStyle);
				}
			}
		}
	}
}
