package com.velocitylight.ExcelToPdf;

import java.io.File;
import java.io.FileOutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.itextpdf.text.Document;
import com.itextpdf.text.Font;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

public class Converter {

    private String excelPath;
    private String pdfPath;
    private String password;

    public Converter(String excelPath, String pdfPath) {
        this.excelPath = excelPath;
        this.pdfPath = pdfPath;
    }

    public void doConvert() throws Exception {
        int maxNum = 1;
        // Read workbook into HSSFWorkbook
        Workbook workbook = WorkbookFactory.create(new File(excelPath));
        // Read worksheet into HSSFSheet
        Sheet worksheet = workbook.getSheetAt(0);
        for (int j = 0; j <= worksheet.getPhysicalNumberOfRows(); j++) {
            Row row = worksheet.getRow(j);
            if (null != row) {
                if (maxNum < row.getPhysicalNumberOfCells()) {
                    maxNum = row.getPhysicalNumberOfCells();
                }
            }
        }
        // We will create output PDF document objects at this point
        Document doc = new Document();
        doc.setPageSize(PageSize.A4);
        PdfWriter writer = PdfWriter.getInstance(doc, new FileOutputStream(pdfPath));
        
        BaseFont bf = BaseFont.createFont("MHei-Medium", "UniCNS-UCS2-H", BaseFont.NOT_EMBEDDED);
        Font font = new Font(bf, 10, Font.NORMAL);
        
        PdfPTable table = new PdfPTable(--maxNum);
        // Loop through rows.
        for (int i = 0; i <= worksheet.getPhysicalNumberOfRows(); i++) {
            PdfPCell cell = new PdfPCell();
            cell.setHorizontalAlignment(PdfPCell.ALIGN_CENTER);
            Row row = worksheet.getRow(i);
            if (null != row) {
                int currentNum = row.getPhysicalNumberOfCells()-1;
                for (int j = 0; j < maxNum; j++) {
                    if (currentNum <= maxNum) {
                        String value = getCellFormatValue(row.getCell(j));
                        cell.setPhrase(new Paragraph(value, font));
                        // 調整取密碼的欄位
                        if (i==20 && j==0 ) {
                            password = value;
                        }
                    } else {
                        cell.setPhrase(new Paragraph("", font));
                    }
                    table.addCell(cell);
                }
            }
        }
        // Encrypt pdf
        writer.setEncryption(password.getBytes(), password.getBytes(), PdfWriter.ALLOW_PRINTING, PdfWriter.DO_NOT_ENCRYPT_METADATA);
        doc.open();
        // Finally add the table to PDF document
        doc.add(table);
        doc.close();
    }
    
    private static String getCellFormatValue(Cell cell) {
        String cellValue = "";
        if (null != cell) {
            switch(cell.getCellTypeEnum()) {
            case NUMERIC:
            case FORMULA:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    cellValue = sdf.format(date);
                } else {
                    DecimalFormat df = new DecimalFormat("########");
                    cellValue = df.format(cell.getNumericCellValue());
                }
                break;
            case STRING:
                cellValue = cell.getRichStringCellValue().getString();
                break;
            default:
                cellValue = " ";
            }
        } else {
            cellValue = "";
        }
        return cellValue;
    }
}
