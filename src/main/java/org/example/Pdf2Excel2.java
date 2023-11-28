//package org.example;
//
//import java.io.FileInputStream;
//import java.io.FileOutputStream;
//import java.io.IOException;
//
//import org.apache.pdfbox.pdmodel.PDDocument;
//import org.apache.pdfbox.text.PDFTextStripper;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//public class Pdf2Excel2 {
//    public static void main(String[] args) {
//        String pdfFilePath = "C:\\Users\\lx\\Desktop\\source.pdf";
//        String excelFilePath = "C:\\Users\\lx\\Desktop\\target.xlsx";
//
//        try {
//            String pdfText = getPdfText(pdfFilePath);
//            writeTextToExcel(pdfText, excelFilePath);
//            System.out.println("PDF successfully converted to Excel!");
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }
//
//    private static String getPdfText(String pdfFilePath) throws IOException {
//        PDDocument pdfDoc = PDDocument.load(new FileInputStream(pdfFilePath));
//        PDFTextStripper stripper = new PDFTextStripper();
//        String pdfText = stripper.getText(pdfDoc);
//        pdfDoc.close();
//        return pdfText;
//    }
//
//    private static void writeTextToExcel(String pdfText, String excelFilePath) throws IOException {
//        boolean isXlsx = excelFilePath.endsWith(".xlsx");
//        Workbook workbook = isXlsx ? new XSSFWorkbook() : new HSSFWorkbook();
//        Sheet sheet = workbook.createSheet("Sheet1");
//        String[] lines = pdfText.split("\n");
//        int rowIndex = 0;
//        for (String line : lines) {
//            Row row = sheet.createRow(rowIndex++);
//            String[] cells = line.split("	");
//            int cellIndex = 0;
//            for (String cellValue : cells) {
//                Cell cell = row.createCell(cellIndex++);
//                cell.setCellValue(cellValue);
//            }
//        }
//        workbook.write(new FileOutputStream(excelFilePath));
//        workbook.close();
//    }
//
//}
