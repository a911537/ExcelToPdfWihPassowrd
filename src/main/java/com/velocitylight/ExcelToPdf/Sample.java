package com.velocitylight.ExcelToPdf;

public class Sample {
    public static void main(String[] args) {
        String excelPath = "C:\\Users\\user\\Desktop\\vm_ip_record.xlsx";
        String pdfPath = "C:\\\\Users\\\\user\\\\Desktop\\\\sample.pdf";

        Converter instance = new Converter(excelPath, pdfPath);
        try {
            instance.doConvert();
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
}
