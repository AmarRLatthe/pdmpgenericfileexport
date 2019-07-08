/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.logicoy.pdmpfileexport.util;

import com.fasterxml.jackson.databind.JsonNode;
import java.io.ByteArrayInputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 *
 * @author Amar
 */
public class GenericExcelBuilder {

    private static Logger logger = LoggerFactory.getLogger(GenericExcelBuilder.class);

    /**
     * 
     * @param jsonInfo
     * @param from
     * @param to
     * @param filter
     * @return 
     */
    public String writeExcel(JsonNode jsonInfo, String from, String to, String filter) {
        HSSFWorkbook hssfWorkbook = null;
        HSSFRow row = null;
        HSSFSheet hssfSheet = null;
        try {
            hssfWorkbook = new HSSFWorkbook();
            hssfSheet = hssfWorkbook.createSheet("Transaction List");
            int counter = 0;

            row = hssfSheet.createRow((short) counter);

            //Create a new font and alter it.
            HSSFFont font = hssfWorkbook.createFont();
            font.setFontHeightInPoints((short) 15);
            font.setFontName("IMPACT");
            font.setItalic(false);
            font.setColor(HSSFColor.GREEN.index);

            //Set font into style
            HSSFCellStyle style = hssfWorkbook.createCellStyle();
            style.setFont(font);
            row.setRowStyle(style);

            // Writing file creation date
            HSSFCell c1 = row.createCell((short) 0);
            c1.setCellStyle(style);
            style.setWrapText(true);
            c1.setCellValue("Execution Date: " + new SimpleDateFormat("MM/dd/yyyy HH:mm:ss").format(new Date())
                    + "\n Total Records: " + jsonInfo.size() + "\n Given Filter Criteria " + "\n Transaction start date: " + from
                    + "\n Transaction end date: " + to + filter);

            logger.info("Total response object : " + jsonInfo.size());

            counter++;
            Iterator<JsonNode> it = jsonInfo.elements();

            if (it.hasNext()) {
                JsonNode next = it.next();
                Iterator<String> names = next.fieldNames();
                int cellCounter = 0;
                HSSFCellStyle headerStyle = hssfWorkbook.createCellStyle();
                HSSFFont headerFont = hssfWorkbook.createFont();

                headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
                headerStyle.setFont(headerFont);
                Row headerRow = null;
                headerRow = hssfSheet.createRow((short) counter);
                headerStyle.setFont(headerFont);
                headerRow.setRowStyle(headerStyle);

                while (names.hasNext()) {
                    String n = names.next();
                    Cell cell = headerRow.createCell((short) cellCounter);
                    cell.setCellValue(n.replace('_', ' ').toUpperCase());
                    cell.setCellStyle(headerStyle);
                    cellCounter++;
                }
            }
            counter++;
            int cellCounter = 0;
            while (it.hasNext()) {
                JsonNode next = it.next();
                Iterator<String> names = next.fieldNames();
                cellCounter = 0;
                row = hssfSheet.createRow((short) counter);
                while (names.hasNext()) {
                    String n = names.next();
                    row.createCell((short) cellCounter).setCellValue(
                            next.path(n).asText());
                    cellCounter++;
                }
                counter++;
                if (counter > 10000) {
                    break;
                }
            }
            for (int columnIndex = 0; columnIndex < cellCounter; columnIndex++) {
                hssfSheet.autoSizeColumn(columnIndex);
            }
            counter += 5;
            java.io.ByteArrayOutputStream bos = new java.io.ByteArrayOutputStream();
            hssfWorkbook.write(bos);
            bos.close();
            byte[] bytes = bos.toByteArray();
            ByteArrayInputStream fis = new ByteArrayInputStream(bytes);
            fis.read(bytes);
            String base64 = new sun.misc.BASE64Encoder().encode(bytes);
            System.out.println("JSON data successfully exported to excel!");
            return base64;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }
}
