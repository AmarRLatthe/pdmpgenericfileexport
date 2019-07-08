/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.logicoy.pdmpfileexport.util;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.dataformat.csv.CsvMapper;
import com.fasterxml.jackson.dataformat.csv.CsvSchema;
import java.io.IOException;
import java.util.logging.Level;

/**
 *
 * @author Amar
 */
public class GenericCSVBuilder {

    /**
     * 
     * @param jsonInfo
     * @return 
     */
    public String writeCSV(JsonNode jsonInfo) {
        String base64 = null;

        try {
            CsvSchema.Builder csvSchemaBuilder = CsvSchema.builder();
            JsonNode firstObject = jsonInfo.elements().next();
            firstObject.fieldNames().forEachRemaining(fieldName -> {
                csvSchemaBuilder.addColumn(fieldName);
            });
            CsvSchema csvSchema = csvSchemaBuilder.build().withHeader();
            CsvMapper csvMapper = new CsvMapper();
            byte[] arr = csvMapper.writerFor(JsonNode.class)
                    .with(csvSchema)
                    .writeValueAsBytes(jsonInfo);
            base64 = new sun.misc.BASE64Encoder().encode(arr);
            return base64;
        } catch (IOException ex) {
            java.util.logging.Logger.getLogger(GenericExcelBuilder.class.getName()).log(Level.SEVERE, null, ex);
        }
        return null;
    }
}
