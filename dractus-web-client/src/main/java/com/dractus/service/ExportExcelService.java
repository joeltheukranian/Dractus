package com.dractus.service;

import com.dractus.bean.util.ExcelListColumn;
import com.dractus.controller.dashboard.business.BusinessExpensesCtrl;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.xssf.usermodel.*; //possible new library we can use
import org.springframework.stereotype.Service;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * User: d.lemeshevsky
 */
@Service
public class ExportExcelService {

    private static final String FILE_TYPE = ".xls";

    public void generateExport(HttpServletResponse response, List dataList,
                               Map<String, ExcelListColumn> excelMap, String headerString,
                               Double amountSum, Double chargedSum, Double withoutTaxSum) throws Exception {
//        BufferedInputStream buf = null;
        ServletOutputStream myOut = null;
        try {
            SimpleDateFormat format = customDataFormat("MM.dd.yyyy_HH-mm-ss");
            Date dateImport = new Date();

            String filename = format.format(dateImport) + FILE_TYPE;

            // Create Excel Workbook and Sheet
            HSSFWorkbook wb = new HSSFWorkbook();


            // Setup the output
            String contentType = "application/vnd.ms-excel";
            response.setHeader("Content-disposition", "attachment; filename="
                    + filename);
            
            response.setHeader("Cache-Control", "no-store, must-revalidate"); // HTTP 1.1.
            response.setHeader("Pragma", "private"); // HTTP 1.0.
            response.setDateHeader("Expires", 1); // Proxi
            
            response.setContentType(contentType);
            
            
            myOut = response.getOutputStream();

//            for(int counterSheet = 1; counterSheet <= 2; ++counterSheet){
                int counterSheet = 1;
                createSheets(wb, counterSheet, dataList, excelMap, headerString,
                        amountSum, chargedSum, withoutTaxSum);
//            }


            // Write out the spreadsheet
            wb.write(myOut);

        } catch (Exception e) {
            e.printStackTrace();
            throw new Exception();
        } finally {
            //close the input/output streams
            if (myOut != null)
                myOut.close();
//            if (buf != null)
//                buf.close();

        }
    }

    private void createSheets(HSSFWorkbook wb, int counterSheet, List dataList,
                              Map<String, ExcelListColumn> excelMap, String headerString,
                              Double amountSum, Double chargedSum, Double withoutTaxSum) throws Exception {

        HSSFSheet sheet = wb.createSheet(String.valueOf(counterSheet));
        
        //Set to Landscape/legal
        sheet.getPrintSetup().setLandscape(true);
        sheet.getPrintSetup().setPaperSize(HSSFPrintSetup.LEGAL_PAPERSIZE);

        //autofit, fitToPage (setAutobreaks is also necessary)
        //See https://issues.apache.org/bugzilla/show_bug.cgi?id=20497
        
//        sheet.getPrintSetup().setFitHeight((short)0);
//        sheet.getPrintSetup().setFitWidth((short)1);
//        sheet.setFitToPage(false); //commented out did NOT work
//        sheet.setAutobreaks(true);

        
//        sheet.getPrintSetup().setFitHeight((short)0);/
//        sheet.getPrintSetup().setFitWidth((short)1);
////        sheet.setFitToPage(false); //commented out did NOT work
////        sheet.setAutobreaks(true);

        
//        sheet.getPrintSetup().setFitHeight((short)0);
//        sheet.getPrintSetup().setFitWidth((short)1);
//        sheet.setFitToPage(false); //commented out did NOT work
//        sheet.setAutobreaks(true);

//another..nope
//        sheet.setFitToPage(true);  // optionally; does not seem to have any effect 
//        sheet.getPrintSetup().setFitWidth((short)1); 
//        sheet.getPrintSetup().setFitHeight((short)0); 

//        sheet.setFitToPage(true);  // optionally; does not seem to have any effect 
//        sheet.getPrintSetup().setFitWidth((short)1); 
//        sheet.getPrintSetup().setFitHeight((short)0); 
//        sheet.setAutobreaks(true);

        sheet.setFitToPage(true);  // optionally; does not seem to have any effect 
        sheet.getPrintSetup().setFitWidth((short)1); 
        sheet.getPrintSetup().setFitHeight((short)99); 
        sheet.setAutobreaks(true);
        
        List oldObjects = new ArrayList();

        Row row;
        Cell cell;
        int counterRow = 0;

        if (headerString != null) {
            // Create head in table
            row = sheet.createRow(counterRow);
            counterRow++;
            row.setHeightInPoints(17.25f);
            cell = createCell(row, 0,
                    getHeaderStyle(wb.createCellStyle(), true));
            cell.setCellValue(headerString);
        }

        // Create head in table
        row = sheet.createRow(counterRow);
        counterRow++;
        row.setHeightInPoints(17.25f);
        for (ExcelListColumn currentColumn : excelMap.values()) {
            currentColumn.addSheetStyle(sheet);
            cell = createCell(row, currentColumn.getRotation(),
                    currentColumn.addCellStyle(wb.createCellStyle(), true));
            cell.setCellValue(currentColumn.getTitle());
        }

        for (Object o : dataList) {

            if (counterRow == 65000) {
                break;
            }

            row = sheet.createRow(counterRow);
            row.setHeightInPoints(17.25f);
            
            //insert page page every 20 rows
//            if((counterRow % 20) == 0){
//            	sheet.setRowBreak(counterRow);
//	            }

            Class currentClass = o.getClass();
            Field fields[] = currentClass.getDeclaredFields();
            
            //make red font
            HSSFCellStyle styleRedFont = wb.createCellStyle();
            HSSFFont font = wb.createFont();
            font.setColor(HSSFColor.RED.index);
            styleRedFont.setFont(font);
            //Use Format ($#,##0.00)
            styleRedFont.setDataFormat((short)8); 

            //end make red font
            
            for (int i = 0; i < fields.length; i++) {
                if (excelMap.containsKey(fields[i].getName())) {
                    ExcelListColumn currentColumn = excelMap.get(fields[i]
                            .getName());

                    cell = createCell(row, currentColumn.getRotation(),
                            currentColumn.addCellStyle(wb.createCellStyle(),
                                    false));
                    fields[i].setAccessible(true);
                    //make font red if negative, while setting value
                    currentColumn.addCellStyle(styleRedFont, false);
                    addCellValueByTypes(wb, i, cell, fields[i].getType(),
                            fields[i].get(o), styleRedFont);
                }
            }
            counterRow++;
            oldObjects.add(o);
        }

        dataList.removeAll(oldObjects);
        if (counterRow == 65000 && dataList.size() > 0) {
            counterSheet++;
            createSheets(wb, counterSheet, dataList, excelMap,
                    headerString, amountSum, chargedSum, withoutTaxSum);
        } else {
            //create total line
            row = sheet.createRow(counterRow);
            row.setHeightInPoints(17.25f);




            ExcelListColumn currentColumn = excelMap.get(BusinessExpensesCtrl.ExpenseViewListColumns.amount
                    .toString());

//            cell = createCell(row, currentColumn.getRotation()-1,
//                    getHeaderStyle(wb.createCellStyle(), true));
//            cell.setCellValue("Totals");

            cell = createCell(row, currentColumn.getRotation(),
                    getHeaderStyle(wb.createCellStyle(), true));
            cell.setCellValue(amountSum);

            //Make Red if negative
            HSSFCellStyle styleRedFontForSum = wb.createCellStyle();
            HSSFFont fontForSum = wb.createFont();
            fontForSum.setColor(HSSFColor.RED.index);
            styleRedFontForSum.setFont(fontForSum);
            //Use Format ($#,##0.00)
            styleRedFontForSum.setDataFormat((short)8);
            getHeaderStyle(styleRedFontForSum, true);
            
            if(amountSum < 0) {
            	cell.setCellStyle(styleRedFontForSum);
            }

            currentColumn = excelMap.get(BusinessExpensesCtrl.ExpenseViewListColumns.hstGstCharged
                    .toString());

            cell = createCell(row, currentColumn.getRotation(),
                    getHeaderStyle(wb.createCellStyle(), true));
            cell.setCellValue(chargedSum);
            //make red if negative
            if(chargedSum < 0) {
            	cell.setCellStyle(styleRedFontForSum);
            }

            currentColumn = excelMap.get(BusinessExpensesCtrl.ExpenseViewListColumns.withoutTax
                    .toString());

            cell = createCell(row, currentColumn.getRotation(),
                    getHeaderStyle(wb.createCellStyle(), true));
            cell.setCellValue(withoutTaxSum);
            //make red if negative
            if(withoutTaxSum < 0) {
            	cell.setCellStyle(styleRedFontForSum);
            }
        }

        //autosize all columns. Done at end so that size can be computed
        for (int i = 0; i < excelMap.values().size(); i++) {
            sheet.autoSizeColumn(i);
        }

    }

    private CellStyle getHeaderStyle(CellStyle currentCellStyle, Boolean isHead) {
        currentCellStyle.setBorderBottom(CellStyle.BORDER_MEDIUM);
        currentCellStyle.setBorderLeft(CellStyle.BORDER_MEDIUM);
        currentCellStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        currentCellStyle.setBorderTop(CellStyle.BORDER_MEDIUM);
        currentCellStyle.setAlignment(CellStyle.ALIGN_LEFT);
        currentCellStyle.setWrapText(false);
        return currentCellStyle;
    }

    private Cell createCell(Row row, int column, CellStyle style)
            throws Exception {
        Cell cell = row.createCell(column);
        cell.setCellStyle(style);
        return cell;
    }

    private void addCellValueByTypes(HSSFWorkbook wb, int index, Cell cell, Class type, Object value, HSSFCellStyle styleRedFont)
            throws Exception {
    	//DEBUG 
    	//System.out.println("***Index: " + index + " Type: " + type);
        if (value == null) {
            cell.setCellValue("");
        } else if (type.equals(Long.class)) {
        	Long cellValue = (Long) value;
            cell.setCellValue(cellValue);
            //make red if negative
            if(cellValue < 0) {
            	cell.setCellStyle(styleRedFont);
            }
        } else if (type.equals(long.class)) {
        	long cellValue = (Long) value;
            cell.setCellValue(cellValue);
            //make red if negative
            if(cellValue < 0) {
            	cell.setCellStyle(styleRedFont);
            }
        } else if (type.equals(Date.class)) {
        	//This should be optimized so we don't create a dateStyle for EACH cell. For now it'll be GC'd after export
        	HSSFCellStyle dateStyle = wb.createCellStyle();
        	dateStyle.setDataFormat((short)0xe);

        	cell.setCellStyle(dateStyle);
        	
            //SimpleDateFormat format = customDataFormat("MM/dd/yyyy");
        	Date date = (Date) value;
            cell.setCellValue(date);
            cell.getCellStyle().setAlignment(CellStyle.ALIGN_LEFT);
        } else if (type.equals(String.class)) {
        	cell.setCellValue((String) value);
        } else if (type.equals(BigDecimal.class)) {
        	BigDecimal cellValue = (BigDecimal) value;
            cell.setCellValue(((BigDecimal) value).doubleValue());
            //make red if negative
            if(cellValue.doubleValue() < 0) {
            	cell.setCellStyle(styleRedFont);
            }
        } else if (type.equals(Boolean.class)) {
            cell.setCellValue((Boolean) value);
        } else if (type.equals(Double.class)) {
        	Double cellValue = (Double) value;
            cell.setCellValue(cellValue);
            //make red if negative
            if(cellValue.doubleValue() < 0) {
            	cell.setCellStyle(styleRedFont);
            } else {
            	//Use Format ($#,##0.00)
            	cell.getCellStyle().setDataFormat((short)8);
            }
        }
    }


    private SimpleDateFormat customDataFormat(String pattern) {
        return new SimpleDateFormat(pattern);
    }
}
