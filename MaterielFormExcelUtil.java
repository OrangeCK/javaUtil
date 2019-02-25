package com.sf.materielmanage.service.util;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.sf.materielmanage.util.ColumnInfoEntity;

/**
 * Title: MaterielFormExcelUtil.java  
 * Description:  
 * Copyright: Copyright (c) 2018
 * @author Kang Chen  
 * @date 2018年7月31日 下午2:20:52
 * @version 1.0  
 */
public class MaterielFormExcelUtil {
    public static String DEFAULT_DATE_PATTERN="yyyy年MM月dd日";//默认日期格式
    public static int DEFAULT_COLOUMN_WIDTH = 25;
    public static SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");  

    /**
     * Web 导出excel .xlsx
     * @param title
     * @param headList
     * @param dataArray
     * @param response
     * @throws IOException 
     */
    public static void exportExcelFile(String title, List<ColumnInfoEntity> headList, JSONArray dataArray, HttpServletResponse response) throws IOException{
    	ByteArrayOutputStream output = new ByteArrayOutputStream();	
    	exportExcelXNoTitle(headList,dataArray,null,0,output);
    	byte[] content = output.toByteArray();
        InputStream input = new ByteArrayInputStream(content);
    	response.reset();
    	response.setContentType("application/octet-stream");
    	title = URLEncoder.encode(title, "UTF-8");
        response.setHeader("Content-Disposition", "attachment;filename="+title+".xlsx");  
        response.setContentLength(content.length);
        ServletOutputStream outputStream = response.getOutputStream();
        BufferedInputStream bis = new BufferedInputStream(input);
        BufferedOutputStream bos = new BufferedOutputStream(outputStream);
        byte[] buff = new byte[8192];
        int bytesRead;
        while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
            bos.write(buff, 0, bytesRead);

        }
        bis.close();
        bos.close();
        outputStream.flush();
        outputStream.close();
    }
    
    /**
     * 正常表格导出
     * @param headList
     * @param jsonArray
     * @param datePattern
     * @param colWidth
     * @param out
     */
    private static void exportExcelXNoTitle(List<ColumnInfoEntity> headList,JSONArray jsonArray,String datePattern,int colWidth, OutputStream out) {
        if(datePattern==null) {
            datePattern = DEFAULT_DATE_PATTERN;
        }
        // 声明一个工作薄 缓存大于1000行时会把之前的行写入硬盘
        SXSSFWorkbook workbook = new SXSSFWorkbook(1000);
        workbook.setCompressTempFiles(true);
        //表头样式
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        Font titleFont = workbook.createFont();
        titleFont.setFontHeightInPoints((short) 20);
        titleFont.setBoldweight((short) 700);
        titleStyle.setFont(titleFont);
        titleStyle.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
        //titleStyle.setFillBackgroundColor(IndexedColors.WHITE.getIndex());
        titleStyle.setFillBackgroundColor(HSSFColor.RED.index);
        // 列头样式
        CellStyle headerStyle = workbook.createCellStyle();
        // 设置边框
        headerStyle.setBorderTop((short) 1);
        headerStyle.setBorderRight((short) 1);
        headerStyle.setBorderBottom((short) 1);
        headerStyle.setBorderLeft((short) 1);
        // 水平居中
        headerStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 生成一个字体
        Font headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setColor(IndexedColors.BLACK.getIndex());
        headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        headerStyle.setFont(headerFont);
        // 单元格样式
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 垂直居中
        cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        // 生成一个字体
        Font cellFont = workbook.createFont();
        cellFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        cellStyle.setFont(cellFont);
        cellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        DataFormat format = workbook.createDataFormat();
        cellStyle.setDataFormat(format.getFormat("@"));
        // 生成一个(带标题)表格
        SXSSFSheet sheet = (SXSSFSheet) workbook.createSheet();
        //设置列宽 //至少字节数
        int minBytes = colWidth<DEFAULT_COLOUMN_WIDTH?DEFAULT_COLOUMN_WIDTH:colWidth;
        int[] arrColWidth = new int[headList.size()];
        // 产生表格标题行,以及设置列宽
        String[] properties = new String[headList.size()];
        String[] headers = new String[headList.size()];
        int ii = 0;

        for(ColumnInfoEntity columnInfo:headList){
            properties[ii] = columnInfo.getColumn();
            headers[ii] = columnInfo.getColumnName();
            int bytes = columnInfo.getColumn().getBytes().length;
            arrColWidth[ii] =  bytes < minBytes ? minBytes : bytes;
            sheet.setColumnWidth(ii,arrColWidth[ii]*256);
            ii++;
        }
        // 遍历集合数据，产生数据行
        int rowIndex = 0;
        for (Object obj : jsonArray) {
            if(rowIndex == 65535 || rowIndex == 0){
                //如果数据超过了，则在第二页显示
                if ( rowIndex != 0 ) {
                    sheet = (SXSSFSheet) workbook.createSheet();
                }
                for(int r = 0;r < 2;r++){
                    SXSSFRow headerRow = (SXSSFRow) sheet.createRow(r);
                    for(int i = 0;i < headers.length;i++){
                    	SXSSFCell headCell = (SXSSFCell) headerRow.createCell(i);	
                    	if("网点代码".equals(headers[i]) || "网点名称".equals(headers[i]) || "月票均".equals(headers[i])){
                    		if(r == 0){
                            	sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex+1, i, i));		
                    		}
                    		headCell.setCellValue(headers[i]);
                        	headCell.setCellStyle(headerStyle);
                    	}else if("运单类".equals(headers[i])){
                    		if(r == 0){
                            	sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, i, i+6));
                            	headCell.setCellValue("综合分析");
                            	headCell.setCellStyle(headerStyle);
                    		}else{
                    			headCell.setCellValue(headers[i]);
                            	headCell.setCellStyle(headerStyle);	
                    		}
                    	}else if("纸质运单".equals(headers[i])){
                    		if(r == 0){
                            	sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, i, i+3));
                            	headCell.setCellValue("运单类");
                            	headCell.setCellStyle(headerStyle);
                    		}else{
                    			headCell.setCellValue(headers[i]);
                            	headCell.setCellStyle(headerStyle);	
                    		}	
                    	}else if("文件封".equals(headers[i])){
                    		if(r == 0){
                            	sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, i, i+4));
                            	headCell.setCellValue("外包装类");
                            	headCell.setCellStyle(headerStyle);
                    		}else{
                    			headCell.setCellValue(headers[i]);
                            	headCell.setCellStyle(headerStyle);	
                    		}	
                    	}else if("泡沫式内包装".equals(headers[i])){
                    		if(r == 0){
                            	sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, i, i+1));
                            	headCell.setCellValue("内包装类");
                            	headCell.setCellStyle(headerStyle);
                    		}else{
                    			headCell.setCellValue(headers[i]);
                            	headCell.setCellStyle(headerStyle);	
                    		}	
                    	}else if("泡沫填充物".equals(headers[i])){
                    		if(r == 0){
                            	sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, i, i+3));
                            	headCell.setCellValue("填充材料类");
                            	headCell.setCellStyle(headerStyle);
                    		}else{
                    			headCell.setCellValue(headers[i]);
                            	headCell.setCellStyle(headerStyle);	
                    		}	
                    	}else if("贴纸类".equals(headers[i])){
                    		if(r == 0){
                            	sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, i, i+5));
                            	headCell.setCellValue("辅助材料类");
                            	headCell.setCellStyle(headerStyle);
                    		}else{
                    			headCell.setCellValue(headers[i]);
                            	headCell.setCellStyle(headerStyle);	
                    		}	
                    	}else if("一次性编织袋".equals(headers[i])){
                    		if(r == 0){
                            	sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, i, i+1));
                            	headCell.setCellValue("编织材料类");
                            	headCell.setCellStyle(headerStyle);
                    		}else{
                    			headCell.setCellValue(headers[i]);
                            	headCell.setCellStyle(headerStyle);	
                    		}	
                    	}else{
                    		headCell.setCellValue(headers[i]);
                        	headCell.setCellStyle(headerStyle);		
                    	}
                    	
                    }
                }
                //数据内容从 rowIndex=2开始
                rowIndex = 2;
            }
            JSONObject jo = (JSONObject) JSONObject.toJSON(obj);
            SXSSFRow dataRow = (SXSSFRow) sheet.createRow(rowIndex);
            for (int i = 0; i < properties.length; i++)
            {
                SXSSFCell newCell = (SXSSFCell) dataRow.createCell(i);
                Object o =  jo.get(properties[i]);
                String cellValue = "";
                if(o==null) {
                    cellValue = "";
                }
                else if(o instanceof Date) {
                    cellValue = sdf.format(o);
                }
                else if(o instanceof Float || o instanceof Double || o instanceof BigDecimal) {
                    cellValue= new BigDecimal(o.toString()).setScale(2,BigDecimal.ROUND_HALF_UP).toString();
                }
                else {
                    cellValue = o.toString();
                }

                newCell.setCellValue(cellValue);
                newCell.setCellStyle(cellStyle);
            }
            rowIndex++;
        }
        try {
            workbook.write(out);
            workbook.dispose();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    
}
