package com.sf.materielmanage.service.util;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
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

import com.sf.materielmanage.constant.MaterielManageConstant;
import com.sf.materielmanage.query.CartonUsageQuery;
import com.sf.materielmanage.util.ColumnInfoEntity;

/**
 * Title: CartonUsageExcelUtil.java  
 * Description:  
 * Copyright: Copyright (c) 2018
 * @author Kang Chen  
 * @date 2018年7月26日 下午8:34:00
 * @version 1.0  
 */
public class CartonUsageExcelUtil {
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
    public static void exportExcelFile(String title, List<ColumnInfoEntity> headList, List<CartonUsageQuery> queryList, HttpServletResponse response) throws IOException{
    	ByteArrayOutputStream output = new ByteArrayOutputStream();	
    	exportExcelOfCartonUsage(headList,queryList,null,0,output);
    	byte[] content = output.toByteArray();
        InputStream input = new ByteArrayInputStream(content);
    	response.reset();
    	response.setContentType("application/octet-stream");
    	title = URLEncoder.encode(title, "UTF-8");
        response.setHeader("Content-Disposition", "attachment;filename="+title+".xlsx");
        response.setHeader("Content-Type", "application/octet-stream");
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
     * 纸箱领用excel导出
     * @param headList
     * @param jsonArray
     * @param datePattern
     * @param colWidth
     * @param out
     */
    private static void exportExcelOfCartonUsage(List<ColumnInfoEntity> headList,List<CartonUsageQuery> queryList,String datePattern,int colWidth, OutputStream out) {
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
        cellFont.setBoldweight(HSSFFont.COLOR_NORMAL);
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
        // 写导出表格标题
        int rowIndex = 0;
        if(rowIndex == 65535 || rowIndex == 0){
    		//如果数据超过了，则在第二页显示
    		if ( rowIndex != 0 ) {
                sheet = (SXSSFSheet) workbook.createSheet();
            }
    		//列头 rowIndex =0
            SXSSFRow headerRow = (SXSSFRow) sheet.createRow(0);
            for(int i=0;i<headers.length;i++)
            {
                headerRow.createCell(i).setCellValue(headers[i]);
                headerRow.getCell(i).setCellStyle(headerStyle);

            }
          //数据内容从 rowIndex=1开始
            rowIndex = 1;
    	}
        // 遍历集合数据，产生数据行
        for(CartonUsageQuery cartonUsage : queryList){
        	if(rowIndex == 65535){
        		//如果数据超过了，则在第二页显示
        		if ( rowIndex != 0 ) {
                    sheet = (SXSSFSheet) workbook.createSheet();
                }
        		//列头 rowIndex =0
                SXSSFRow headerRow = (SXSSFRow) sheet.createRow(0);
                for(int i=0;i<headers.length;i++)
                {
                    headerRow.createCell(i).setCellValue(headers[i]);
                    headerRow.getCell(i).setCellStyle(headerStyle);

                }
              //数据内容从 rowIndex=1开始
                rowIndex = 1;
        	}
        	Field[] fieldArray =cartonUsage.getClass().getDeclaredFields();
        	for(int k = 0;k < 6;k++){
                SXSSFRow dataRow = (SXSSFRow) sheet.createRow(rowIndex);
                for(int i = 0;i < properties.length;i++){
                	for(int c = 0;c < fieldArray.length;c++){
                        //设置对象的访问权限，保证对private的属性的访问
                		fieldArray[c].setAccessible(true);
                		if("type".equals(properties[i])){
                			SXSSFCell newCell = (SXSSFCell) dataRow.createCell(i);	
                			switch(k){
	                			case 0:newCell.setCellValue("一号纸箱");break;
	                			case 1:newCell.setCellValue("二号纸箱");break;
	                			case 2:newCell.setCellValue("三号纸箱");break;
	                			case 3:newCell.setCellValue("四号纸箱");break;
	                			case 4:newCell.setCellValue("五号纸箱");break;
	                			case 5:newCell.setCellValue("六号纸箱");break;
                			}
                			newCell.setCellStyle(cellStyle);
                		}else if("cartonAppliedQty".equals(properties[i])){
                			getDefinedCell(i,cartonUsage,dataRow,k,cellStyle,"cartonAppliedQty");	
                		}else if("cartonMonthUsageQty".equals(properties[i])){
                			getDefinedCell(i,cartonUsage,dataRow,k,cellStyle,"cartonMonthUsageQty");	
                		}else if("cartonTimeUsageQty".equals(properties[i])){
                			getDefinedCell(i,cartonUsage,dataRow,k,cellStyle,"cartonTimeUsageQty");	
                		}else if(properties[i].equals(fieldArray[c].getName())){
                			SXSSFCell newCell = (SXSSFCell) dataRow.createCell(i);	
                			if(k == 0){
                				// 合并单元格
                                sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex+5, i, i));	
                			}
                            Object o = null;
    						try {
    							o = fieldArray[c].get(cartonUsage);
    						} catch (IllegalArgumentException | IllegalAccessException e) {
    							e.printStackTrace();
    						}
                            newCell.setCellValue(getStrData(o));
                            newCell.setCellStyle(cellStyle);
                            break;
                		}
                	}
                }
            	rowIndex++;
            }	
        }
        try {
            workbook.write(out);
            workbook.dispose();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    private static void getDefinedCell(int i,CartonUsageQuery query,SXSSFRow dataRow,int k,CellStyle cellStyle,String qtyType){
    	SXSSFCell newCell = (SXSSFCell) dataRow.createCell(i);	
		CartonUsageQuery mapQuery = query.getUsageQtyMap().get(qtyType);
		String carton1 = "";
		String carton2 = "";
		String carton3 = "";
		String carton4 = "";
		String carton5 = "";
		String carton6 = "";
		if(mapQuery != null){
			carton1 = getStrData(mapQuery.getCarton1());
			carton2 = getStrData(mapQuery.getCarton2());
			carton3 = getStrData(mapQuery.getCarton3());
			carton4 = getStrData(mapQuery.getCarton4());
			carton5 = getStrData(mapQuery.getCarton5());
			carton6 = getStrData(mapQuery.getCarton6());
		}
		switch(k){
			case 0:newCell.setCellValue(carton1);break;
			case 1:newCell.setCellValue(carton2);break;
			case 2:newCell.setCellValue(carton3);break;
			case 3:newCell.setCellValue(carton4);break;
			case 4:newCell.setCellValue(carton5);break;
			case 5:newCell.setCellValue(carton6);break;
		}
		newCell.setCellStyle(cellStyle);
    }
    
    /**
     * 格式化数据
     * @param obj
     * @return
     */
    private static String getStrData(Object obj){
    	String cellValue = "";
	    if(obj==null) {
	        cellValue = "0";
	    }
	    else if(obj instanceof Date) {
	        cellValue = new SimpleDateFormat("yyyy/MM/dd").format(obj);
	    }
	    else if(obj instanceof Float || obj instanceof Double || obj instanceof BigDecimal) {
	        cellValue= new BigDecimal(obj.toString()).setScale(2,BigDecimal.ROUND_HALF_UP).toString();
	    }
	    else {
	        cellValue = obj.toString();
	    }
    	return cellValue;
    }

}
