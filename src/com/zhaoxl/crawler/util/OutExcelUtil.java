/*
 * Project: seaway-p2p-biz-mgr
 * 
 * File Created at 2014-2-20 下午3:44:04
 * 
 * Copyright 2012 seaway.com Corporation Limited.
 * All rights reserved.
 *
 * This software is the confidential and proprietary information of
 * Seaway Company. ("Confidential Information").  You shall not
 * disclose such Confidential Information and shall use it only in
 * accordance with the terms of the license agreement you entered into
 * with seaway.com.
 */
package com.zhaoxl.crawler.util;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.*;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map.Entry;
import java.util.Set;

public class OutExcelUtil {

    /**
     * 
     * @desz
     * @param row    从第几行开始导入数据库的数据
     * @param dataName   导入的名字 
     * @param data   数据
     * @return
     * @throws IOException
     * @throws NoSuchMethodException
     * @throws SecurityException
     * @throws IllegalAccessException
     * @throws IllegalArgumentException
     * @throws InvocationTargetException
     * @throws InvalidFormatException 
     */
    public static <T> XSSFWorkbook getTemplate(String fileName, int row, List<String> dataName, List<T> data)
            throws IOException, NoSuchMethodException, SecurityException, IllegalAccessException,
            IllegalArgumentException, InvocationTargetException, InvalidFormatException {
        // 模版路径
        // String filePath = UserMessReportCtr.class.getResource("/").getPath();
    	OPCPackage pkg = OPCPackage.open(fileName);
        XSSFWorkbook book = new XSSFWorkbook(pkg);
        XSSFSheet sheet = (XSSFSheet) book.getSheet("Sheet1");
        // 样式对象 居中
        XSSFCellStyle style = book.createCellStyle();
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 垂直
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 水平
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN); // 边框加粗
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);// 边框加粗
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);// 边框加粗
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);// 边框加粗
        XSSFCell cell = null;

        if (data != null && !data.isEmpty()) {
            // 第二行开始 循环数据库里的数据
            for (int i = 0; i < data.size(); i++) {
                XSSFRow row2 = sheet.createRow(i + row);
                T mdl = data.get(i);
                Class<? extends Object> clazz = mdl.getClass();
                cell = row2.createCell(0);// 创建序号列
                cell.setCellValue(i + 1); // 设置序号
                cell.setCellStyle(style); // 样式
                int j = 1;
                for (String key : dataName) {
                    String upcase = key.substring(0, 1).toUpperCase();
                    String getMethodName = "get" + upcase + key.substring(1);
                    Method getMethod = clazz.getMethod(getMethodName);
                    Object obj = getMethod.invoke(mdl);
                    cell = row2.createCell(j);
                    cell.setCellStyle(style); // 样式
                    if (null == obj) {
                        cell.setCellValue("");
                    } else if (obj instanceof Integer) {
                        cell.setCellValue(((Integer) obj).intValue());
                    } else if (obj instanceof String) {
                        cell.setCellValue((String) obj);
                    } else if (obj instanceof Date) {
                        String dateString = new SimpleDateFormat("yyyy-MM-dd").format((Date) obj);
                        cell.setCellValue(dateString);
                    } else if (obj instanceof Boolean) {
                        cell.setCellValue(obj.toString());
                    } else if (obj instanceof Double) {
                        cell.setCellValue(((Double) obj).doubleValue());
                    } else {
                        cell.setCellValue(obj.toString());
                    }
                    j++;
                }
            }
        }
        return book;
    }

    /**
     * 
     * @desc 导出表格 
     * @param  number  true导出序号，  false不导出序号
     * @param linkedHashMap
     * @param list2
     * @return
     * @throws Exception
     */
    public static <T> XSSFWorkbook getExcel(boolean number, LinkedHashMap<String, String> linkedHashMap, List<T> list2,
            String[] str) throws Exception {
        // 创建Excel的工作书册 Workbook,对应到一个excel文档
        XSSFWorkbook wb = new XSSFWorkbook();
        // 创建Excel的工作sheet,对应到一个excel文档的tab
        XSSFSheet sheet = wb.createSheet("sheet1");
        // 样式
        XSSFCellStyle style = wb.createCellStyle();
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 垂直
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 水平
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN); // 边框加粗
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);// 边框加粗
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);// 边框加粗
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);// 边框加粗
        style.setWrapText(true);//强制换行，当内容超过宽度时换行
        // 创建第一行
        XSSFRow row = sheet.createRow(0);
        // 创建n个Excel的单元格 (n列)
        XSSFCell cell = null;
        int k = 0;
        // 如果需要导出序号，那么第一列为序号列 标题为序号
        if (number) {
            sheet.setColumnWidth(k, 4500);// 设置单元格宽度
            cell = row.createCell(0);// 创建 序号列
            cell.setCellValue("序号"); // 设置序号列内容
            cell.setCellStyle(style); // 样式

            k++;
        }
        // 设置其他列的标题
        Set<Entry<String, String>> entrySet = linkedHashMap.entrySet();
        for (Entry<String, String> entry : entrySet) {
            sheet.setColumnWidth(k, 4500);// 设置单元格宽度
            String value = entry.getValue();
            cell = row.createCell(k);// 创建第 i+1列
            cell.setCellValue(value); // 设置列里的标题
            cell.setCellStyle(style); // 样式

            k++;
        }
        if (list2 != null && !list2.isEmpty()) {
            // 第二行开始 循环数据库里的数据
            for (int i = 0; i < list2.size(); i++) {
                XSSFRow row2 = sheet.createRow(i + 1);
                T mdl = list2.get(i);
                Class<? extends Object> clazz = mdl.getClass();
                Set<String> keySet = linkedHashMap.keySet();
                int j = 0;
                // 设置序号
                if (number) {
                    cell = row2.createCell(0);// 创建序号列
                    cell.setCellValue(i + 1); // 设置序号
                    cell.setCellStyle(style); // 样式
                    j++;
                }
                // 设置数据
                for (String key : keySet) {
                    String upcase = key.substring(0, 1).toUpperCase();
                    String getMethodName = "get" + upcase + key.substring(1);
                    Method getMethod = clazz.getMethod(getMethodName);
                    Object obj = getMethod.invoke(mdl);
                    cell = row2.createCell(j);
                    cell.setCellStyle(style); // 样式
                    // cell.setCellType(Cell.CELL_TYPE_STRING);
                    if (null == obj) {
                        cell.setCellValue("");
                    } else if (obj instanceof Integer) {
                        // cell.setCellValue((String) obj);
                        cell.setCellValue(((Integer) obj).intValue());
                    } else if (obj instanceof String) {
                        cell.setCellValue((String) obj);
                    } else if (obj instanceof Date) {
                        String dateString = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format((Date) obj);
                        cell.setCellValue(dateString);
                    } else if (obj instanceof Boolean) {
                        cell.setCellValue(obj.toString());
                    } else if (obj instanceof Double) {
                        cell.setCellValue(((Double) obj).doubleValue());
                    } else {
                        cell.setCellValue(obj.toString());
                    }
                    j++;
                }
            }

            if (str != null) {
                XSSFRow rowX = sheet.createRow(list2.size() + 1);
                XSSFCell cellX = null;
                sheet.setColumnWidth(0, 4500);// 设置单元格宽度
                cellX = rowX.createCell(0);
                cellX.setCellValue("总计");
                cellX.setCellStyle(style); // 样式
                for (int i = 0; i < str.length; i++) {
                    cellX = rowX.createCell(i + 1);
                    cellX.setCellValue(str[i]);
                    cellX.setCellStyle(style); // 样式
                }
            }

        }
        return wb;
    }
    
   /* *//**
     * 
     * 带序号excel
     * @param linkedHashMap
     * @param list2
     * @return
     * @throws Exception
     * <br>----------------------------------------------------变更记录--------------------------------------------------
     * <br> 序号      |           时间                        	|   作者      |                          描述                                                         
     * <br> 0     | 2014年4月14日 下午7:41:41  	|  姚鑫炜    | 创建
     *//*
    public static <T> XSSFWorkbook getExcelXh(LinkedHashMap<String, String> linkedHashMap, List<T> list2)
            throws Exception {
        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
        // 创建Excel的工作书册 Workbook,对应到一个excel文档
        XSSFWorkbook wb = new XSSFWorkbook();
        // 创建Excel的工作sheet,对应到一个excel文档的tab
        XSSFSheet sheet = wb.createSheet("sheet1");
        // 第一行
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = null;
        // 创建n个Excel的单元格 (n列)
        Set<Entry<String, String>> entrySet = linkedHashMap.entrySet();
        sheet.setColumnWidth(0, 1500);// 设置单元格宽度
        cell = row.createCell(0);// 创建第 序号列
        cell.setCellValue("序号"); // 设置序号列
        int k = 1;
        for (Entry<String, String> entry : entrySet) {
            sheet.setColumnWidth(k, 4500);// 设置单元格宽度
            String value = entry.getValue();
            cell = row.createCell(k);// 创建第 i+1列
            cell.setCellValue(value); // 设置列里的标题
            k++;
        }
        if (list2 != null && !list2.isEmpty()) {
            // 第二行开始 循环数据库里的数据
            for (int i = 0; i < list2.size(); i++) {
                XSSFRow row2 = sheet.createRow(i + 1);
                T mdl = list2.get(i);
                Class<? extends Object> clazz = mdl.getClass();
                Set<String> keySet = linkedHashMap.keySet();

                cell = row2.createCell(0);// 创建序号列
                cell.setCellValue(i + 1); // 设置序号
                int j = 1;
                for (String key : keySet) {
                	cell = row2.createCell(j);
                	if(key.equals("quantity")){
                		Method getMethod = clazz.getMethod("getProductDetail");
                		Object obj = getMethod.invoke(mdl);
                		CreateOrderReq order = JSON.parseObject(obj.toString(), CreateOrderReq.class);
                		Long num = order.getProducts().get(0).getNum();
                		cell.setCellValue(num);
                	} else if(key.equals("contacts")){
                		Method getMethod = clazz.getMethod("getProductDetail");
                		Object obj = getMethod.invoke(mdl);
                		CreateOrderReq order = JSON.parseObject(obj.toString(), CreateOrderReq.class);
                		String contacts = order.getContacts();
                		cell.setCellValue(contacts);
                	} else if(key.equals("telephone")){
                		Method getMethod = clazz.getMethod("getProductDetail");
                		Object obj = getMethod.invoke(mdl);
                		CreateOrderReq order = JSON.parseObject(obj.toString(), CreateOrderReq.class);
                		String tepephone = order.getTelephone();
                		cell.setCellValue(tepephone);
                	} else if(key.equals("deliveryAddress")){
                		Method getMethod = clazz.getMethod("getProductDetail");
                		Object obj = getMethod.invoke(mdl);
                		CreateOrderReq order = JSON.parseObject(obj.toString(), CreateOrderReq.class);
                		String deliveryAddress = order.getDeliveryAddress();
                		cell.setCellValue(deliveryAddress);
                	} else {
                		String upcase = key.substring(0, 1).toUpperCase();
                		String getMethodName = "get" + upcase + key.substring(1);
                		Method getMethod = clazz.getMethod(getMethodName);
                		Object obj = getMethod.invoke(mdl);
                		if (null == obj) {
                			cell.setCellValue("");
                		} else if (obj instanceof Integer) {
                			cell.setCellValue(((Integer) obj).intValue());
                		} else if (obj instanceof String) {
                			cell.setCellValue((String) obj);
                		} else if (obj instanceof Date) {
                			String dateString = formatter.format((Date) obj);
                			cell.setCellValue(dateString);
                		} else if (obj instanceof Boolean) {
                			cell.setCellValue(obj.toString());
                		} else if (obj instanceof Double) {
                			cell.setCellValue(((Double) obj).doubleValue());
                		} else {
                			cell.setCellValue(obj.toString());
                		}
                	}
                    j++;
                }
            }
        }
        return wb;
    }
*/
    /**
     * 
     * @desc  导出表格  不带序号  特殊的
     * @param linkedHashMap
     * @param list
     * @return
     * @throws Exception
     */
    public static <T> XSSFWorkbook getExcel3(LinkedHashMap<String, String> linkedHashMap, List<T> list, String s1,
            String s2) throws Exception {

        // 创建Excel的工作书册 Workbook,对应到一个excel文档
        XSSFWorkbook wb = new XSSFWorkbook();
        // 创建Excel的工作sheet,对应到一个excel文档的tab
        XSSFSheet sheet = wb.createSheet("sheet1");
        // 样式对象 居中
        XSSFCellStyle style = wb.createCellStyle();
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 垂直
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 水平
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN); // 边框加粗
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);// 边框加粗
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);// 边框加粗
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);// 边框加粗
        // 第一行
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = null;
        // 创建n个Excel的单元格 (n列)
        Set<Entry<String, String>> entrySet = linkedHashMap.entrySet();
        int k = 0;
        for (Entry<String, String> entry : entrySet) {
            sheet.setColumnWidth(k, 4500);// 设置单元格宽度
            cell = row.createCell(k);// 创建第 i+1列
            cell.setCellValue(entry.getValue()); // 设置列里的标题
            cell.setCellStyle(style); // 样式
            k++;
        }
        // 判断表格内容是否为空
        if (list != null && !list.isEmpty()) {
            // 第二行开始 循环数据库里的数据
            for (int i = 0; i < list.size(); i++) {
                XSSFRow row2 = sheet.createRow(i + 1);
                T mdl = list.get(i);
                Class<? extends Object> clazz = mdl.getClass();
                Set<String> keySet = linkedHashMap.keySet();
                int j = 0;
                for (String key : keySet) {
                    String upcase = key.substring(0, 1).toUpperCase();
                    String getMethodName = "get" + upcase + key.substring(1);
                    Method getMethod = clazz.getMethod(getMethodName);
                    Object obj = getMethod.invoke(mdl);
                    cell = row2.createCell(j);
                    cell.setCellStyle(style); // 样式
                    if (null == obj) {
                        cell.setCellValue("");
                        // 判断去到的属性类型 整形
                    } else if (obj instanceof Integer) {
                        cell.setCellValue(((Integer) obj).intValue());
                        // 字符串
                    } else if (obj instanceof String) {
                        cell.setCellValue((String) obj);
                        // 日期
                    } else if (obj instanceof Date) {
                        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                        String dateString = formatter.format((Date) obj);
                        cell.setCellValue(dateString);
                        // boolean
                    } else if (obj instanceof Boolean) {
                        cell.setCellValue(obj.toString());
                        // 浮点型
                    } else if (obj instanceof Double) {
                        cell.setCellValue(((Double) obj).doubleValue());
                        // 其他转成字符型
                    } else {
                        cell.setCellValue(obj.toString());
                    }
                    j++;
                }
            }

            XSSFRow rowX = sheet.createRow(list.size() + 1);
            XSSFCell cellX = null;
            sheet.setColumnWidth(0, 4500);// 设置单元格宽度
            cellX = rowX.createCell(0);
            cellX.setCellValue("总计");
            cellX = rowX.createCell(4);
            cellX.setCellValue(s1);
            cellX = rowX.createCell(5);
            cellX.setCellValue(s2);
        }
        return wb;
    }

    
    /**
     * 
     * @desc 
     * @param filePath  模版路径
     * @param row    从第几行开始导入数据库的数据
     * @param dataName   导入的名字 
     * @param data   数据
     * @return
     * @throws IOException
     * @throws NoSuchMethodException
     * @throws SecurityException
     * @throws IllegalAccessException
     * @throws IllegalArgumentException
     * @throws InvocationTargetException
     * @throws InvalidFormatException 
     */
    public static <T> XSSFWorkbook getCountTemplate(String fileName, int row, List<String> dataName, List<T> data,String[] str, int year)
            throws IOException, NoSuchMethodException, SecurityException, IllegalAccessException,
            IllegalArgumentException, InvocationTargetException, InvalidFormatException {
        // 模版路径
        // String filePath = UserMessReportCtr.class.getResource("/").getPath();
    	OPCPackage pkg = OPCPackage.open(fileName);
        XSSFWorkbook book = new XSSFWorkbook(pkg);
        XSSFSheet sheet = (XSSFSheet) book.getSheet("Sheet1");
        // 样式对象 居中
        XSSFCellStyle style = book.createCellStyle();
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 垂直
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 水平
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN); // 边框加粗
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);// 边框加粗
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);// 边框加粗
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);// 边框加粗
        XSSFCell cell = null;
        
        if (data != null && !data.isEmpty()) {
            // 第二行开始 循环数据库里的数据
            for (int i = 0; i < data.size(); i++) {
                XSSFRow row2 = sheet.createRow(i + row);
                T mdl = data.get(i);
                Class<? extends Object> clazz = mdl.getClass();
                int j = 0;
                for (String key : dataName) {
                    String upcase = key.substring(0, 1).toUpperCase();
                    String getMethodName = "get" + upcase + key.substring(1);
                    Method getMethod = clazz.getMethod(getMethodName);
                    Object obj = getMethod.invoke(mdl);
                    cell = row2.createCell(j);
                    cell.setCellStyle(style); // 样式
                    if (null == obj) {
                        cell.setCellValue("");
                    } else if (obj instanceof Integer) {
                        cell.setCellValue(((Integer) obj).intValue());
                    } else if (obj instanceof String) {
                        cell.setCellValue((String) obj);
                    } else if (obj instanceof Date) {
                        String dateString = new SimpleDateFormat("yyyy-MM-dd").format((Date) obj);
                        cell.setCellValue(dateString);
                    } else if (obj instanceof Boolean) {
                        cell.setCellValue(obj.toString());
                    } else if (obj instanceof Double) {
                        cell.setCellValue(((Double) obj).doubleValue());
                    } else {
                        cell.setCellValue(obj.toString());
                    }
                    j++;
                }
            }
            
            if(year != 0){
            	XSSFRow yearRow = sheet.getRow(2);
            	XSSFCell yearCell = yearRow.getCell(2);
            	yearCell.setCellValue(String.valueOf(year));
            }
            
            if (str != null) {
                XSSFRow rowX = sheet.createRow(data.size() + row);
                XSSFCell cellX = null;
                sheet.setColumnWidth(0, 4500);// 设置单元格宽度
                cellX = rowX.createCell(0);
                cellX.setCellValue("总计");
                cellX.setCellStyle(style); // 样式
                for (int i = 0; i < str.length; i++) {
                    cellX = rowX.createCell(i + 1);
                    cellX.setCellValue(str[i]);
                    cellX.setCellStyle(style); // 样式
                }
            }
        }
        return book;
    }
    
    
    // *********************************************
    /**
     * 
     * @desc java 里写表头
     * @param list
     * @param list2
     * @return
     * @throws Exception
     */
    /*
     * public static <T> XSSFWorkbook getExcel2(LinkedList<String> list, List<T>
     * list2) throws Exception { // 创建Excel的工作书册 // Workbook,对应到一个excel文档
     * XSSFWorkbook wb = new XSSFWorkbook();
     * 
     * // 创建Excel的工作sheet,对应到一个excel文档的tab XSSFSheet sheet =
     * wb.createSheet("sheet1"); // 样式对象 居中 边框加粗 XSSFCellStyle style =
     * wb.createCellStyle();
     * style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 垂直
     * style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 水平 //
     * style.setBorderBottom(HSSFCellStyle.BORDER_THIN); // 边框加粗 //
     * style.setBorderLeft(HSSFCellStyle.BORDER_THIN);// 边框加粗 //
     * style.setBorderRight(HSSFCellStyle.BORDER_THIN);// 边框加粗 //
     * style.setBorderTop(HSSFCellStyle.BORDER_THIN);// 边框加粗 // 1、创建第一行第二行第三行
     * XSSFRow row1 = sheet.createRow(0); XSSFRow row2 = sheet.createRow(1);
     * XSSFRow row3 = sheet.createRow(2); // 2、设置单元格宽度 第一列
     * sheet.setColumnWidth(0, 2000); // 合并单元格 ，四个参数分别是：起始行，结束行，起始列，结束列
     * 
     * sheet.addMergedRegion(new CellRangeAddress(0, 2, 0, 0));
     * sheet.setColumnWidth(1, 2000); sheet.addMergedRegion(new
     * CellRangeAddress(0, 2, 1, 1)); sheet.setColumnWidth(2, 2000);
     * sheet.addMergedRegion(new CellRangeAddress(0, 1, 2, 4));
     * sheet.setColumnWidth(5, 5000); sheet.addMergedRegion(new
     * CellRangeAddress(0, 2, 5, 5)); sheet.setColumnWidth(6, 2000);
     * sheet.addMergedRegion(new CellRangeAddress(0, 1, 6, 8));
     * sheet.setColumnWidth(9, 4000); sheet.addMergedRegion(new
     * CellRangeAddress(0, 2, 9, 9)); sheet.setColumnWidth(10, 4000);
     * sheet.addMergedRegion(new CellRangeAddress(0, 2, 10, 10));
     * sheet.setColumnWidth(11, 2000); sheet.addMergedRegion(new
     * CellRangeAddress(0, 0, 11, 18)); //
     * ----------------------------------------第一组 XSSFCell a13 =
     * row1.createCell(0); a13.setCellValue("序号"); // 表格的第一行第一列显示的数据
     * a13.setCellStyle(style); //
     * ------------------------------------------------------ 第二组 XSSFCell b13 =
     * row1.createCell(1); b13.setCellValue("月份"); b13.setCellStyle(style); //
     * ------------------------------------------------------ 第三组 XSSFCell
     * c12e12 = row1.createCell(2); c12e12.setCellValue("本期新增注册用户");
     * c12e12.setCellStyle(style); XSSFCell c3 = row3.createCell(2);
     * c3.setCellValue("企业"); c3.setCellStyle(style); XSSFCell d3 =
     * row3.createCell(3); d3.setCellValue("个人"); d3.setCellStyle(style);
     * XSSFCell e3 = row3.createCell(4); e3.setCellValue("合计");
     * e3.setCellStyle(style); //
     * ------------------------------------------------------第四组 XSSFCell f13 =
     * row1.createCell(5); f13.setCellValue("本期新增临时用户数 ");
     * f13.setCellStyle(style); //
     * ------------------------------------------------------ 第五组 XSSFCell
     * g12i12 = row1.createCell(6); g12i12.setCellValue("总注册用户数 ");
     * g12i12.setCellStyle(style); XSSFCell g3 = row3.createCell(6);
     * g3.setCellValue("企业"); g3.setCellStyle(style); XSSFCell h3 =
     * row3.createCell(7); h3.setCellValue("个人"); h3.setCellStyle(style);
     * XSSFCell i3 = row3.createCell(8); i3.setCellValue("合计");
     * i3.setCellStyle(style); //
     * ---------------------------------------------------- 第六组 XSSFCell g13 =
     * row1.createCell(9); g13.setCellValue("总临时用户数 "); g13.setCellStyle(style);
     * // ------------------------------------------------------ 第七组 XSSFCell
     * h13 = row1.createCell(10); h13.setCellValue("总用户数");
     * h13.setCellStyle(style); //
     * ------------------------------------------------------ 第八组 XSSFCell i1s1
     * = row1.createCell(11); i1s1.setCellValue("用户活跃度 ");
     * i1s1.setCellStyle(style); sheet.addMergedRegion(new CellRangeAddress(1,
     * 1, 11, 14)); XSSFCell i2o2 = row2.createCell(11);
     * i2o2.setCellValue("本期活跃用户数"); i2o2.setCellStyle(style); XSSFCell l3 =
     * row3.createCell(11); l3.setCellValue("企业 "); l3.setCellStyle(style);
     * XSSFCell m3 = row3.createCell(12); m3.setCellValue("个人 ");
     * m3.setCellStyle(style); XSSFCell n3 = row3.createCell(13);
     * n3.setCellValue("合计 "); n3.setCellStyle(style); XSSFCell o3 =
     * row3.createCell(14); o3.setCellValue("占本月总注册用户数百分比 ");
     * o3.setCellStyle(style); sheet.addMergedRegion(new CellRangeAddress(1, 1,
     * 15, 18)); XSSFCell p2s2 = row2.createCell(15);
     * p2s2.setCellValue("本期沉默用户数 "); p2s2.setCellStyle(style); XSSFCell p3 =
     * row3.createCell(15); p3.setCellValue("企业 "); p3.setCellStyle(style);
     * XSSFCell q3 = row3.createCell(16); q3.setCellValue("个人 ");
     * q3.setCellStyle(style); XSSFCell r3 = row3.createCell(17);
     * r3.setCellValue("合计 "); r3.setCellStyle(style); XSSFCell s3 =
     * row3.createCell(18); s3.setCellValue("占本月总注册用户数百分比 ");
     * s3.setCellStyle(style);
     * 
     * return wb; }
     */

    
}
