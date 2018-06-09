package com.wit.utils;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;


/**
 * 导出excel工具类
 * @author Wit
 *
 */
public class exportUtils {
	
	/**
	 * 生成表头Map的方法
	 * @param value  值
	 * @param rowStart 起始行标
	 * @param rowEnd   终止行标
	 * @param columnStart 起始列标
	 * @param columeEnd  终止列标
	 * @return Map<String,Object>
	 */
	public static Map<String, Object> setTitleCell(String value ,int rowStart , int rowEnd , int columnStart , int columeEnd ) {
		Map<String, Object> TitleCell = new HashMap<>();
		TitleCell.put("value", value);
		TitleCell.put("rowStart", rowStart);
		TitleCell.put("rowEnd", rowEnd);
		TitleCell.put("columnStart", columnStart);
		TitleCell.put("columeEnd", columeEnd);	
		return TitleCell;
	}
	
	
	/**
	 * 生成表头
	 * @param TitleCells 
	 *        表头的显示和位置信息，包括值(value)、起始行标(rowStart)、终止行标(rowEnd)、起始列标(columnStart)、终止列标(columeEnd)
	 * 		     具体格式为Map<String , Object> setTitleCell(String value ,int rowStart , int rowEnd , int columnStart , int columeEnd)
	 *        位置参数的起始值为0
	 * @param sheet
	 * @param wb
	 * @return int 返回下一行的行标，方便导出数据部分使用
	 */
	public static int SetTitle(List<Map<String, Object>> TitleCells,Sheet sheet,Workbook wb) {
		int rownum = -1 ;
		Map<String, Row> rows = new HashMap<>();
		Font font = wb.createFont();
	    font.setBold(true);	
		for (Map<String, Object> TitleCell : TitleCells) {
			int rowN = (int)TitleCell.get("rowStart");
			int columnN = (int)TitleCell.get("columnStart");
			String rowName = "row" + TitleCell.get("rowStart").toString();
			addMergeCell(TitleCell , sheet);
				
			if ((int)TitleCell.get("rowStart") > rownum) {
				Row row = sheet.createRow(rowN);
				rows.put(rowName, row);
				rownum = rowN;
			}
			
			
//			System.out.println(rowName + " " + (int)TitleCell.get("columnStart") + " " + TitleCell.get("value").toString());
			Cell cell = rows.get(rowName).createCell(columnN);
			cell.setCellValue(TitleCell.get("value").toString());	
			CellStyle cellStyle = wb.createCellStyle();
	        cellStyle.setAlignment(HorizontalAlignment.CENTER);
	        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
	        cellStyle.setFont(font);   
	        cell.setCellStyle(cellStyle);
	        sheet.setColumnWidth(columnN, (TitleCell.get("value").toString().length()+4)*500);
//	        System.out.println(TitleCell.get("value").toString().length()*500);

		}
		return rownum+1;
	}
	
	/**
	 * 针对List<Map<Object>>类型数据的导出
	 * @param DataCells 数据对象
	 * @param titles List对象，存储的是按title顺序存放的DataCells的键值
	 * @param sheet
	 * @param wb
	 * @param rownum SetTitle返回的值，若不需要表头则应为0
	 * @param maxNum 允许导出的最大数据条数，默认为5000
	 */
	public static void SetData(List<Map<String, Object>> DataCells,List<String> titles,Sheet sheet,Workbook wb,int rownum,int maxNum) {
		int count = 0;
		for (Map<String, Object> map : DataCells) {
			Row row = sheet.createRow(rownum);
			int column = 0;
			for (String str : titles) {
				Cell cell = row.createCell(column);
				cell.setCellValue(map.get(str).toString());
				column++;
			}
			count++;
			rownum++;
			if (count>=maxNum) {
				break;
			}
		}
		return;
	}
	public static void SetData(List<Map<String, Object>> DataCells,List<String> titles,Sheet sheet,Workbook wb,int rownum) {
		SetData(DataCells,titles,sheet,wb,rownum,5000);
	}
	
	/**
	 * 针对List<List<String>>类型数据的导出
	 * @param DataCells 数据对象
	 * @param sheet
	 * @param wb
	 * @param rownum SetTitle返回的值，若不需要表头则应为0
	 * @param maxNum 允许导出的最大数据条数，默认为5000
	 */
	public static void SetData(List<List<String>> DataCells,Sheet sheet,Workbook wb,int rownum,int maxNum) {
		int count = 0;
		for (List<String> list: DataCells) {
			Row row = sheet.createRow(rownum);
			int column = 0;
			for (String str : list) {
				Cell cell = row.createCell(column);
				cell.setCellValue(str);
				column++;
			}
			count++;
			rownum++;
			if (count>=maxNum) {
				break;
			}
		}
		return;
	}
	public static void SetData(List<List<String>> DataCells,Sheet sheet,Workbook wb,int rownum) {
		SetData(DataCells,sheet,wb,rownum,5000);
	}
	
	/**
	 * 生成单元格，若起止坐标是多个单元格则会生成范围内的合并单元格
	 * @param MergeCell 单元格的起止坐标，详见上方setTitleCell的说明
	 * @param sheet
	 * @return true表示生成单个单元格，false表示生成合并单元格
	 */
	public static boolean addMergeCell(Map<String, Object> MergeCell,Sheet sheet) {

		int rowStart = (int)MergeCell.get("rowStart");
		int rowEnd = (int)MergeCell.get("rowEnd"); 
		int columnStart = (int)MergeCell.get("columnStart"); 
		int columeEnd = (int)MergeCell.get("columeEnd");
		
		if ((rowStart==rowEnd) && (columnStart==columeEnd)) {
			return true;
		} else {
			sheet.addMergedRegion(new CellRangeAddress(
					rowStart, //first row (0-based)
					rowEnd, //last row  (0-based)
					columnStart, //first column (0-based)
					columeEnd  //last column  (0-based)
		    ));
			return false;
		}
	}
}
