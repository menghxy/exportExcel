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
 * ����excel������
 * @author Wit
 *
 */
public class exportUtils {
	
	/**
	 * ���ɱ�ͷMap�ķ���
	 * @param value  ֵ
	 * @param rowStart ��ʼ�б�
	 * @param rowEnd   ��ֹ�б�
	 * @param columnStart ��ʼ�б�
	 * @param columeEnd  ��ֹ�б�
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
	 * ���ɱ�ͷ
	 * @param TitleCells 
	 *        ��ͷ����ʾ��λ����Ϣ������ֵ(value)����ʼ�б�(rowStart)����ֹ�б�(rowEnd)����ʼ�б�(columnStart)����ֹ�б�(columeEnd)
	 * 		     �����ʽΪMap<String , Object> setTitleCell(String value ,int rowStart , int rowEnd , int columnStart , int columeEnd)
	 *        λ�ò�������ʼֵΪ0
	 * @param sheet
	 * @param wb
	 * @return int ������һ�е��б꣬���㵼�����ݲ���ʹ��
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
	 * ���List<Map<Object>>�������ݵĵ���
	 * @param DataCells ���ݶ���
	 * @param titles List���󣬴洢���ǰ�title˳���ŵ�DataCells�ļ�ֵ
	 * @param sheet
	 * @param wb
	 * @param rownum SetTitle���ص�ֵ��������Ҫ��ͷ��ӦΪ0
	 * @param maxNum ���������������������Ĭ��Ϊ5000
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
	 * ���List<List<String>>�������ݵĵ���
	 * @param DataCells ���ݶ���
	 * @param sheet
	 * @param wb
	 * @param rownum SetTitle���ص�ֵ��������Ҫ��ͷ��ӦΪ0
	 * @param maxNum ���������������������Ĭ��Ϊ5000
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
	 * ���ɵ�Ԫ������ֹ�����Ƕ����Ԫ��������ɷ�Χ�ڵĺϲ���Ԫ��
	 * @param MergeCell ��Ԫ�����ֹ���꣬����Ϸ�setTitleCell��˵��
	 * @param sheet
	 * @return true��ʾ���ɵ�����Ԫ��false��ʾ���ɺϲ���Ԫ��
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
