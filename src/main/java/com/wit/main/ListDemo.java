package com.wit.main;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.wit.utils.*;


/**
 * List<List<String>>类型数据的导出Demo
 * @author Wit
 *
 */
public class ListDemo {
	public static void main(String[] args) {
		
		//生成excel对象
		Workbook wb = new HSSFWorkbook();
	    Sheet sheet = wb.createSheet("new sheet");
		
		//设置表头信息
		List<Map<String, Object>> TitleCells = new ArrayList<>();
		Map<String, Object> TitleCell = exportUtils.setTitleCell("时间",0,1,0,0);
		Map<String, Object> TitleCell2 = exportUtils.setTitleCell("B2C",0,0,1,2);
		Map<String, Object> TitleCell3 = exportUtils.setTitleCell("店铺数(万家)",1,1,1,1);
		Map<String, Object> TitleCell4 = exportUtils.setTitleCell("店铺数占比(%)",1,1,2,2);
		Map<String, Object> TitleCell5 = exportUtils.setTitleCell("C2C",0,0,3,4);
		Map<String, Object> TitleCell6 = exportUtils.setTitleCell("店铺数(万家)",1,1,3,3);
		Map<String, Object> TitleCell7 = exportUtils.setTitleCell("店铺数占比(%)",1,1,4,4);
		TitleCells.add(TitleCell);
		TitleCells.add(TitleCell2);
		TitleCells.add(TitleCell3);
		TitleCells.add(TitleCell4);
		TitleCells.add(TitleCell5);
		TitleCells.add(TitleCell6);
		TitleCells.add(TitleCell7);
		
	    //生成表头
	    int row_n = exportUtils.SetTitle(TitleCells,sheet,wb);
	    
	    //设置数据信息
	    List<List<String>> datas = new ArrayList<>();
	    List<String> data1 = new ArrayList<>();
	    String a1 = "2017-01-01";
	    String a2 = "100";
	    String a3 = "20";
	    String a4 = "200";
	    String a5 = "30";
	    data1.add(a1);data1.add(a2);data1.add(a3);data1.add(a4);data1.add(a5);
	    datas.add(data1);
	    List<String> data2 = new ArrayList<>();
	    String b1 = "2017-02-01";
	    String b2 = "400";
	    String b3 = "10";
	    String b4 = "300";
	    String b5 = "36";
	    data2.add(b1);data2.add(b2);data2.add(b3);data2.add(b4);data2.add(b5);
	    datas.add(data2);
	    List<String> data3 = new ArrayList<>();
	    String c1 = "2017-03-01";
	    String c2 = "103";
	    String c3 = "27";
	    String c4 = "500";
	    String c5 = "34";
	    data3.add(c1);data3.add(c2);data3.add(c3);data3.add(c4);data3.add(c5);
	    datas.add(data3);
	    
	    //生成数据
	    exportUtils.SetData(datas,sheet,wb,row_n);

	    // Write the output to a file
	    try (OutputStream fileOut = new FileOutputStream("D:\\workbook.xls")) {
	        wb.write(fileOut);
	        wb.close();
	    } catch (Exception e) {
			// TODO: handle exception
	    	e.printStackTrace();
		}
	}
}
