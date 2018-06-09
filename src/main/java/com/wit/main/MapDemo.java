package com.wit.main;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.wit.utils.exportUtils;


/**
 * List<Map<String,Object>>类型数据的导出Demo
 * @author Wit
 *
 */
public class MapDemo {
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
	    List<Map<String, Object>> datas = new ArrayList<>();
	    Map<String, Object> data1 = new HashMap<>();
	    data1.put("time", "2017-01-01");
	    data1.put("b2c_num", "100");
	    data1.put("b2c_percent", "20");
	    data1.put("c2c_num", "200");
	    data1.put("c2c_percent", "30");
	    datas.add(data1);
	    Map<String, Object> data2 = new HashMap<>();
	    data2.put("time", "2017-02-01");
	    data2.put("b2c_num", "103");
	    data2.put("b2c_percent", "27");
	    data2.put("c2c_num", "270");
	    data2.put("c2c_percent", "40");
	    datas.add(data2);
	    Map<String, Object> data3 = new HashMap<>();
	    data3.put("time", "2017-03-01");
	    data3.put("b2c_num", "160");
	    data3.put("b2c_percent", "27");
	    data3.put("c2c_num", "400");
	    data3.put("c2c_percent", "40");
	    datas.add(data3);
	    
	    List<String> titles = new ArrayList<>();
	    titles.add("time");
	    titles.add("b2c_num");
	    titles.add("b2c_percent");
	    titles.add("c2c_num");
	    titles.add("c2c_percent");
	    
	    //生成数据
	    exportUtils.SetData(datas, titles, sheet, wb, row_n);

	    // Write the output to a file
	    try (OutputStream fileOut = new FileOutputStream("D:\\workbook2.xls")) {
	        wb.write(fileOut);
	        wb.close();
	    } catch (Exception e) {
			// TODO: handle exception
	    	e.printStackTrace();
		}
	    
	}
}
