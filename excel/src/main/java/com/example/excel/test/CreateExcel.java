package com.example.excel.test;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.filechooser.FileSystemView;
import java.awt.*;
import java.io.*;
import java.util.Date;

public class CreateExcel {
	public static void main(String[] args) throws IOException {
		copyWookbook();

		//存储路径--获取桌面位置
		FileSystemView view = FileSystemView.getFileSystemView();
		File directory = view.getHomeDirectory();
		System.out.println(directory);

		//存储Excel的路径
		String path = directory+"\\test11.xlsx";
		System.out.println(path);
		try {

			//定义一个Excel表格
			SXSSFWorkbook wb = new SXSSFWorkbook();  //创建工作薄
			Sheet sheet = wb.createSheet("sheet1"); //创建工作表

			//添加表头数据
			for (int j = 0; j < 5; j++) {
				Row row = sheet.createRow(j); //行
				for (int i = 0; i < 5; i++) {//列
					//将取到的值依次写到Excel的第一行的cell中
					Cell cell1=row.createCell(i);
					cell1.setCellType(1);
					cell1.setCellValue("行"+j+"-"+i+"列");
				}
			}


			//创建输出流
			FileOutputStream outputStream = new FileOutputStream(path);
			wb.write(outputStream);
			//记得关闭流
			outputStream.flush();
			outputStream.close();
			System.out.println("写入成功");
		} catch (Exception e) {
			System.out.println("写入失败");
			e.printStackTrace();
		}finally {

		}
	}
//复制文件 这样如果有需求需要对Excel数据进行校对 那么就很好处理 处理的时候其实是文件副本
	public static void  copyWookbook() throws IOException {
		File file = new File("C:\\Users\\Administrator\\Desktop\\test11.xlsx");
		Workbook book = new XSSFWorkbook(new FileInputStream(file));


		Sheet shee=book.getSheetAt(0);
		shee.getCellComment(1,1).getString();

		String path="C:\\Users\\Administrator\\Desktop\\test12.xlsx";

		FileOutputStream outputStream = new FileOutputStream(path);
		book.write(outputStream);
		outputStream.flush();
		outputStream.close();
	}

}
