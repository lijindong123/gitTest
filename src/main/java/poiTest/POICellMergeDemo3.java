package poiTest;

import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

public class POICellMergeDemo3 {

	public static void main(String[] args) {

		HSSFWorkbook wb = new HSSFWorkbook();// 创建一个Excel文件
		HSSFSheet sheet = wb.createSheet("表1");// 创建一个Excel的Sheet
		
		//设置一个样式，下面可以调用
		CellStyle style = wb.createCellStyle();
		//设置下边框的线条粗细（有14种选择，可以根据需要在BorderStyle这个类中选取）
		style.setBorderBottom(BorderStyle.THIN);
		//设置下边框的边框线颜色（颜色和上述的颜色对照表是一样的）
		style.setBorderBottom(BorderStyle.THIN);//下边框
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(BorderStyle.THIN);//左边框
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderTop(BorderStyle.THIN); //上边框
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderRight(BorderStyle.THIN);//右边框
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setAlignment(HorizontalAlignment.CENTER);//水平居中
		style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
		
		// 设置列宽
		sheet.setColumnWidth(0, 2800);
		sheet.setColumnWidth(1, 2800);
		sheet.setColumnWidth(2, 2800);
		sheet.setColumnWidth(3, 2800);
		sheet.setColumnWidth(4, 4500);
		sheet.setColumnWidth(5, 2800);
		sheet.setColumnWidth(6, 2800);
		sheet.setColumnWidth(7, 2800);
		sheet.setColumnWidth(8, 2800);
		try {
			HSSFRow row = null;
			HSSFCell cell = null;
			// ---------------------------1.初始化带边框的表头------------------------------
			for (int i = 0; i < 6; i++) {
				row = sheet.createRow(i);
				for (int j = 0; j <= 8; j++) {
					cell = row.createCell(j);
					cell.setCellStyle(style);
				}
			}
			// ---------------------------2.指定单元格填充数据------------------------------
			
			cell = sheet.getRow(0).getCell(0);
			cell.setCellValue(new HSSFRichTextString("转出人员查询展示"));
			
			cell = sheet.getRow(1).getCell(0);
			cell.setCellValue(new HSSFRichTextString("基本信息"));
			cell = sheet.getRow(1).getCell(6);
			cell.setCellValue(new HSSFRichTextString("状态"));
			
			cell = sheet.getRow(2).getCell(0);
			cell.setCellValue(new HSSFRichTextString("序号"));
			cell = sheet.getRow(2).getCell(1);
			cell.setCellValue(new HSSFRichTextString("姓名"));
			cell = sheet.getRow(2).getCell(2);
			cell.setCellValue(new HSSFRichTextString("年龄"));
			cell = sheet.getRow(2).getCell(3);
			cell.setCellValue(new HSSFRichTextString("地址"));
			cell = sheet.getRow(2).getCell(4);
			cell.setCellValue(new HSSFRichTextString("手机号"));
			cell = sheet.getRow(2).getCell(5);
			cell.setCellValue(new HSSFRichTextString("图画"));
			cell = sheet.getRow(2).getCell(6);
			cell.setCellValue(new HSSFRichTextString("审批"));
			cell = sheet.getRow(2).getCell(7);
			cell.setCellValue(new HSSFRichTextString("通过"));
			cell = sheet.getRow(2).getCell(8);
			cell.setCellValue(new HSSFRichTextString("驳回"));
			
			// ---------------------------3.合并单元格------------------------------
			sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 8));  // 开始行，结束行，开始列，结束列      包头和尾(从0到8的这9列都合并了)
			sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 5));
			sheet.addMergedRegion(new CellRangeAddress(1, 1, 6, 8));
			
			
			List<UserTest> list=new LinkedList();
			list.add(new UserTest(1, "lili", 17, "新东方", new BigDecimal(1328742894)));
			list.add(new UserTest(2, "mola", 34, "赤道几内亚", new BigDecimal(77689372)));
			list.add(new UserTest(3, "kobe", 24, "拉不拉卡", new BigDecimal(1897789439)));
			list.add(new UserTest(4, "季芳新", 11, "东方帕米尔", new BigDecimal(1991839482)));
			int rowIndex=3;
			Row nRow = null;
		    Cell nCell = null;
			for(UserTest u:list){
				//创建数据行
		        nRow = sheet.createRow(rowIndex++);
		        //从第一列开始写
		        int cellIndex = 0;
		        //ID
		        nCell = nRow.createCell(cellIndex++);
		        nCell.setCellStyle(style);
		        nCell.setCellValue(u.getId());
		        //姓名
		        nCell = nRow.createCell(cellIndex++);
		        nCell.setCellStyle(style);
		        nCell.setCellValue(u.getName());
		        //年龄
		        nCell = nRow.createCell(cellIndex++);
		        nCell.setCellStyle(style);
		        nCell.setCellValue(u.getAge());
		        //地址
		        nCell = nRow.createCell(cellIndex++);
		        nCell.setCellStyle(style);
		        nCell.setCellValue(u.getAddr());
		        //手机号
		        nCell = nRow.createCell(cellIndex++,CellType.NUMERIC);
		        nCell.setCellStyle(style);
		        nCell.setCellValue(u.getPhone().doubleValue());
			}
			
			FileOutputStream fileOut = new FileOutputStream("d:\\转出查询报表1.xls");
			wb.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
	
	

}
