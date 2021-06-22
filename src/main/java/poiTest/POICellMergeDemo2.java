package poiTest;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

public class POICellMergeDemo2 {

	public static void main(String[] args) {

		HSSFWorkbook wb = new HSSFWorkbook();// 创建一个Excel文件
		HSSFSheet sheet = wb.createSheet("表1");// 创建一个Excel的Sheet
		HSSFCellStyle boderStyle=wb.createCellStyle();
		//垂直居中
		boderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		boderStyle.setAlignment(HorizontalAlignment.CENTER); // 创建一个居中格式
		//设置一个边框
		boderStyle.setBorderTop(BorderStyle.THIN);
		
		
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
			for (int i = 0; i < 8; i++) {
				row = sheet.createRow(i);
				for (int j = 0; j <= 8; j++) {
					cell = row.createCell(j);
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
			sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 8));// 开始行，结束行，开始列，结束列      包头和尾(从0到8的这9列都合并了)
			sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 5));
			sheet.addMergedRegion(new CellRangeAddress(1, 1, 6, 8));
			FileOutputStream fileOut = new FileOutputStream("d:\\转出查询报表.xls");
			wb.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
