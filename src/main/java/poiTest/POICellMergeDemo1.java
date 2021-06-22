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

public class POICellMergeDemo1 {

	public static void main(String[] args) {

		HSSFWorkbook wb = new HSSFWorkbook();// 创建一个Excel文件
		HSSFSheet sheet = wb.createSheet("银行存余额表(1)");// 创建一个Excel的Sheet
		HSSFCellStyle boderStyle=wb.createCellStyle();
		
		// 定义样式
		HSSFCellStyle cellStyleCenter = ExportFileNameUtils.initColumnHeadStyle(wb);//表头样工
		HSSFCellStyle cellStyleRight = ExportFileNameUtils.initColumnCenterstyle(wb);//单元格样式
		HSSFCellStyle cellStyleLeft = ExportFileNameUtils.initColumnCenterstyle(wb);
		cellStyleRight.setAlignment(HorizontalAlignment.CENTER);//右对齐
		cellStyleLeft.setAlignment(HorizontalAlignment.LEFT);//左对齐
		
		
		// 设置列宽
		sheet.setColumnWidth(0, 7200);
		sheet.setColumnWidth(1, 5000);
		sheet.setColumnWidth(2, 5000);
		sheet.setColumnWidth(3, 5000);
		sheet.setColumnWidth(4, 5000);
		sheet.setColumnWidth(5, 5000);
		try {
			HSSFRow row = null;
			HSSFCell cell = null;
			// ---------------------------1.初始化带边框的表头------------------------------
			for (int i = 0; i < 5; i++) {
				row = sheet.createRow(i);
				for (int j = 0; j <= 5; j++) {
					cell = row.createCell(j);
				}
			}
			// ---------------------------2.指定单元格填充数据------------------------------
			cell = sheet.getRow(0).getCell(0);
			cell.setCellValue(new HSSFRichTextString("银行存余额表"));
			cell = sheet.getRow(1).getCell(0);
			cell.setCellValue(new HSSFRichTextString("2015-08-05"));
			cell = sheet.getRow(2).getCell(0);
			cell.setCellValue(new HSSFRichTextString("开户行"));
			cell = sheet.getRow(2).getCell(1);
			cell.setCellValue(new HSSFRichTextString("活期"));
			cell = sheet.getRow(2).getCell(3);
			cell.setCellValue(new HSSFRichTextString("定期"));
			cell = sheet.getRow(2).getCell(5);
			cell.setCellValue(new HSSFRichTextString("存款合计"));
			cell = sheet.getRow(3).getCell(1);
			cell.setCellValue(new HSSFRichTextString(" "));
			cell = sheet.getRow(3).getCell(4);
			cell.setCellValue(new HSSFRichTextString("折合本位币合计"));
			cell = sheet.getRow(4).getCell(1);
			cell.setCellValue(new HSSFRichTextString("人民币"));
			cell = sheet.getRow(4).getCell(2);
			cell.setCellValue(new HSSFRichTextString("折合本位币合计"));

			// ---------------------------3.合并单元格------------------------------
			sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));// 开始行，结束行，开始列，结束列
			sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 5));
			sheet.addMergedRegion(new CellRangeAddress(2, 4, 0, 0));
			sheet.addMergedRegion(new CellRangeAddress(2, 3, 1, 2));
			sheet.addMergedRegion(new CellRangeAddress(2, 2, 3, 4));
			sheet.addMergedRegion(new CellRangeAddress(3, 4, 4, 4));
			sheet.addMergedRegion(new CellRangeAddress(2, 4, 5, 5));
			FileOutputStream fileOut = new FileOutputStream("d:\\银行存款余额表.xls");
			wb.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
