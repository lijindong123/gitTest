package poiTest;

import java.net.URLEncoder;

import javax.servlet.http.HttpServletRequest;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import com.sun.xml.internal.messaging.saaj.packaging.mime.internet.MimeUtility;

//import com.sun.xml.internal.messaging.saaj.packaging.mime.internet.MimeUtility;
//import com.sun.xml.internal.ws.util.StringUtils;

public class ExportFileNameUtils {
	
	/**
	 * 
	 * <br>
	 * <b>功能：</b>设置下载文件中文件的名称<br>
	 * <b>作者：</b>yixq<br>
	 * <b>@param filename
	 * <b>@param request
	 * <b>@return</b>
	 */
	public static String encodeFilename(String filename,
			HttpServletRequest request) {
		/**
		 * 获取客户端浏览器和操作系统信息 在IE浏览器中得到的是：User-Agent=Mozilla/4.0 (compatible; MSIE
		 * 6.0; Windows NT 5.1; SV1; Maxthon; Alexa Toolbar)
		 * 在Firefox中得到的是：User-Agent=Mozilla/5.0 (Windows; U; Windows NT 5.1;
		 * zh-CN; rv:1.7.10) Gecko/20050717 Firefox/1.0.6
		 */
		String agent = request.getHeader("USER-AGENT");
		try {
			if ((agent != null) && (-1 != agent.indexOf("MSIE"))) {
				String newFileName = URLEncoder.encode(filename, "UTF-8");
				newFileName =  newFileName.replaceAll("+", "%20");
				if (newFileName.length() > 150) {
					newFileName = new String(filename.getBytes("GB2312"),
							"ISO8859-1");
					newFileName = newFileName.replaceAll(" ", "%20");
				}
				return newFileName;
			}
			if ((agent != null) && (-1 != agent.indexOf("Mozilla")))
				return MimeUtility.encodeText(filename, "UTF-8", "B");
 
			return filename;
		} catch (Exception ex) {
			return filename;
		}
	}
 
	/**
	 * 
	 * <br>
	 * <b>功能：</b>列头样式<br>
	 * <b>作者：</b>yixq<br>
	 * <b>@param wb
	 * <b>@return</b>
	 */
	public static HSSFCellStyle initColumnHeadStyle(HSSFWorkbook wb) {
		HSSFCellStyle columnHeadStyle = wb.createCellStyle();
		HSSFFont columnHeadFont = wb.createFont();
		columnHeadFont.setFontName("宋体");
		columnHeadFont.setFontHeightInPoints((short) 10);
		//columnHeadFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		columnHeadStyle.setFont(columnHeadFont);
		columnHeadStyle.setAlignment(HorizontalAlignment.CENTER);// 左右居中
		columnHeadStyle.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
		columnHeadStyle.setLocked(true);
		columnHeadStyle.setWrapText(true);
		//设置边框颜色和粗细
		columnHeadStyle.setBorderBottom(BorderStyle.THIN);
		columnHeadStyle.setBottomBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
		columnHeadStyle.setBorderLeft(BorderStyle.THIN);
		columnHeadStyle.setLeftBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
		columnHeadStyle.setBorderRight(BorderStyle.THIN);
		columnHeadStyle.setRightBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
		columnHeadStyle.setBorderTop(BorderStyle.THIN);
		columnHeadStyle.setTopBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
		// 设置单元格的背景颜色（单元格的样式会覆盖列或行的样式）
		//columnHeadStyle.setFillForegroundColor();
		return columnHeadStyle;
	}
 
	 /**
	 * 
	 * <br>
	 * <b>功能：</b>单元格的默认样式<br>
	 * <b>作者：</b>yixq<br>
	 * <b>@param wb
	 * <b>@return</b>
	 */
	public static HSSFCellStyle initColumnCenterstyle(HSSFWorkbook wb) {
		HSSFFont font = wb.createFont();
		font.setFontName("宋体");
		font.setFontHeightInPoints((short) 10);
		HSSFCellStyle centerstyle = wb.createCellStyle();
		centerstyle.setFont(font);
		centerstyle.setAlignment(HorizontalAlignment.CENTER);// 左右居中
		centerstyle.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
		centerstyle.setWrapText(true);
		//设置边框颜色和粗细
		centerstyle.setBorderBottom(BorderStyle.THIN);
		centerstyle.setBottomBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
		centerstyle.setBorderLeft(BorderStyle.THIN);
		centerstyle.setLeftBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
		centerstyle.setBorderRight(BorderStyle.THIN);
		centerstyle.setRightBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
		centerstyle.setBorderTop(BorderStyle.THIN);
		centerstyle.setTopBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
		return centerstyle;

	}
}
