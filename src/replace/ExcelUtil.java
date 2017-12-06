package replace;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
	/**
	 * 替换Excel模板文件内容
	 * 
	 * @param datas
	 *            文档数据
	 * @param sourceFilePath
	 *            Excel模板文件路径
	 * @param targetFilePath
	 *            Excel生成文件路径
	 */
	public static boolean replaceModel(List<ExcelReplaceDataVO> datas,
			String sourceFilePath, String targetFilePath) {
		boolean bool = true;
		try {
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(
					sourceFilePath));
			HSSFWorkbook wb = new HSSFWorkbook(fs);
			HSSFSheet sheet = wb.getSheetAt(0);
			for (ExcelReplaceDataVO data : datas) {
				// 获取单元格内容
				HSSFRow row = sheet.getRow(data.getRow());
				HSSFCell cell = row.getCell((short) data.getColumn());
				// 写入单元格内容
				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				// cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue(data.getValue());
			}
			// 输出文件
			FileOutputStream fileOut = new FileOutputStream(targetFilePath);
			wb.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			bool = false;
			e.printStackTrace();
		}
		return bool;
	}
	
	public static boolean replaceModel2007(List<List<ExcelReplaceDataVO>> allDatas,
			String sourceFilePath, String targetFilePath) {
		boolean bool = true;
		try {
			InputStream inp = new FileInputStream(sourceFilePath);
			XSSFWorkbook wb = new XSSFWorkbook(inp);

			for (int i = 0; i < allDatas.size(); i ++) {
                // 读取第一章表格内容
                List<ExcelReplaceDataVO> sdata = allDatas.get(i);
                XSSFSheet sheet = wb.getSheetAt(i);
                for (ExcelReplaceDataVO data : sdata) {
                    // 获取单元格内容
                    XSSFRow row = sheet.getRow(data.getRow());
                    XSSFCell cell = row.getCell((short) data.getColumn());
                    // 写入单元格内容
                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                    // cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(data.getValue());
                }
            }

			// 输出文件
			FileOutputStream fileOut = new FileOutputStream(targetFilePath);
			wb.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			bool = false;
			e.printStackTrace();
		}
		return bool;
	}
}
