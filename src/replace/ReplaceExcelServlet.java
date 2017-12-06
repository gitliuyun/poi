package replace;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import util.EPlatform;
import util.OSinfo;

public class ReplaceExcelServlet extends HttpServlet {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;


	public static final String fileName = "大兴终端监控情况数据分析-12.5.xlsx";
	public static final String queryDate = "20171205";


	public void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {

		System.out.println("开始时间>>>>>" + new Date().toLocaleString());
		String docsPath = request.getSession().getServletContext()
				.getRealPath("docs");

		String sfileName = fileName;
		String tFileName = "result.xlsx";// 导出Excel文件名 String filePath =
		String sfilePath = docsPath;
		String tFilePath = docsPath;
		if (EPlatform.Windows.equals(OSinfo.getOSname())) {
			sfilePath = sfilePath + "\\" + sfileName;
			tFilePath = tFilePath + "\\" + tFileName;
		} else {
			sfilePath = sfilePath + "/" + sfileName;
			tFilePath = tFilePath + "/" + tFileName;
		}

		Connection conn = null; // 数据库连接
		Statement stmt = null; // 数据库表达式
		ResultSet rs = null; // 结果集
		try {
			// 加载驱动
			Class.forName("com.mysql.jdbc.Driver");
			// 连接到数据库
			conn = DriverManager.getConnection(address, userName, passWord);
			// 获取表达式
			stmt = (Statement) conn.createStatement();

			// 构造 XSSFWorkbook 对象，strPath 传入文件路径
			XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(sfilePath));

            List<List<ExcelReplaceDataVO>> allDatas = new ArrayList<List<ExcelReplaceDataVO>>();

			// 读取第一章表格内容
			for (int m = 0; m <= 1; m ++) {
                List<ExcelReplaceDataVO> datas = new ArrayList<ExcelReplaceDataVO>();
                XSSFSheet sheet = xwb.getSheetAt(m);
                Object value = null;
                XSSFRow row = null;
                XSSFCell cell = null;

                int counter = 0;

                for (int i = sheet.getFirstRowNum(); counter < sheet
                        .getPhysicalNumberOfRows(); i++) {
                    if (i < 3) {
                        continue;
                    }
                    row = sheet.getRow(i);
                    if (row == null) {
                        continue;
                    } else {
                        counter++;
                    }

                    for (int j = 4; j <= 9; j++) {
                        if (j != 4 && j != 9) {
                            continue;
                        }
                        cell = row.getCell(j);
                        if (cell == null) {
                            continue;
                        }
                        DecimalFormat df = new DecimalFormat("0");// 格式化 number
                        // String
                        // 字符
                        SimpleDateFormat sdf = new SimpleDateFormat(
                                "yyyy-MM-dd HH:mm:ss");// 格式化日期字符串
                        DecimalFormat nf = new DecimalFormat("0.00");// 格式化数字
                        switch (cell.getCellType()) {
                            case XSSFCell.CELL_TYPE_STRING:
                                value = cell.getStringCellValue();
                                break;
                            case XSSFCell.CELL_TYPE_NUMERIC:
                                if ("@".equals(cell.getCellStyle()
                                        .getDataFormatString())) {
                                    value = df.format(cell.getNumericCellValue());
                                } else if ("General".equals(cell.getCellStyle()
                                        .getDataFormatString())) {
                                    value = nf.format(cell.getNumericCellValue());
                                } else {
                                    value = sdf.format(HSSFDateUtil.getJavaDate(cell
                                            .getNumericCellValue()));
                                }
                                break;
                            case XSSFCell.CELL_TYPE_BOOLEAN:
                                value = cell.getBooleanCellValue();
                                break;
                            case XSSFCell.CELL_TYPE_BLANK:
                                value = "";
                                break;
                            default:
                                value = cell.toString();
                        }
                        if (value == null || "".equals(value)) {
                            continue;
                        }

                        value = value.toString().trim().substring(0, 8);
                        System.out.println(value);

                        // 插入数据
                        // stmt.executeUpdate("insert into student (name,age) values ('test',20)");

                        rs = stmt.executeQuery("SELECT pos_online_time "
                                + "FROM tbl_base_pos_online_time " + "WHERE"
                                + " date = '" + queryDate + "' " + "and "
                                + "pos_id = '" + Long.parseLong(value.toString(), 16) + "'");
                        while (rs.next()) {
                            System.out.println("在线时长="
                                    + rs.getString("pos_online_time"));
                            ExcelReplaceDataVO vo = new ExcelReplaceDataVO();
                            vo.setRow(i);
                            vo.setColumn(j + 1);
                            vo.setValue(rs.getString("pos_online_time"));
                            datas.add(vo);
                        }

                        rs = stmt.executeQuery("SELECT count(*) as number "
                                + "FROM " + "tbl_base_card_order " + "WHERE "
                                + "term_trans_date = '" + queryDate + "' " + "AND"
                                + " pos_id = '"
                                + Long.parseLong(value.toString(), 16) + "'");

                        while (rs.next()) {
                            System.out.println("实体卡刷卡数=" + rs.getString("number"));
                            ExcelReplaceDataVO vo1 = new ExcelReplaceDataVO();
                            vo1.setRow(i);
                            vo1.setColumn(j + 2);
                            vo1.setValue(rs.getString("number"));
                            datas.add(vo1);
                        }
                        rs = stmt.executeQuery("SELECT count(*) as number "
                                + "FROM tbl_base_qrcode_ride_record trans "
                                + "WHERE " + "trans.trans_recv_date = '"
                                + queryDate + "' " + "AND "
                                + "pos_id = '" + Long.parseLong(value.toString(), 16) + "'");
                        while (rs.next()) {
                            System.out.println("刷卡笔数" + rs.getString("number"));
                            ExcelReplaceDataVO vo2 = new ExcelReplaceDataVO();
                            vo2.setRow(i);
                            vo2.setColumn(j + 3);
                            vo2.setValue(rs.getString("number"));
                            datas.add(vo2);
                        }

                    }
                }

                allDatas.add(datas);

            }

            System.out.println("查询时间>>>>>" + new Date().toLocaleString());

            ExcelUtil.replaceModel2007(allDatas, sfilePath, tFilePath);



			System.out.println("replaceModel时间>>>>>" + new Date().toLocaleString());
			download(tFilePath, response);

			System.out.println("download时间>>>>>" + new Date().toLocaleString());

			rs.close();
			stmt.close();
			conn.close();
		} catch (SQLException se) {
			// 处理 JDBC 错误
			se.printStackTrace();
		} catch (Exception e) {
			// 处理 Class.forName 错误
			e.printStackTrace();
		} finally {
			// 关闭资源
			try {
				if (stmt != null)
					stmt.close();
			} catch (SQLException se2) {
			}// 什么都不做
			try {
				if (conn != null)
					conn.close();
			} catch (SQLException se) {
				se.printStackTrace();
			}
		}

	}

	private void download(String path, HttpServletResponse response) {
		try {
			// path是指欲下载的文件的路径。
			File file = new File(path);
			// 取得文件名。
			String filename = file.getName();
			// 以流的形式下载文件。
			InputStream fis = new BufferedInputStream(new FileInputStream(path));
			byte[] buffer = new byte[fis.available()];
			fis.read(buffer);
			fis.close();
			// 清空response
			response.reset();
			// 设置response的Header
			response.addHeader("Content-Disposition", "attachment;filename="
					+ new String(filename.getBytes()));
			response.addHeader("Content-Length", "" + file.length());
			OutputStream toClient = new BufferedOutputStream(
					response.getOutputStream());
			response.setContentType("application/vnd.ms-excel;charset=gb2312");
			toClient.write(buffer);
			toClient.flush();
			toClient.close();
		} catch (IOException ex) {
			ex.printStackTrace();
		}
	}


    @Override
    protected void doPost(HttpServletRequest request,
                          HttpServletResponse response) throws ServletException, IOException {
        doGet(request, response);
    }
}