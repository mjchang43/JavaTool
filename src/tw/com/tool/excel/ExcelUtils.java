package tw.com.tool.excel;

import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.TreeMap;
import java.util.UUID;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import jodd.bean.BeanUtil;
import jodd.typeconverter.impl.IntegerConverter;
import jodd.util.StringUtil;

public class ExcelUtils {

	public Logger log = LoggerFactory.getLogger(ExcelUtils.class);

	  /**
	   * 回傳錯誤代碼
	   */
	  private static final String ERROR = "err";

	  /**
	   * 回傳資料代碼
	   */
	  private static final String DATA = "data";

	  /**
	   * 累計錯誤次數
	   */
	  private int errCnt = 0;

	  /**
	   * 錯誤次數上限
	   */
	  private int errLimit = 0;

	  /**
	   * 錯誤訊息字串
	   */
	  private StringBuffer sb = new StringBuffer();

	  /*錯誤Map*/
	  //Map<Integer, Map<String,String>> errMap = new TreeMap<Integer,Map<String, String>>();
	  TreeMap<Integer, String> errMap = new TreeMap<Integer, String>();

	  /**
	   * Excel 內容處理 (重複資料已過濾)
	   *
	   * @param xls         Excel 檔案
	   * @param xlsName     Excel 檔案名稱
	   * @param columnlt    DB 欄位檢核設定條件
	   * @param deflt       DB 欄位預設值設定條件
	   * @param ao          DB Table Bean
	   * @param iHeaderLine excel表頭行數
	   * @return
	   * @throws FileNotFoundException
	   * @throws IOException
	   * @throws IllegalAccessException
	   * @throws InvocationTargetException
	   * @throws InstantiationException
	   * @throws ClassNotFoundException
	   * @version 1.0
	   * <p>[Change history]</p>
	   * <p>Initial: (2013/4/17) 初始版本 by Hank</p>
	   */

	  public HashMap excelHandle(InputStream xls, String xlsName, List<Object[]> columnlt, List<String[]> deflt, Object ao, int iHeaderLine) throws Exception {
	    return excelHandle(xls, xlsName, columnlt, deflt, ao, iHeaderLine, "");
	  }

	  /**
	   * Excel 內容處理 (重複資料已過濾)
	   *
	   * @param xls         Excel 檔案
	   * @param xlsName     Excel 檔案名稱
	   * @param columnlt    DB 欄位檢核設定條件
	   * @param deflt       DB 欄位預設值設定條件
	   * @param ao          DB Table Bean
	   * @param iHeaderLine excel表頭行數
	   * @param sheetName   Sheet Name
	   * @return
	   * @throws FileNotFoundException
	   * @throws IOException
	   * @throws IllegalAccessException
	   * @throws InvocationTargetException
	   * @throws InstantiationException
	   * @throws ClassNotFoundException
	   * @version 1.0
	   * <p>[Change history]</p>
	   * <p>Initial: (2013/4/17) 初始版本 by Hank</p>
	   */
	  public HashMap excelHandle(InputStream xls, String xlsName, List<Object[]> columnlt, List<String[]> deflt, Object ao, int iHeaderLine, String sheetName) throws Exception {

	    HashMap hm = new HashMap();
	    List exllt = new ArrayList();
	    LinkedHashSet<HashMap> lhs = new LinkedHashSet<>();

	    Workbook workbook;

	    if (xlsName.toLowerCase().endsWith(".xls")) {
	      // 2003版本Excel(.xls)
	      workbook = new HSSFWorkbook(xls);
	    } else if (xlsName.toLowerCase().endsWith(".xlsx")) {
	      // 2007版本Excel或更高版本(.xlsx)
	      workbook = new XSSFWorkbook(xls);
	    } else {
	      throw new Exception("檔案格式非Excel!");
	      //hm.put( ERROR, "檔案格式非Excel!" );
	      //return hm;
	    }

	    Sheet sheet = null;
	    if (StringUtil.isBlank(sheetName)) {
	      sheet = workbook.getSheetAt(0);
	    } else {
	      sheet = workbook.getSheet(sheetName);
	    }

	    if (sheet == null) {
	      throw new Exception("檔案無[" + sheetName + "]頁籤!");
	      //hm.put( ERROR, "檔案無[" + sheetName + "]頁籤!" );
	      //return hm;
	    }

	    int rowNum = sheet.getLastRowNum();

	    for (int i = iHeaderLine; i <= rowNum; i++) {
	      Row row = sheet.getRow(i);
	      HashMap cellhm = new HashMap();

	      //Excel 內容欄位處理
	      String errLine = columnValue(cellhm, columnlt, row, (i + 1));

	      if ("break".equalsIgnoreCase(errLine)) {
	        sb.append("\n<br/>已達檢查錯誤次數上限[" + errLimit + "]個，請重新檢視上傳檔案格式是否正確。");
	        break;
	      }

	      if ("false".equalsIgnoreCase(errLine)) {// 整行沒錯誤才寫入加入list
	        if (!cellhm.isEmpty()) {
	          // 預設欄位值處理(Excel 內容不存在資料,但DB欄位需給 DEFAULT) (先執行避免Excel相同資料被排除)
	          defColumnValue(cellhm, deflt, i - iHeaderLine + 1);
	          lhs.add(cellhm);
	        }
	      }
	    }

	    //過濾重複資料後,再處理 DB 欄位需給 DEFAULT及轉成 DB Bean
	    if (!lhs.isEmpty()) {
//	            int seq = 1;
	      for (HashMap lhm : lhs) {
	        //預設欄位值處理(Excel 內容不存在資料,但DB欄位需給 DEFAULT)
//	                defColumnValue( lhm, deflt , seq );

	        //將資料存入 DB Bean
	        Object aoN = Class.forName(ao.getClass().getName()).newInstance();
	        BeanUtil.populateBean(aoN, lhm);
	        exllt.add(aoN);
	        lhm.clear();
//	                seq++;
	      }

	      lhs.clear();
	    }

	    if (!exllt.isEmpty()) {
	      hm.put(DATA, exllt);
	    }
	    if (sb.length() > 0) {
	      hm.put(ERROR, sb.toString());
	      sb.delete(0, sb.length());
	    }

	    return hm;
	  }

	  public HashMap excelHandleIgnoreErr(InputStream xls, String xlsName, List<Object[]> columnlt, List<String[]> deflt,
	                                      Object ao, int iHeaderLine, String sheetName)
	      throws Exception {

	    HashMap hm = new HashMap();
	    List exllt = new ArrayList();
	    LinkedHashSet<HashMap> lhs = new LinkedHashSet<>();

	    Workbook workbook;

	    if (xlsName.toLowerCase().endsWith(".xls")) {
	      // 2003版本Excel(.xls)
	      workbook = new HSSFWorkbook(xls);
	    } else if (xlsName.toLowerCase().endsWith(".xlsx")) {
	      // 2007版本Excel或更高版本(.xlsx)
	      workbook = new XSSFWorkbook(xls);
	    } else {
	      throw new Exception("檔案格式非Excel!");
	      // hm.put( ERROR, "檔案格式非Excel!" );
	      // return hm;
	    }

	    Sheet sheet = null;
	    if (StringUtil.isBlank(sheetName)) {
	      sheet = workbook.getSheetAt(0);
	    } else {
	      sheet = workbook.getSheet(sheetName);
	    }

	    if (sheet == null) {
	      throw new Exception("檔案無[" + sheetName + "]頁籤!");
	      // hm.put( ERROR, "檔案無[" + sheetName + "]頁籤!" );
	      // return hm;
	    }

	    int rowNum = sheet.getLastRowNum();

	    for (int i = iHeaderLine; i <= rowNum; i++) {
	      Row row = sheet.getRow(i);
	      HashMap cellhm = new HashMap();

	      // Excel 內容欄位處理
	      String errLine = columnValue(cellhm, columnlt, row, (i + 1));

	      if ("break".equalsIgnoreCase(errLine)) {
	        sb.append("\n<br/>已達檢查錯誤次數上限[" + errLimit + "]個，請重新檢視上傳檔案格式是否正確。");
	        break;
	      }

	      if ("false".equalsIgnoreCase(errLine)) {// 整行沒錯誤才寫入加入list
	        if (!cellhm.isEmpty()) {
	          // 預設欄位值處理(Excel 內容不存在資料,但DB欄位需給 DEFAULT)
	          // (先執行避免Excel相同資料被排除)
	          defColumnValue(cellhm, deflt, i - iHeaderLine + 1);
	          cellhm.put("rowNum", (i + 1));
	          lhs.add(cellhm);
	        }
	      }
	    }

	    // 過濾重複資料後,再處理 DB 欄位需給 DEFAULT及轉成 DB Bean
	    if (!lhs.isEmpty()) {
	      // int seq = 1;
	      for (HashMap lhm : lhs) {
	        // 預設欄位值處理(Excel 內容不存在資料,但DB欄位需給 DEFAULT)
	        // defColumnValue( lhm, deflt , seq );

	        // 將資料存入 DB Bean
	        Object aoN = Class.forName(ao.getClass().getName()).newInstance();
	        BeanUtil.populateBean(aoN, lhm);
	        exllt.add(aoN);
	        lhm.clear();
	        // seq++;
	      }

	      lhs.clear();
	    }

	    if (!exllt.isEmpty()) {
	      hm.put(DATA, exllt);
	    }
	    if (sb.length() > 0) {
	      hm.put(ERROR, sb.toString());
	      sb.delete(0, sb.length());
	    }

	    hm.put("errMap", errMap);

	    return hm;
	  }


	  /**
	   * Excel 每行資料處理
	   *
	   * @param cellhm   每筆資料或錯誤訊息
	   * @param columnlt DB 欄位檢核設定條件
	   * @param row      Excel 該行資料
	   * @param rowNum   Excel 行數
	   * @return errLine 本行是否有錯誤
	   * @version 1.0
	   * <p>[Change history]</p>
	   * <p>Initial: (2013/4/17) 初始版本 by Hank</p>
	   */
	  private String columnValue(HashMap cellhm, List<Object[]> columnlt, Row row, int rowNum) {
	    boolean errLine = false;// 本行是否有錯誤

	    if (columnlt != null && !columnlt.isEmpty()) {
	      Cell cell = null;
	      for (int i = 0; i < columnlt.size(); i++) {
	        // errLimit 為0 表示為不限制錯誤次數
	        if (errLimit != 0) {
	          if (errCnt >= errLimit) return "break";
	        }

	        Object[] aryValue = columnlt.get(i);
	        String cellValue = "";
	        if (row != null && row.getCell(i) != null) {
	          cell = row.getCell(i);
	          cell.setCellType(Cell.CELL_TYPE_STRING);
	          cellValue = cell.getStringCellValue().trim();
	        }

	        switch ((Integer) aryValue[1]) {
	          case ExcelConstant.VARCHAR:
	            if (xlsVARCHAR(cellhm, cellValue, aryValue, i, rowNum)) {
	              errLine = true;
	            }
	            break;
	          case ExcelConstant.INT:
	            if (xlsINT(cellhm, cellValue, aryValue, i, rowNum)) {
	              errLine = true;
	            }
	            break;
	          case ExcelConstant.DATE:
	            if (xlsDATE(cellhm, cellValue, aryValue, i, rowNum)) {
	              errLine = true;
	            }
	            break;
	          default:
	            break;
	        }
	      }
	    }

	    return errLine ? "true" : "false";
	  }

	  /**
	   * 特殊處理 非 Excel 內容值,但DB 欄位需給預設值
	   *
	   * @param cellhm 每筆資料或錯誤訊息
	   * @param deflt  DB 欄位預設值設定條件
	   * @version 1.0
	   * <p>[Change history]</p>
	   * <p>Initial: (2013/4/17) 初始版本 by Hank</p>
	   */
	  private void defColumnValue(HashMap cellhm, List<String[]> deflt, int seq) {
	    if (deflt != null && !deflt.isEmpty()) {
	      for (int i = 0; i < deflt.size(); i++) {
	        String[] aryValue = deflt.get(i);
	        Object def = new Object();
	        if ("uid".equalsIgnoreCase(aryValue[1].toLowerCase())) {
	          def = UUID.randomUUID().toString().replaceAll("-", "").replaceAll("-", "");
	        } else if ("def".equalsIgnoreCase(aryValue[1].toLowerCase())) {
	          def = aryValue[2];
	        } else if ("int".equalsIgnoreCase(aryValue[1].toLowerCase())) {
	          def = aryValue[2];
	        } else if ("date".equalsIgnoreCase(aryValue[1].toLowerCase())) {
	          if (aryValue.length >= 3 && !StringUtil.isBlank(aryValue[2])) {
	            def = aryValue[2];
	          } else {
	            def = new Date();
	          }
	        } else if ("seq".equalsIgnoreCase(aryValue[1].toLowerCase())) {
	          def = seq;
	        }
	        
	        cellhm.put(aryValue[0], def);
	      }
	    }
	  }

	/***************************************************************************************************************************/

	  /**
	   * 欄位為VARCHAR處理
	   *
	   * @param cellhm    每筆資料或錯誤訊息
	   * @param cellValue Excel 該欄位值
	   * @param aryValue  該筆Record設定檢核資料
	   * @param colNum    列
	   * @param rowNum    行
	   * @return 檢核是否有錯
	   * @version 1.0
	   * <p>[Change history]</p>
	   * <p>Initial: (2013/4/18) 初始版本 by Hank</p>
	   */
	  private boolean xlsVARCHAR(HashMap cellhm, String cellValue, Object[] aryValue, int colNum, int rowNum) {
	    boolean isErr = false;
	    if (!StringUtil.isBlank(cellValue)) {
	      //   isErr = chkLength( cellhm, cellValue, (Integer)aryValue[2], colNum, rowNum );
	      isErr = chkLength(cellhm, cellValue, (Integer) aryValue[2], colNum, rowNum, StringUtil.toSafeString(aryValue[4]));
	    } else if (!(Boolean) aryValue[3]) {
	      // isErr = chkISEMPTY( cellhm, cellValue, colNum, rowNum );
	      isErr = chkISEMPTY(cellhm, cellValue, colNum, rowNum, StringUtil.toSafeString(aryValue[4]));
	    }

	    if (!isErr) {
	      cellhm.put(aryValue[0], cellValue);
	    }
	    return isErr;
	  }

	  /**
	   * 欄位為 INT 處理
	   *
	   * @param cellhm    每筆資料或錯誤訊息
	   * @param cellValue Excel 該欄位值
	   * @param aryValue  該筆Record設定檢核資料
	   * @param colNum    列
	   * @param rowNum    行
	   * @return 檢核是否有錯
	   * @version 1.0
	   * <p>[Change history]</p>
	   * <p>Initial: (2013/4/18) 初始版本 by Hank</p>
	   */
	  private boolean xlsINT(HashMap cellhm, String cellValue, Object[] aryValue, int colNum, int rowNum) {
	    boolean isErr = false;
	    if (!StringUtil.isBlank(cellValue)) {
	      //    isErr = chkLength( cellhm, cellValue, (Integer)aryValue[2], colNum, rowNum );
	      isErr = chkLength(cellhm, cellValue, (Integer) aryValue[2], colNum, rowNum, StringUtil.toSafeString(aryValue[4]));
	      IntegerConverter intC = new IntegerConverter();
	      if (!isErr) {
	        try {
	          cellhm.put(aryValue[0], intC.convert(cellValue));
	        } catch (Exception e) {
	          //  sb.append( "第 " + rowNum + " 列,第 " + (colNum+1) + " 欄,數字格式錯誤;\n<br/>" );
	          sb.append("第 " + rowNum + " 列,【" + StringUtil.toSafeString(aryValue[4]) + "】數字格式錯誤;\n<br/>");
	          errCnt++;
	          isErr = true;
	          errMap.put(rowNum, "第 " + rowNum + " 列,【" + StringUtil.toSafeString(aryValue[4]) + "】數字格式錯誤");
	          log.info("", e);
	        }
	      }
	    } else {
	      if (!(Boolean) aryValue[3]) {
	        //     isErr = chkISEMPTY( cellhm, cellValue, colNum, rowNum );
	        isErr = chkISEMPTY(cellhm, cellValue, colNum, rowNum, StringUtil.toSafeString(aryValue[4]));
	      }

	      if (!isErr) {
	        cellhm.put(aryValue[0], 0);
	      }
	    }
	    return isErr;
	  }

	  /**
	   * 欄位為 Date 處理
	   *
	   * @param cellhm    每筆資料或錯誤訊息
	   * @param cellValue Excel 該欄位值
	   * @param aryValue  該筆Record設定檢核資料
	   * @param colNum    列
	   * @param rowNum    行
	   * @return 檢核是否有錯
	   * @version 1.0
	   * <p>[Change history]</p>
	   * <p>Initial: (2013/4/18) 初始版本 by Hank</p>
	   */
	  private boolean xlsDATE(HashMap cellhm, String cellValue, Object[] aryValue, int colNum, int rowNum) {
	    boolean isErr = false;
	    if (!StringUtil.isBlank(cellValue)) {
	      cellValue = cellValue.replace("/", "-");
	      try {
	        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
	        sdf.setLenient(false);
	        Date dt = sdf.parse(cellValue);

	        cellhm.put(aryValue[0], dt);
	      } catch (Exception e) {
	        //  sb.append( "第 " + rowNum + " 列,第 " + (colNum+1) + " 欄,欄位日期格式錯誤;\n<br/>" );
	        sb.append("第 " + rowNum + " 列,【" + StringUtil.toSafeString(aryValue[4]) + "】欄位日期格式錯誤;\n<br/>");
	        errCnt++;
	        isErr = true;
	        errMap.put(rowNum, "第 " + rowNum + " 列,【" + StringUtil.toSafeString(aryValue[4]) + "】欄位日期格式錯誤");
	        log.info("", e);
	      }
	    } else {
	      if (!(Boolean) aryValue[3]) {
	        //  isErr = chkISEMPTY( cellhm, cellValue, colNum, rowNum );
	        isErr = chkISEMPTY(cellhm, cellValue, colNum, rowNum, StringUtil.toSafeString(aryValue[4]));
	      }

	      if (!isErr) {
	        cellhm.put(aryValue[0], null);
	      }
	    }
	    return isErr;
	  }
	/***************************************************************************************************************************/
	  /**
	   * 長度檢核
	   *
	   * @param cellhm    每筆資料或錯誤訊息
	   * @param cellValue Excel 該欄位值
	   * @param maxLength 最大長度
	   * @param colNum    列
	   * @param rowNum    行
	   * @return
	   * @version 1.0
	   * <p>[Change history]</p>
	   * <p>Initial: (2013/4/18) 初始版本 by Hank</p>
	   */
	  private boolean chkLength(HashMap cellhm, String cellValue, int maxLength, int colNum, int rowNum) {
	    boolean isErr = false;
	    if (cellValue.length() > maxLength) {
	      sb.append("第 " + rowNum + " 列,第 " + (colNum + 1) + " 欄,欄位長度超過 " + maxLength + ";\n<br/>");
	      isErr = true;
	      errMap.put(rowNum, "第 " + rowNum + " 列,第 " + (colNum + 1) + " 欄,欄位長度超過 " + maxLength);
	      errCnt++;
	    }

	    return isErr;
	  }

	  private boolean chkLength(HashMap cellhm, String cellValue, int maxLength, int colNum, int rowNum, String colName) {
	    boolean isErr = false;
	    if (cellValue.length() > maxLength) {
	      //  sb.append( "第 " + rowNum + " 列,第 " + (colNum+1) + " 欄,欄位長度超過 " + maxLength + ";\n<br/>");
	      sb.append("第 " + rowNum + " 列,【" + colName + "】欄位長度超過 " + maxLength + ";\n<br/>");
	      isErr = true;

	      errMap.put(rowNum, "第 " + rowNum + " 列,【" + colName + "】欄位長度超過 " + maxLength);

	      errCnt++;
	    }

	    return isErr;
	  }

	  /**
	   * 是否允許空值檢核
	   *
	   * @param cellhm    每筆資料或錯誤訊息
	   * @param cellValue Excel 該欄位值
	   * @param colNum    列
	   * @param rowNum    行
	   * @return
	   * @version 1.0
	   * <p>[Change history]</p>
	   * <p>Initial: (2013/4/18) 初始版本 by Hank</p>
	   */
	  private boolean chkISEMPTY(HashMap cellhm, String cellValue, int colNum, int rowNum) {
	    boolean isErr = false;

	    if (StringUtil.isBlank(cellValue)) {
	      sb.append("第 " + rowNum + " 列,第 " + (colNum + 1) + " 欄,不允許為空值;\n<br/>");
	      isErr = true;
	      errMap.put(rowNum, "第 " + rowNum + " 列,第 " + (colNum + 1) + " 欄,不允許為空值");
	      errCnt++;
	    }

	    return isErr;
	  }

	  private boolean chkISEMPTY(HashMap cellhm, String cellValue, int colNum, int rowNum, String colName) {
	    boolean isErr = false;
	    log.debug("colName:" + colName);
	    if (StringUtil.isBlank(cellValue)) {
	      log.debug(colName + "isBlank");
	      sb.append("第 " + rowNum + " 列,【" + colName + "】不允許為空值;\n<br/>");
	      isErr = true;
	      errMap.put(rowNum, "第 " + rowNum + " 列,【" + colName + "】不允許為空值");
	      errCnt++;
	    }

	    return isErr;
	  }


	  public int getErrLimit() {
	    return errLimit;
	  }

	  public void setErrLimit(int errLimit) {
	    this.errLimit = errLimit;
	  }
	  
	  public List<Object> getXls(InputStream input){
			
		try{
			List<Object> result = new LinkedList<Object>();
			HSSFWorkbook wb = new HSSFWorkbook(input);
			HSSFSheet sheet = wb.getSheetAt(0);
			int rows = sheet.getPhysicalNumberOfRows();
			
			for (int r = 0; r < rows; r++) {
				  
				HSSFRow row = sheet.getRow(r);
				int cells = row.getPhysicalNumberOfCells();
				for (int c = 0; c < cells; c++) {
					
					HSSFCell cell = row.getCell(c);
					if (c == 0 && (cell == null || cell.getCellType() == HSSFCell.CELL_TYPE_BLANK )) {
						System.out.println("行:" + c + "首行為空，停止讀取excel檔案");
						break;
					}
					if (null != cell && cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
						if(StringUtil.isNotBlank(cell.getStringCellValue()))
							result.add(cell.getStringCellValue());
					} 
				}
			}
			return result;
		}catch(Exception e){
			e.printStackTrace();
		}
		
		return null;
	}
	
	public List<Object> getXlsx(InputStream input){
		
		try{
			List<Object> result = new LinkedList<Object>();
			XSSFWorkbook wb = new XSSFWorkbook(input);
			XSSFSheet sheet = wb.getSheetAt(0);
			int rows = sheet.getPhysicalNumberOfRows();
			
			for (int r = 0; r < rows; r++) {
				  
				XSSFRow row = sheet.getRow(r);
				int cells = row.getPhysicalNumberOfCells();
				for (int c = 0; c < cells; c++) {
					
					XSSFCell cell = row.getCell(c);
					if (c == 0 && (cell == null || cell.getCellType() == XSSFCell.CELL_TYPE_BLANK )) {
						System.out.println("行:" + c + "首行為空，停止讀取excel檔案");
						break;
					}
					if (null != cell && cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
						if(StringUtil.isNotBlank(cell.getStringCellValue()))
							result.add(cell.getStringCellValue());
					} 
				}
			}
			return result;
		}catch(Exception e){
			e.printStackTrace();
		}
		
		return null;
	}
}
