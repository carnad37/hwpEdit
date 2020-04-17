package util.excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelParse {
	
	private XSSFWorkbook workbook = null;
	private List<Map<String, String>> dataList = null;
	private List<Integer> dateColNum = null;
	private int sRow = 0;
	private int sCol = 0;
	private int eRow = 0;
	private int eCol = 0;	
	private int mRow = 0;
	
	interface Callback {
		void rowFuncMapping();
	}
	
	
	public ExcelParse(FileInputStream file) throws IOException {
		this.workbook = new XSSFWorkbook(file);
		dateColNum = new ArrayList<Integer>();
		dataList = new ArrayList<Map<String, String>>();
	}

//	public Parse(FileInputStream file, ) throws IOException {
//		this.workbook = new XSSFWorkbook(file);
//		this.row = row;
//		this.col = column;
//	}
	
	public XSSFWorkbook getWorkBook() {
		return this.workbook;
	}
	
	public void setRow(int sRow, int eRow) {
		this.mRow = sRow;
		this.sRow = sRow + 1;
		this.eRow = eRow;
	}
	
	public void setCol(int sCol, int eCol) {
		this.sCol = sCol;
		this.eCol = eCol;
	}
	
	public void setDateCol(int colNum) {
		if (dateColNum.contains(colNum)) {
			return;
		}
		dateColNum.add(colNum);
	}
	
	public void setDateCol(List<Integer> colNumList) {
		for (Integer num : colNumList) {
			if (dateColNum.contains(num)) {
				continue;
			}
			dateColNum.add(num);
		}
	}
	
	//임시로 파싱구조는 하드코딩
	public List<Map<String, String>> parse(int sheetNumber) {
		//맵에다가 파싱한 데이터를 삽입
		
		dataList.clear();
		
		XSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		rowLoop:for (int i = sRow; i <= eRow; i++) {
			XSSFRow row = sheet.getRow(i);
			if (row != null) {
				boolean breakFlag = false;
				Map<String, String> colMap = new HashMap<String, String>();
				colLoop:for (int j = sCol; j <= eCol; j++) {
					XSSFCell cell = row.getCell(j);
					String value = "";
					if (cell != null) {
						switch (cell.getCellType()) {
                        case FORMULA:
                            value=cell.getCellFormula();
                            break;
                        case NUMERIC:
                        	if (dateColNum.contains(j)) {
                        		Date date = cell.getDateCellValue();
                        		value = new SimpleDateFormat("yyyy-MM-dd").format(date);
                        	} else {
                        		//숫자 소수점으로 새자
                    			value = String.format("%,d", (int)cell.getNumericCellValue());               		
							}
                            break;
                        case STRING:
                            value=cell.getStringCellValue();
                            break;
                        case BLANK:
                            value="";
                            break;
                        case ERROR:
                            value=cell.getErrorCellValue()+"";
                            break;
                        }
					}
					switch (j) {
					case 1:
						//날짜 정리
						colMap.put("{date}", value);
						break;
					case 2:
						//주소
						colMap.put("{address}", value);
						break;
					case 3:
						//사업자
						colMap.put("{name}", value);
						if (value == null || value.equals("")) {
							breakFlag = true;
							break colLoop;
						}
						break;
					case 4:
						//사업자
						colMap.put("{number}", value);
						break;
					case 5:
						//연락처
						colMap.put("{phone}", value);
						break;
					case 6:
						//지원사업명
						colMap.put("{main_title}", value);
						break;
					case 7:
						//세부사업명
						colMap.put("{detail_title}", value);
						break;
					case 8:
						//규격
						colMap.put("{size}", value);
						break;
					case 9:
						//수량
						colMap.put("{n}", value);
						break;
					case 10:
						//단위
						colMap.put("{u}", value);
						break;
					case 11:
						//단가
						colMap.put("{price}", value);
						break;
					case 12:
						//공급금액
						colMap.put("{i_price}", value);
						break;
					case 13:
						//세액
						colMap.put("{t_price}", value);
						break;
					case 14:
						//사업금액
						colMap.put("{total}", value);
						break;
					case 15:
						//사업금액(한글)
						colMap.put("{korean_total}", value);
						break;
					case 16:
						//자부담
						colMap.put("{address}", value);
						break;
					case 17:
						//보조금
						colMap.put("{address}", value);
						break;
					}
				}
				if (!breakFlag) dataList.add(colMap);
			}
			
		}	
		
		return this.dataList;
//        int rowindex=0;
//        int columnindex=0;
//		
//		XSSFSheet sheet=workbook.getSheetAt(0);
		
	}
	
	
}
