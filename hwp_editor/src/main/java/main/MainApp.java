package main;

import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.bodytext.Section;
import kr.dogfoot.hwplib.object.bodytext.control.Control;
import kr.dogfoot.hwplib.object.bodytext.control.ControlTable;
import kr.dogfoot.hwplib.object.bodytext.control.ControlType;
import kr.dogfoot.hwplib.object.bodytext.control.table.Cell;
import kr.dogfoot.hwplib.object.bodytext.control.table.Row;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.HWPChar;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.HWPCharNormal;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.HWPCharType;
import kr.dogfoot.hwplib.reader.HWPReader;
import kr.dogfoot.hwplib.tool.objectfinder.ControlFilter;
import kr.dogfoot.hwplib.tool.objectfinder.ControlFinder;
import kr.dogfoot.hwplib.writer.HWPWriter;
import util.excel.ExcelParse;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MainApp {
	
	private static String path = "D:\\workspace\\java\\shopping_mall\\hwp_editor\\src\\main\\resources\\hwp\\";
	private static String input = "platform_father.hwp";
	private static String output = "edit-platform_father.hwp";
	
	private static Map<String, String> mainMap = new HashMap<String, String>();
	
	//셀의 텍스트 변경
    public static class MyControlFilter implements ControlFilter {
        public boolean isMatched(Control control, Paragraph paragrpah,
                                 Section section) {
            if (control.getType() == ControlType.Table) {
                ControlTable table = (ControlTable) control;
                for (Row raw : table.getRowList()) {
                	for (Cell cell : raw.getCellList()) {
                		try {
                    		for (Paragraph pg : cell.getParagraphList()) {
                    			for (String key : mainMap.keySet()) {
                    				changeNewCharList(pg, key, mainMap.get(key));
								}
							}
						} catch (Exception e) {
							// TODO: handle exception
						}
					}
				}
            }
            return false;
        }
    }

    private static void changeNewCharList(Paragraph paragraph, String targetWord, String replaceWord)
    	throws Exception{
    	if (!paragraph.getNormalString().contains(targetWord)) {
			return;
		}
    	int targetLength = targetWord.length();
    	int replaceLength = replaceWord.length();
    	ArrayList<HWPChar> charList = paragraph.getText().getCharList();
    	short targetFirstCode = (short)targetWord.codePointAt(0);
    	
    	//먼저 해당 문자열의 위치를 찾음.
    	//HWPCharType이 Normal이 아닌경우 오류가 발생하기에 직접 타입 체크후 변경해줄 필요가 있음.
    	boolean returnFlag = true;
    	mainLoop: for (int i = 0; i < charList.size(); i++) {
    		HWPChar ch = charList.get(i);
    		if (ch.getType() != HWPCharType.Normal) {
				continue mainLoop;
			}
    		if (ch.getCode() == targetFirstCode) {
    			for (int j = 1; j < targetLength; j++) {
    				HWPChar nextCh = charList.get(i + j);
					if (nextCh.getCode() != (short)targetWord.codePointAt(j)) {
						continue mainLoop;
					}
				}
    			//여기에 실제 기능 추가
    			for (int j = i; j < targetLength + i; j++) {
					charList.remove(i);
				}
				for (int j = replaceLength - 1; j >= 0; j--) {
					HWPCharNormal nch = new HWPCharNormal();
					nch.setCode((short)replaceWord.codePointAt(j));
					charList.add(i, nch);
				}
    			returnFlag = false;
    		}
    	}
    	if (returnFlag) return;
    	
    	if (paragraph.getNormalString().contains(targetWord)) {
    		changeNewCharList(paragraph, targetWord, replaceWord);
		}
    }
    
    private static void changeParagraphText(Paragraph paragraph) throws Exception {
//        ArrayList<HWPChar> newCharList = getNewCharList(paragraph.getText().getCharList());
		for (String key : mainMap.keySet()) {
			changeNewCharList(paragraph, key, mainMap.get(key));
		}
//        changeNewCharList(paragraph, newCharList);
//        removeLineSeg(paragraph);
//        removeCharShapeExceptFirstOne(paragraph);
    }

//    public static ArrayList<HWPChar> getNewCharList(ArrayList<HWPChar> oldList) throws UnsupportedEncodingException {
//        ArrayList<HWPChar> newList = new ArrayList<HWPChar>();
//        ArrayList<HWPChar> listForText = new ArrayList<HWPChar>();
//        for (HWPChar ch : oldList) {
//            if (ch.getType() == HWPCharType.Normal) {
//                listForText.add(ch);
//            } else {
//                if (listForText.size() > 0) {
//                    String text = toString(listForText);
//                    listForText.clear();
//                    String newText = changeText(text);
//
//                    newList.addAll(toHWPCharList(newText));
//                }
//                newList.add(ch);
//            }
//        }
//
//        if (listForText.size() > 0) {
//            String text = toString(listForText);
//            listForText.clear();
//            String newText = changeText(text);
//
//            newList.addAll(toHWPCharList(newText));
//        }
//        return newList;
//    }


    public static void main(String[] args) throws Exception {
    	
        FileInputStream inputFile = new FileInputStream(new File(path + "project_detail.xlsx"));

        ExcelParse excel = new ExcelParse(inputFile);
        excel.setRow(2, 13);
        excel.setCol(0, 21);
        excel.setDateCol(1);
        List<Map<String,String>> dataList = excel.parse(0);
                

    	
    	MyControlFilter myFilter = new MyControlFilter();
    	for (Map<String, String> map : dataList) {
    		mainMap = map;
    		HWPFile hwpFile = HWPReader.fromFile(path + input);
    		if (hwpFile == null) {
				break;
			}
    		//테이블 수정
            ArrayList<Control> result = ControlFinder.find(hwpFile, myFilter);
            Section s = hwpFile.getBodyText().getSectionList().get(0);
            int count = s.getParagraphCount();
            for (int index = 0; index < count; index++) {
                changeParagraphText(hwpFile.getBodyText().getSectionList().get(0).getParagraph(index));
            }
            HWPWriter.toFile(hwpFile, path + "20계약서(" + map.get("{name}") + ")농산물상자.hwp");
		}
            
        
    }
    
    
}
