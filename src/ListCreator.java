import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.commons.collections4.iterators.IteratorChain;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.helpers.XSSFXmlColumnPr;

public class ListCreator {
	/*
	 * output.xlsx and input.xlsx should in same directory of running jar
	 */
	public class RowData {
		public int sentenceId;
		public int rowId;
		public String sentence;
		public RowData(int sentenceId, int rowId, String sentence) {
			this.sentenceId = sentenceId;
			this.rowId = rowId;
			this.sentence = sentence;
		}

	}

	public void LoadAndCreatLists(String filePath){
		File f=new File("output.xlsx");
		if(!f.exists()){
			try {
				ArrayList<RowData> mydata=new ArrayList<>();
				FileInputStream file=new FileInputStream(new File("input.xlsx"));
				XSSFWorkbook input=new XSSFWorkbook(file);
				XSSFSheet datasheet=input.getSheetAt(0);
				Iterator<Row> rowIterator = datasheet.iterator();
				rowIterator.next();
				int sentence=1;
				int rowid=-1;
				while(rowIterator.hasNext()){
					Row row=rowIterator.next();
					Iterator<Cell> cellIterator=row.cellIterator();
					while(cellIterator.hasNext()){
						Cell cell = cellIterator.next();
						//Check the cell type and format accordingly
						switch (cell.getCellType())
						{
						case Cell.CELL_TYPE_NUMERIC:
							rowid=(int) cell.getNumericCellValue();
							break;
						case Cell.CELL_TYPE_STRING:
							String str=cell.getStringCellValue();
							String[] tok=str.split("[.?]");
							for(int i=0;i<tok.length;i++){
								if(!(tok[i].trim().equals("") || (tok[i].length()<5))){
									//System.out.println(sentence+"  "+rowid+": "+tok[i]);
									mydata.add(new RowData(sentence, rowid, tok[i].trim()));
									sentence++;
								}
							}
							break;
						}
					}
				}
				file.close();
				createLists(mydata);
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

	}
	private void createLists(ArrayList<RowData> data){
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Task 1");
		XSSFSheet sheet2=workbook.createSheet("Task 2");
		XSSFSheet sheet3=workbook.createSheet("Task 3");
		//Task_1 Header
		XSSFRow row=sheet.createRow(0);
		XSSFCell c1=row.createCell(0);
		c1.getCellStyle().setWrapText(true);
		c1.setCellValue("sentence_id");
		XSSFCell c2=row.createCell(1);
		c2.setCellValue((String) "row id");
		XSSFCell c3=row.createCell(2);
		c3.setCellValue((String)"Sentence");
		//Task_2 Header
		Row row2=sheet2.createRow(0);
		row2.createCell(0).setCellValue("ID");
		row2.createCell(1).setCellValue("Sentence_ID");
		row2.createCell(2).setCellValue("Sentence");
		row2.createCell(3).setCellValue("CheckBox 1");
		row2.createCell(4).setCellValue("CheckBox 2");
		row2.createCell(5).setCellValue("CheckBox 3");
		row2.createCell(6).setCellValue("CheckBox 4");
		row2.createCell(7).setCellValue("CheckBox 5");
		row2.createCell(8).setCellValue("CheckBox 6");
		row2.createCell(9).setCellValue("CheckBox 7");
		row2.createCell(10).setCellValue("Skip Count");
		row2.createCell(11).setCellValue("Submit Count");
		int rownum = 1;
		for(RowData r:data){

			row=sheet.createRow(rownum);
			row2=sheet2.createRow(rownum);
			System.out.println(r.sentenceId+"  "+r.rowId+"  "+r.sentence);
			c1=row.createCell(0);
			c1.setCellValue((Integer) r.sentenceId);
			c2=row.createCell(1);
			c2.setCellValue((Integer) r.rowId);
			c3=row.createCell(2);
			c3.setCellValue((String)r.sentence);
			row2.createCell(0).setCellValue(r.sentenceId);
			row2.createCell(1).setCellValue(r.sentenceId);
			row2.createCell(2).setCellValue(r.sentence);
			row2.createCell(3).setCellValue(0);
			row2.createCell(4).setCellValue(0);
			row2.createCell(5).setCellValue(0);
			row2.createCell(6).setCellValue(0);
			row2.createCell(7).setCellValue(0);
			row2.createCell(8).setCellValue(0);
			row2.createCell(9).setCellValue(0);
			row2.createCell(10).setCellValue(0);
			row2.createCell(11).setCellValue(0);
			rownum++;
		}
		try
		{
			//Write the workbook in file system
			FileOutputStream out = new FileOutputStream(new File("Output.xlsx"));

			workbook.write(out);
			out.close();
			System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
		}
		catch (Exception e)
		{

		}
	}
}
