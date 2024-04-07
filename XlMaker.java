package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlMaker {

	static XlMaker xlMaker=new XlMaker(); 
	String buffer="";
	String[]parms;
	FileInputStream fisWorkBook=null;
	Workbook oWorkBook;

	int iTemplate=0;

	public static void main(String[] args) {
		
	    long start = System.currentTimeMillis();

	    xlMaker.interpret("C:/Users/Kingdel/eclipse-workspace/temp/in/xlMaker.xCmd");
        long end = System.currentTimeMillis();
	    
        float sec = (end - start) / 1000F; 
	    System.out.println(sec + " seconds");

	}

	private void interpret(String sFile) {
		File myFile = new File(sFile);
		Scanner scanner=null;
		try {
			scanner = new Scanner(myFile);
		} catch (FileNotFoundException e) {
			System.out.println("read failure:"+sFile);
			return;
		}
		while (scanner.hasNextLine()) {
			String srcDta = scanner.nextLine();
			loadBuffer(srcDta);
		}
		scanner.close();
	}

	private void loadBuffer(String srcDta) {
		srcDta=srcDta.trim();

		if (srcDta.isEmpty() || srcDta.startsWith("//")) return;

		buffer+=srcDta;
		cmdExec();
	}

	private void cmdExec() {
		int p1, p2;
		String command="", parmx="";
		if (!buffer.endsWith(";")) return;

		p1=buffer.indexOf("(");
		command=buffer.substring(0, p1);

		p2=buffer.lastIndexOf(")");

		parmx=buffer.substring(p1+1,p2);
		parms=parmx.split(",");

		if (command.equals("OVRXLT")) ovrXlt(parms);
		else if (command.equals("NEWXL")) newXl(parms);
		else if (command.equals("ADDSHEET")) addSheet(parms);
		else if (command.equals("SAVXL")) savXl(parms);
		else if (command.equals("RLSXLT")) rlsXlt(parms);
		else System.out.println("unknown command:"+command);

		buffer="";
	}

	private void ovrXlt(String[] parms) {
		try {
			fisWorkBook =new FileInputStream(new File(parms[0]));

		} catch (Throwable e) {
			System.out.println("failure on OVRXLT:"+parms[0]);
		}
	}

	private void newXl(String[] parms2) {
		try {
			oWorkBook=new XSSFWorkbook(fisWorkBook);

			iTemplate=oWorkBook.getNumberOfSheets();


		} catch (IOException e) {
			System.out.println("failure on newXL:");
		}
	}

	private void addSheet(String[] parms2) {

		String sTemplate=parms[0];
		String sNew=parms[1];
		String sFileName=((parms.length==3)?parms[2]:parms[1]);

		Sheet shTemplate = oWorkBook.getSheet(sTemplate);

		int iTemplate = oWorkBook.getSheetIndex(shTemplate);
		Sheet shNew = oWorkBook.cloneSheet(iTemplate);
		oWorkBook.setSheetName(oWorkBook.getSheetIndex(shNew), sNew);

		for (Name name:oWorkBook.getAllNames()) {

			if (name.getSheetName().equals(sTemplate) ) {
				fillSheet(name,shTemplate,shNew, sFileName);
			}
		}
	}

	private void fillSheet(Name name, Sheet shTemplate,Sheet shNew, String sFileName) {

		if (name.getNameName().endsWith("FilterDatabase")) return;
		
		AreaReference aRef = new AreaReference(name.getRefersToFormula(), null);
		CellReference cellRef = aRef.getFirstCell();
		int offset= cellRef.getCol(); 
		int y= cellRef.getRow(); 

		Row tRow = shTemplate.getRow(y);

		String sExtension = name.getNameName();
		String sFilePath= "C:/Users/Kingdel/eclipse-workspace/temp/in/"+sFileName+"."+sExtension;

		File myFile = new File(sFilePath);
		Scanner myReader=null;
		try {
			myReader = new Scanner(myFile);
		} catch (FileNotFoundException e) {
			System.out.println("read failure:"+sFilePath);
			return;
		}
		while (myReader.hasNextLine()) {
			String srcDta = myReader.nextLine();
			String[]strings = srcDta.split(";");  

			fillRow(tRow,offset,strings, shNew, y);
			y++;
		}
		myReader.close();
	}

	private void fillRow(Row tRow, int offset, String[] strings, Sheet shNew,int y) {

		Row oRow = shNew.createRow(y);

		for (int s=0;s<strings.length;s++) {

			Cell tCell= tRow.getCell(s+offset);
	
			Cell oCell = oRow.createCell(s+offset);

			String string = strings[s];

			if      (CellType.NUMERIC==tCell.getCellType())	oCell.setCellValue(Double.valueOf(string));
			else if (CellType.STRING==tCell.getCellType())	oCell.setCellValue(string);
			else if (CellType.BOOLEAN==tCell.getCellType())	oCell.setCellValue(Boolean.valueOf(string));
			else                                            oCell.setCellValue(string);

			oCell.setCellStyle(tCell.getCellStyle());
			
		}

	}

	private void savXl(String[] parms) {
		String pathName = parms[0];

		try {
			for (int t= iTemplate-1;t>=0;t--) oWorkBook.removeSheetAt(t);
			oWorkBook.write(new FileOutputStream(pathName));
			System.out.println("fin");
		} catch (IOException e) {
			System.out.println("failure on savXL: "+pathName);
		}
	}

	private void rlsXlt(String[] parms2) {
		fisWorkBook=null;
	}

}
