package whatsapp;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Takes care of operations such as extracting existing contacts from masterExcel ,removing duplicates and updating them to the master excel
 * @author vishal sundararajan
 *
 */
public class MasterExcel {

	
	private HashSet<String> kktContacts=new HashSet<String>();  
	private HashSet<String> yuvaVikasContacts=new HashSet<String>();  
	/**
	 * Gets the entries from the master excel and remove duplicates in the contacts HashSet by comparing with the master excel
	 */
	void processMasterExcel(){
		
		try{
			processKKtContacts();
			processYuvaVikas();
			createExcel();
			createYuvaVikasExcel();
		
		}
		catch(Exception e){
			Global.exception.error("Error occurred in processMasterExcel :", e);
		}
		
	}
	void processYuvaVikas(){
		try{
			System.out.println("Processing kkt contacts");
			Pattern pattern = Pattern.compile(".*[a-zA-Z]+.*");
			
			for(String raw:Global.yuvaVikas){
				Matcher matcher = pattern.matcher(raw);
				if((!matcher.matches()) && (!raw.contains("_"))){
					if(checkRegionalLanguages(raw)){
						String rawParsed=raw.replaceAll("\\)", "").replaceAll("\\(", "").replaceAll("-", "").replaceAll(" ", "").trim();
						yuvaVikasContacts.add(rawParsed);
					}
				}
			}
			Global.excelLog.info("contents after parsing processYuvaVikas : "+yuvaVikasContacts);
		}
		catch(Exception e){
			Global.exception.error("Error occurred in processYuvaVikas :", e);
		}
	}
	private void createYuvaVikasExcel(){
		try{
			if(yuvaVikasContacts!=null && !yuvaVikasContacts.isEmpty()){
				Date date = new Date() ;
				SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy hh-mm-ss") ;
				XSSFWorkbook workbook = new XSSFWorkbook();
				XSSFSheet sheet = workbook.createSheet("Yuva sheet");
				int i=0;
				for(String contact:yuvaVikasContacts){
					Row row = sheet.createRow(i);
					Cell cell = row.createCell(0);
					cell.setCellValue(contact);
					cell.setCellType(Cell.CELL_TYPE_STRING);
					i++;
				}
				FileOutputStream out = 
						new FileOutputStream(new File(Global.homePath+File.separator+"Contacts"+File.separator+"YuvaVikas "+dateFormat.format(date)+".xlsx"));
				workbook.write(out);
				workbook.close();
				out.close();
				System.out.println("Yuva Vikas Excel written successfully..");
			}else{
				System.out.println("No yuva vikas contact to write to excel");
			}
		}
		catch(Exception e){
			
		}
	}
	private void createExcel(){
		try{
		
			Date date = new Date() ;
			SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy hh-mm-ss") ;
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet("Kkt sheet");
			int i=0;
			for(String contact:kktContacts){
				Row row = sheet.createRow(i);
				Cell cell = row.createCell(0);
				cell.setCellValue(contact);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				i++;
			}
			FileOutputStream out = 
					new FileOutputStream(new File(Global.homePath+File.separator+"Contacts"+File.separator+"KKT "+dateFormat.format(date)+".xlsx"));
			workbook.write(out);
			workbook.close();
			out.close();
			System.out.println("KKT Excel written successfully..");
			
			
		}
		catch(Exception e){
			Global.exception.error("Exception occurred while creating excel ", e);
		}
	}

	@SuppressWarnings("unused")
	private static boolean isRowEmpty(Row row) {
	    for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
	        Cell cell = row.getCell(c);
	        if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK)
	            return false;
	    }
	    return true;
	}
	private void processKKtContacts(){
		try{
			System.out.println("Processing kkt contacts");
			Pattern pattern = Pattern.compile(".*[a-zA-Z]+.*");
			
			for(String raw:Global.kkt){
				Matcher matcher = pattern.matcher(raw);
				if((!matcher.matches()) && (!raw.contains("_"))){
					if(checkRegionalLanguages(raw)){
						String rawParsed=raw.replaceAll("\\)", "").replaceAll("\\(", "").replaceAll("-", "").replaceAll(" ", "").trim();
						kktContacts.add(rawParsed);
					}
				}
			}
			Global.excelLog.info("contents after parsing raw contacts : "+kktContacts);
		}
		catch(Exception e){
			Global.exception.error("Exception occurred while processing raw contacts ", e);
		}
	}
	/**
	 * This method will remove groups that contain regional language characters in it , will return true if the element does not contain any regional characters
	 * @param element
	 * @return
	 */
	private boolean checkRegionalLanguages(String element){
		try{
			for (char c: element.toCharArray()) {
			     if ((Character.UnicodeBlock.of(c) == Character.UnicodeBlock.DEVANAGARI) || (Character.UnicodeBlock.of(c) == Character.UnicodeBlock.TAMIL) 
			    		 || (Character.UnicodeBlock.of(c) == Character.UnicodeBlock.TELUGU)) {
			    	 return false;
			     }
			 }
		}
		catch(Exception e){
			Global.exception.error("Exception in checkRegionalLanguages", e);
		}
		return true;
	}

}
