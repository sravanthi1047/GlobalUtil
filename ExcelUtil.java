package GenericUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;

public class ExcelUtil {
	XSSFWorkbook wb=null;
	XSSFSheet active_sheet=null;
	File xlfile=null;
	FileInputStream fIn=null;
	FileOutputStream fOut=null;
	
	//getters & setters of file, fin and fout objects
	private File getXlfile() {
		return xlfile;
	}

	private void setXlfile(File xlfile) {
		this.xlfile = xlfile;
	}

	private FileInputStream getfIn() {
		return fIn;
	}

	private void setfIn(FileInputStream fIn) {
		this.fIn = fIn;
	}

	private FileOutputStream getfOut() {
		return fOut;
	}

	private void setfOut(FileOutputStream fOut) {
		this.fOut = fOut;
	}

	// Convenient functions for wb creation
	private void createwb_object(String file_path) throws FileNotFoundException, IOException {
		setXlfile(new File(file_path)); 
		setfIn(new FileInputStream(getXlfile()));
		wb = new XSSFWorkbook(getfIn());
	}
	
	private void createwbfile(String file_path) throws IOException {
		setXlfile(new File(file_path)); 
		setfIn(new FileInputStream(getXlfile()));
		wb = XSSFWorkbookFactory.createWorkbook(getfIn());
	}
	
	//existing wb instance 
	public ExcelUtil(String file_path) throws IOException {
		createwb_object(file_path);
	}

	//new wb creation
	public ExcelUtil(String file_path, boolean newfile) throws FileNotFoundException, IOException {
		if(newfile)
			createwbfile(file_path);
		else
			createwb_object(file_path);
	}
	
	//existing wb and sheet
	public ExcelUtil(String file_path, String sheet_name) throws FileNotFoundException, IOException {
		createwb_object(file_path);
		
		active_sheet=wb.getSheet(sheet_name);
	}
	
	public ExcelUtil(String file_path, int sheet_index) throws FileNotFoundException, IOException {
		createwb_object(file_path);
		
		active_sheet=wb.getSheetAt(sheet_index);
	}
	
	//new wb & new sheet would be created
	public ExcelUtil(String file_path, boolean newfile, String new_sheet_name) throws FileNotFoundException, IOException {
		if(newfile)
			createwbfile(file_path);
		else
			createwb_object(file_path);
		
		createSheet(new_sheet_name);
	}
	
	//new wb & list of new sheets to be created
	public ExcelUtil(String file_path, boolean newfile, String[] new_sheet_names) throws FileNotFoundException, IOException {
		if(newfile)
			createwbfile(file_path);
		else
			createwb_object(file_path);
		
		for(String sht : new_sheet_names) {
			createSheet(sht);
		}
		
	}

	//create new sheet
	private void createSheet(String sheet_name) {
		wb.createSheet(sheet_name);
	}
	
	//sets the active sheet
	public void SetActiveSheet(String sheetname) {
		active_sheet=wb.getSheet(sheetname);
	}
	
	//returns the sheet instance with name specified
	public XSSFSheet getSheet(String sheetname) {
		return wb.getSheet(sheetname);
	}
	
	//returns the requested cell value
	public String getCelldata(int rownumber, int columnnumber) {
		return active_sheet.getRow(rownumber).getCell(columnnumber).getStringCellValue();
	}
	
	//returns the requested cell value of the sheet other than the active one
	public String getCelldata(String sheetname, int rownumber, int columnnumber) {
		return getSheet(sheetname).getRow(rownumber).getCell(columnnumber).getStringCellValue();
	}
	

	//returns the requested range of cells' value of the sheet other than the active one
	private Object[][] getCellsdata(XSSFSheet sheet, int min_row, int max_row, int min_col, int max_col){
		Object[][] cellsdata = new Object[max_row-min_row][max_col-min_col];
		int r=0, c=0;
		for(int i=min_row; i<=max_row;) {
			XSSFRow _row= sheet.getRow(min_row);
			for(int j=min_col; j<=max_col;) {
				cellsdata[r][c]=_row.getCell(j).getRawValue();
				c++;
			}
			r++;
		}
		return cellsdata;
	}
	
	public Object[][] getAllCellsdata(){
		return getCellsdata(active_sheet,0, active_sheet.getLastRowNum(), 0, active_sheet.getRow(0).getLastCellNum());
	}
	
	public Object[][] getAllCellsdata(String sheet_name){
		return getCellsdata(this.getSheet(sheet_name),0, this.getSheet(sheet_name).getLastRowNum(), 0, this.getSheet(sheet_name).getRow(0).getLastCellNum());
	}
	
	//returns the requested range of cells' value of the active sheet
	public Object[][] getCellsdata(int min_row, int max_row, int min_col, int max_col){
		return getCellsdata(active_sheet,min_row, max_row, min_col, max_col);
	}
	
	//returns the requested cell value of the sheet other than the active one
	public Object[][] getCellsdata(String sheetname, int min_row, int max_row, int min_col, int max_col){
		return getCellsdata(this.getSheet(sheetname),min_row, max_row, min_col, max_col);
	}
	
	//writes into inputted sheet and inputted cell
	private void writeCelldata(XSSFSheet sht, int row, int column, String data) throws IOException {
		if(row>sht.getLastRowNum()) 
			sht.createRow(row).createCell(column);			
		else if(column>sht.getRow(row).getPhysicalNumberOfCells())
			sht.getRow(row).createCell(column);
			
		sht.getRow(row).getCell(column).setCellValue(data);
		this.SaveWb();
	}
	
	//writes the data into the cell of active sheet
	private void writeCelldata(int row, int column, String data) throws IOException {
		writeCelldata(active_sheet, row, column, data);		
	}

	//writes the data into the cell of input sheet
	public void writeCelldata(String sheetname, int row, int column, String data) throws IOException {
		writeCelldata(this.getSheet(sheetname), row, column, data);	
	}
	
	
	public void writeCellsdata(int min_row, int max_row, int min_col, int max_col, Object[][] data) throws IOException {
		_writeCellsdata(active_sheet, min_row, max_row, min_col, max_col, data);	
	}
	
	public void writeCellsdata(String sheetname, int min_row, int max_row, int min_col, int max_col, Object[][] data) throws IOException {
		_writeCellsdata(this.getSheet(sheetname), min_row, max_row, min_col, max_col, data);	
	}
	
	private void _writeCellsdata(XSSFSheet sht, int min_row, int max_row, int min_col, int max_col, Object[][] data) throws IOException{
		int r=0, c=0;
		for(int i=min_row; i<=max_row;) {
			XSSFRow _row= sht.getRow(min_row);
			for(int j=min_col; j<=max_col;) {
				_row.getCell(j).setCellValue(data[r][c].toString());
				c++;
			}
			r++;
		}
		this.SaveWb();
	}
	
	public void SaveWb() throws IOException {
		fIn.close();
		setfOut(new FileOutputStream(getXlfile()));
		wb.write(getfOut());		
	}
	//public void writeCellsdata(String sheetname, int min_row, int max_row, int min_col, int max_col, String[][] data){
	public void Close() throws IOException {
		getfOut().close();
		wb.close();
		
	}
	
//	public static void main(String[] args) throws FileNotFoundException, IOException {
//		// TODO Auto-generated method stub
//		//ExcelUtil xl = new ExcelUtil("C:\\Selenium\\Data.xlsx","ss");
//		
//		File f= new File("C:\\Selenium\\Data.xlsx");
//		FileInputStream f_in = new FileInputStream(f);
//		Workbook wb = new XSSFWorkbook(f_in);
//		System.out.println(wb.getSheetName(1));
//		
//		wb.close();
////		System.out.println(xl.getAllCellsdata());
////		
////		xl.Close();
//	}
}
