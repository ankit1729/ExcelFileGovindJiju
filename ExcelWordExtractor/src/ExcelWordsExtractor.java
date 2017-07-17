import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.FileAlreadyExistsException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileSystemView;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWordsExtractor {

	public static void main(String[] args) {
		
		JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());

		int returnValue = jfc.showOpenDialog(null);
		// int returnValue = jfc.showSaveDialog(null);
		File selectedFile=null;
		if (returnValue == JFileChooser.APPROVE_OPTION) {
			selectedFile = jfc.getSelectedFile();
			selectedFile.getAbsoluteFile();
			//System.out.println(selectedFile.getAbsolutePath().replaceAll("\\\\+", File.separator));
		}
		
		try {
			
			FileInputStream inputStream=new FileInputStream(selectedFile.getAbsoluteFile());
			//FileInputStream inputStream=new FileInputStream(new File("C://Users//ankchopr.ORADEV//Desktop//Test Sample 2.xlsx"));
			XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
			Sheet sheet=workbook.getSheetAt(0);
			Iterator<Row> rowIterator=sheet.iterator();
			List<String> queryList=new ArrayList<>();
			boolean queryListPopulatedFlag=false;
			while(rowIterator.hasNext()){
				Row row=rowIterator.next();
				/*if(row.getRowNum()==30)
					System.out.println(row.getRowNum());*/
				
				//System.out.println("nmbr of cols : "+row.getLastCellNum());
				if(isRowEmpty(row)){
					continue;
				}
				else if(!queryListPopulatedFlag){
					int cellCount=0;
					Iterator<Cell> cellIterator=row.cellIterator();
					while(cellIterator.hasNext()){
						Cell cell=cellIterator.next();
						if(cellCount==0){
							cellCount++;
							continue;
						}
						else{
							queryList.add(cell.getStringCellValue());
							cellCount++;
						}
					}
					queryListPopulatedFlag=true;
					System.out.println("query list is : "+queryList);
				}
				else{
					//int cellCount=0;
					String input="";
					StringBuffer freqCellval=new StringBuffer();
					int cellCount=0;
					for(;cellCount<=queryList.size();cellCount++)/*while(cellIterator.hasNext())*/{
						int queryCount=0;
						StringBuffer buffer=new StringBuffer("");
						//Cell cell=cellIterator.next();
						Cell cell=row.getCell(cellCount, MissingCellPolicy.CREATE_NULL_AS_BLANK);
						if(cellCount==0){
							input=cell.getStringCellValue();
							//cellCount++;
						}
						else{
							int queryIndex=input.toLowerCase().indexOf(queryList.get(cellCount-1).toLowerCase());
							boolean queryFoundFlag=false;
							while(queryIndex!=-1){
								queryFoundFlag=true;
								queryCount++;
								if(queryIndex+20<=input.length()-1 && queryIndex-20>=0){
									buffer.append(input.substring(queryIndex-20, queryIndex+20)+"\n");
								}
								else if(queryIndex+10<=input.length()-1 && queryIndex-10>=0){
									buffer.append(input.substring(queryIndex-10, queryIndex+10)+"\n");
								}
								else if (queryIndex+5<=input.length()-1 && queryIndex-5>=0){
									buffer.append(input.substring(queryIndex-5, queryIndex+5)+"\n");
								}
								else{
									buffer.append(queryList.get(cellCount-1)+"\n");
								}
								
								
								queryIndex=input.toLowerCase().indexOf(queryList.get(cellCount-1).toLowerCase(), queryIndex+1);
							}
							if(queryFoundFlag){
								freqCellval.append(queryList.get(cellCount-1)+ " : "+queryCount+"\n");
							}
							//cellCount++;
						}
						
						if(!("".equals(buffer.toString())) && cellCount!=0){
							//System.out.println("query : "+queryList.get(cellCount-2)+"buffer : "+buffer.toString());
							cell.setCellValue(buffer.toString());
							CellStyle cellStyle=workbook.createCellStyle();
							cellStyle.setWrapText(true);
							cell.setCellStyle(cellStyle);
							sheet.autoSizeColumn(cell.getColumnIndex());
						}
					}
						Cell freqCell=row.createCell(cellCount);
						freqCell.setCellValue(freqCellval.toString());
						CellStyle cellStyle=workbook.createCellStyle();
						cellStyle.setWrapText(true);
						freqCell.setCellStyle(cellStyle);
						sheet.autoSizeColumn(freqCell.getColumnIndex());
				}
			}
			inputStream.close();
			FileOutputStream outputStream=new FileOutputStream(selectedFile.getAbsoluteFile());
			workbook.write(outputStream);
			
			workbook.close();
			outputStream.close();
			
			JOptionPane optionPane = new JOptionPane("Done Please open file now",JOptionPane.PLAIN_MESSAGE);
			JDialog dialog = optionPane.createDialog("Information!!");
			dialog.setAlwaysOnTop(true); // to show top of all other application
			dialog.setVisible(true); // to visible the dialog
			System.exit(0);

		} catch (Exception e) {
			
			JOptionPane optionPane = new JOptionPane("Some unexpected error occurred.",JOptionPane.ERROR_MESSAGE);
			JDialog dialog = optionPane.createDialog("Error!!");
			dialog.setAlwaysOnTop(true); // to show top of all other application
			dialog.setVisible(true);
		}
	
	}
	
	private static void drawGraph(){
		
	}
	
	public static boolean isRowEmpty(Row row) {
	    for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
	        Cell cell = row.getCell(c);
	        if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK)
	            return false;
	    }
	    return true;
	}

}
