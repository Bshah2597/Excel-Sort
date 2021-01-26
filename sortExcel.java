import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	public class sortExcel {

		private static String[] columns = { "Sid (Not In CyberArk)", "Account", "Platform", " SR #", "Status"}; // Setting up column names for CyberArk Reconciliation Excel Report
		
		private static String[] columnsInfra = {"Sid (Not In Infrastructure)", "Account Name", "Safe" }; // Setting column names for Infrastructure Reconciliation Report 

		public static List<Map<String, String>> processFiles(String[] args) { // List of individual hashmaps 
			final String ENABLED = "Enabled"; // String that later helps filters for only enabled accounts
			List<Map<String, String>> reportslist = null;

			try {
				// Windows Data
				Map<String, String> userReportSids = new HashMap<>();  // Creating a map with key value as sids linking to account name 
				Map<String, String> userReportPlatform = new HashMap<>(); // Creating map with key value as sids linking to platform 

				String file1 = null;
				String file2 = null;
				String file3 = null;

				file1 = args[0]; // Each file selected will be saved in these string values of "file1" "file2" exc. 
				file2 = args[1];
				file3 = args[2];

				InputStream SourceDataFile;
				XSSFWorkbook wb1 = null;
				try {
					SourceDataFile = new FileInputStream(file1); // Input Stream to take in First Source Data File to read and parse data 
					wb1 = new XSSFWorkbook(SourceDataFile);
				} catch (FileNotFoundException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				// Source Data
				Row rowWD;
				XSSFSheet sheet1 = wb1.getSheetAt(0);
				int colIndexWD = 1; // change this for corresponding ID value for source data. Index 0 = Column A in excel. Index 1 = Column B in excel. Index always starts at 0 for first column
				int colSidNameIndex = 0; // The column in which the account names are located
				int colStatusIndex = 2; // Status of account whether it is enabled or disabled
				int colPlatformIndex = 3; //Column for which the platform is specified in the User List 
				Cell cellWD = null;
				Cell cellSidName = null;
				Cell cellStatus = null;
				Cell cellPlatfrom = null;
				for (int rowIndexWD = 2; rowIndexWD <= sheet1.getLastRowNum(); rowIndexWD++) { 
					// change rowIndexWD for the location where you want to start counting the row. Index 0 = Row 1 in excel.
					rowWD = sheet1.getRow(rowIndexWD);
					if (rowWD != null) {
						cellWD = rowWD.getCell(colIndexWD);
						cellSidName = rowWD.getCell(colSidNameIndex);
						cellStatus = rowWD.getCell(colStatusIndex);
						cellPlatfrom = rowWD.getCell(colPlatformIndex);
						if (null != cellWD && null != cellStatus && ENABLED.equalsIgnoreCase(cellStatus.getStringCellValue())) {
							userReportSids.put(cellWD.getStringCellValue(), (null != cellSidName) ? cellSidName.getStringCellValue() : "");
							userReportPlatform.put(cellWD.getStringCellValue(), (null != cellPlatfrom) ? cellPlatfrom.getStringCellValue() : "");
						}
					}
				}
				wb1.close();

				// CyberArk Data
				ArrayList<String> CyberarkData = new ArrayList<String>();
				InputStream CyberarkDatafile = new FileInputStream(file2);
				XSSFWorkbook wb2 = new XSSFWorkbook(CyberarkDatafile);
				Row rowCA;
				int colIndexCA = 16;
				XSSFSheet sheet2 = wb2.getSheetAt(0);
				Cell cellCA;
				for (int rowIndexCA = 1; rowIndexCA < sheet2.getLastRowNum(); rowIndexCA++) {
					rowCA = sheet2.getRow(rowIndexCA);
					if (rowCA != null) {
						cellCA = rowCA.getCell(colIndexCA);
						if (null != cellCA) {
							CyberarkData.add(cellCA.getStringCellValue());
						}
					}
				}
				wb2.close();

				// BaselineReport accounts
				ArrayList<String> BaselineData = new ArrayList<String>();
				InputStream BaselinePersonalAccounts = new FileInputStream(file3);
				XSSFWorkbook wb3 = new XSSFWorkbook(BaselinePersonalAccounts);
				Row rowBA;
				XSSFSheet sheet3 = wb3.getSheetAt(0);
				int colIndexBA = 1;
				for (int rowIndexBA = 1; rowIndexBA <= sheet3.getLastRowNum(); rowIndexBA++) {
					rowBA = sheet3.getRow(rowIndexBA);
					if (rowBA != null) {
						Cell cellBA = rowBA.getCell(colIndexBA);
						if (null != cellBA) {
							BaselineData.add(cellBA.getStringCellValue());
						}
					} else {
						break;
					}
				}
				try {
					wb3.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				// System.out.println(BaselineData.size());

				// First Compare Database Report with Cyberark
				// Add Database data i.e Windows data which also contains the Cyberark Data
				// Remove Cyberark Accounts in the Windows data (basically removing the cyberark
				// accounts which were in CA file)
				// Now we have Windows report without the accounts in Cyberark, so all the
				// accounts shown will not be in cyberark
				// Get the report of ONLY the Personal accounts from Baseline Report - Joe
				// Now remove those Baseline Accounts from new data set
				// Leaves you with NPA accounts that are not in Cyberark and are also not in
				// Baseline Report

				// removing the MainData
				userReportSids.keySet().removeAll(CyberarkData);
				userReportPlatform.keySet().removeAll(CyberarkData);

				// removing the BaselineData
				userReportSids.keySet().removeAll(BaselineData);
				userReportPlatform.keySet().removeAll(BaselineData);
				
				reportslist = new ArrayList<>();
				reportslist.add(userReportSids);
				reportslist.add(userReportPlatform);
				
				// displaying the final list
				for (Map.Entry<String,String> entry  : userReportSids.entrySet()) {
					System.out.println(entry.getValue());
				}
				// displaying the final list size
				System.out.println(userReportSids.size());
			} catch (FileNotFoundException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			return reportslist;
		}

		public static List<Map<String, String>> processInfraFiles(String[] args) {
			List<Map<String, String>> reportslist = null;
//			final String ENABLED = "Enabled";
			try {
				// Windows Data
				Map<String, String> CyberarkMap = new HashMap<>(); // Links Sid key back with Account name 
				Map<String, String> Safe = new HashMap<>();  // Links Sid to associated Safe value 
				ArrayList<String> SourceData = new ArrayList<>(); 
				String file1 = null;
				String file2 = null;
				String file3 = null; 
				
				file1 = args[0];
				file2 = args[1];
				file3 = args[2]; 
				
				
				InputStream SourceDataFile;
				XSSFWorkbook wb1 = null;
				try {
					SourceDataFile = new FileInputStream(file1);
					wb1 = new XSSFWorkbook(SourceDataFile);
				} catch (FileNotFoundException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				// Source Data
				Row rowWD;
				XSSFSheet sheet1 = wb1.getSheetAt(0); // Determines which sheet to obtain to parse the data (set this index to whatever sheet your data is on) Index 0 = Sheet 1 
				int colIndexWD = 1; // change this for corresponding ID value for source data. Index 0 = Column A in excel. Index 1 = Column B in excel. Index always starts at 0 for first column
				Cell cellWD = null;
				for (int rowIndexWD = 2; rowIndexWD <= sheet1.getLastRowNum(); rowIndexWD++) { // change rowIndexWD for the location where you want to start counting the row. Index 0 = Row 1 in excel.
					rowWD = sheet1.getRow(rowIndexWD);
					if (rowWD != null) {
						cellWD = rowWD.getCell(colIndexWD);
						if (null != cellWD){
							SourceData.add(cellWD.getStringCellValue()); 
						}
					}
				}
				wb1.close();

				// CyberArk Data
//				ArrayList<String> CyberarkData = new ArrayList<String>();
				InputStream CyberarkDatafile = new FileInputStream(file2);
				XSSFWorkbook wb2 = new XSSFWorkbook(CyberarkDatafile); // initializing a new workbook and parsing data from CyberArk file that user entered in UI 
				Row rowCA;
				int colIndexCA = 16; // This column is for Sid Values 
				int colSidNameIndex = 4; // This column is for Name associated with each Sid
				int colSafeIndex = 0; // Safe associated with each Sid 
				XSSFSheet sheet2 = wb2.getSheetAt(0); // sheet index, index 0 = sheet 1 of excel workbook 
				Cell cellCA;
				Cell cellSidName = null; 
				Cell cellSafeName = null; 
				for (int rowIndexCA = 1; rowIndexCA < sheet2.getLastRowNum(); rowIndexCA++) {
					rowCA = sheet2.getRow(rowIndexCA);
					if (rowCA != null) {
						cellCA = rowCA.getCell(colIndexCA); // Puts SID values in initialized Cell
						cellSidName = rowCA.getCell(colSidNameIndex);  //Puts Name linked to sid in initialized cell 
						cellSafeName = rowCA.getCell(colSafeIndex); // Puts Safe linked to sid 
						if (cellCA != null) {
							CyberarkMap.put(cellCA.getStringCellValue(), (null != cellSidName) ? cellSidName.getStringCellValue() : ""); //places key value first (SID), then links that to Account name IF the account name isnt blank, if it is blank it just links blank account name with SID
							Safe.put(cellCA.getStringCellValue(), (null != cellSafeName) ? cellSafeName.getStringCellValue() : ""); // Does the same except with Safe Name 
						} 
					}
				}
				
				wb2.close();
				
				// BaselineReport accounts
				ArrayList<String> BaselineData = new ArrayList<String>();
				InputStream BaselinePersonalAccounts = new FileInputStream(file3);
				XSSFWorkbook wb3 = new XSSFWorkbook(BaselinePersonalAccounts);
				Row rowBA;
				XSSFSheet sheet3 = wb3.getSheetAt(0);
				int colIndexBA = 1;
				for (int rowIndexBA = 1; rowIndexBA <= sheet3.getLastRowNum(); rowIndexBA++) {
					rowBA = sheet3.getRow(rowIndexBA);
					if (rowBA != null) {
						Cell cellBA = rowBA.getCell(colIndexBA);
						if (null != cellBA) {
							BaselineData.add(cellBA.getStringCellValue());
						}
					} else {
						break;
					}
				}
				try {
					wb3.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				
				CyberarkMap.keySet().removeAll(SourceData); 
				Safe.keySet().removeAll(SourceData); 
				
				CyberarkMap.keySet().removeAll(BaselineData); 
				Safe.keySet().removeAll(BaselineData); 
//				newSet = new HashSet<>(CyberarkMap.keySet()); 

//				for (String str : newSet) {
//					System.out.println(str);
//				}
				reportslist = new ArrayList<>();
				reportslist.add(CyberarkMap); 
				reportslist.add(Safe); 
				
			} catch (FileNotFoundException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} 
			return reportslist;
			
		}

		public static boolean generateInfraFile(List<Map<String, String>> Sid, String fileName )   {
			
			Map<String, String> Safe = Sid.get(1);  
			
			//Create Workbook
			Workbook workbookInfra = new XSSFWorkbook(); 
			
			//Create Sheet
			Sheet sheet = workbookInfra.createSheet("Accounts not in Infrastructure Reports"); 
			
			// Create a Font for styling header cells
					Font headerFont = workbookInfra.createFont();
					headerFont.setBold(true);
					headerFont.setFontHeightInPoints((short) 14);
					headerFont.setColor(IndexedColors.BLUE.getIndex());

					// Create a CellStyle with the font
					CellStyle headerCellStyle = workbookInfra.createCellStyle();
					headerCellStyle.setFont(headerFont);

					// Create a Row
					Row headerRow = sheet.createRow(0);

					// Creating cells
					Cell cell1 = headerRow.createCell(0);
					cell1.setCellValue(columnsInfra[0]);
					cell1.setCellStyle(headerCellStyle);
					
					Cell cell2 = headerRow.createCell(1); 
					cell2.setCellValue(columnsInfra[1]);
					cell2.setCellStyle(headerCellStyle);
					
					Cell cell3 = headerRow.createCell(2); 
					cell3.setCellValue(columnsInfra[2]);
					cell3.setCellStyle(headerCellStyle);
					
					// Create Other rows and cells with data
					int rowNum = 1;
					for (Map.Entry<String, String> entry: Sid.get(0).entrySet()) {
						Row row = sheet.createRow(rowNum++);
						row.createCell(0).setCellValue(entry.getKey());
						row.createCell(1).setCellValue(entry.getValue());
						row.createCell(2).setCellValue(Safe.get(entry.getKey()));
					}

					sheet.autoSizeColumn(0);
					
					// Write the output to a file
					FileOutputStream fileOut;
					try {
						fileOut = new FileOutputStream(fileName);
						workbookInfra.write(fileOut);
						fileOut.close();
						workbookInfra.close();
						return true;
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
			
			return false; 
		}
		
		public static boolean generateFile(List<Map<String, String>> filteredDomains, String fileName) {
			//List<Map<String, String>>
			//List<String> AccNames = new ArrayList<>(filteredDomains);
			Map<String, String> platforms = filteredDomains.get(1);

			// Create a Workbook
			Workbook workbook = new XSSFWorkbook();

			// Create a Sheet
			Sheet sheet = workbook.createSheet("Accounts Not in CyberArk");

			// Create a Font for styling header cells
			Font headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerFont.setFontHeightInPoints((short) 14);
			headerFont.setColor(IndexedColors.BLUE.getIndex());

			// Create a CellStyle with the font
			CellStyle headerCellStyle = workbook.createCellStyle();
			headerCellStyle.setFont(headerFont);

			// Create a Row
			Row headerRow = sheet.createRow(0);

			// Creating cells
			Cell cell1 = headerRow.createCell(0);
			cell1.setCellValue(columns[0]);
			cell1.setCellStyle(headerCellStyle);
			
			Cell cell2 = headerRow.createCell(1);
			cell2.setCellValue(columns[1]);
			cell2.setCellStyle(headerCellStyle);
			
			Cell cell3 = headerRow.createCell(2);
			cell3.setCellValue(columns[2]);
			cell3.setCellStyle(headerCellStyle);
			
			Cell cell4 = headerRow.createCell(3);
			cell4.setCellValue(columns[3]);
			cell4.setCellStyle(headerCellStyle);
			
			Cell cell5 = headerRow.createCell(4);
			cell5.setCellValue(columns[4]);
			cell5.setCellStyle(headerCellStyle);

			// Create Other rows and cells with data
			int rowNum = 1;
			for (Map.Entry<String,String> entry  : filteredDomains.get(0).entrySet()) {
				Row row = sheet.createRow(rowNum++);
				row.createCell(0).setCellValue(entry.getKey());
				row.createCell(1).setCellValue(entry.getValue());
				row.createCell(2).setCellValue(platforms.get(entry.getKey()));
			}

			sheet.autoSizeColumn(0);

			// Write the output to a file
			FileOutputStream fileOut;
			try {
				fileOut = new FileOutputStream(fileName);
				workbook.write(fileOut);
				fileOut.close();
				workbook.close();
				return true;
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			return false;

		}


}
