package com.demo.ExcelProject;

import java.sql.*;
 

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Properties;
import java.util.Calendar;
import java.text.DateFormatSymbols;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class CreateMonthlyReport {
	public static void main(String[] args) {

		Properties prop = new Properties();
		ArrayList<Users> a = new ArrayList();

		//**************************************
		//get value from properties file
		//**************************************

		String output = readConfigFile.main("app.output");
		String branchCodeConfig = readConfigFile.main("app.branchCode");

		String month = "";
		String year = "";
		String monthStr = "";
		String configMonth = readConfigFile.main("app.month");
		String configYear = readConfigFile.main("app.year");

		Calendar cal = Calendar.getInstance();
		int prevMonth = cal.get(Calendar.MONTH) ; // beware of month indexing from zero
		String prevMonthStr = new DateFormatSymbols().getMonths()[prevMonth - 1];
		SimpleDateFormat dateOnly = new SimpleDateFormat("MM/dd/yyyy HH:mm");
		System.out.println("date: " + dateOnly.format(cal.getTime()));
		System.out.println("month (config): " + configMonth);
		System.out.println("previous month: " + prevMonth);

		//if month in config file is not empty, it will generate report based on config file month.
		//if month in config file is empty, it will generate report for previous month

		if(configMonth.length() > 0){
			//config month got value. generate report for this month
			month = configMonth;
			monthStr = new DateFormatSymbols().getMonths()[Integer.parseInt(configMonth) - 1];
			if(configYear.equals("")){
				year = Integer.toString(cal.get(Calendar.YEAR));
			}else{
				year = configYear;
			}

		}else{
			//config month is empty. generate report for previous month
			month = Integer.toString(prevMonth);
			monthStr = prevMonthStr;

			if(month.equals("12")){
				year = Integer.toString(cal.get(Calendar.YEAR) - 1);
			}else{
				year = Integer.toString(cal.get(Calendar.YEAR));
			}

		}

		if(month.length() <2){
			month = "0"+month;
		}

		//**************************************
		//get data from db
		//**************************************
       	a =  getData(configYear, month, branchCodeConfig);

		//**************************************
		//create report
		//**************************************
		try {
			System.out.println("\n");
			System.out.println("Creating Report for "+ monthStr + " "+year+".....");
			System.out.println("\n");

			//Create workbook in .xlsx format
			Workbook workbook = new XSSFWorkbook();
			//For .xsl workbooks use new HSSFWorkbook();
			
			//Create Sheet
			Sheet sh = workbook.createSheet(monthStr + " " + year);

			//-------------------
			//Create main header
			//-------------------
			String[] mainHeader2 = {"","","", "AMANAH SAHAM NASIONAL BERHAD", "", "", "", "Printed Date:", dateOnly.format(cal.getTime())};
			String[] mainHeader3 = {"","","","","", "", "", "", "", "", "", ""};
			String[] mainHeader4 = {"","","", "Month:    "+ monthStr + " " + year, "", "", "","", "", ""};
			String[] mainHeader7 = {"","","", "Agent Code: "+branchCodeConfig,"","", "", "", "", ""};

			//We want to make table header bold with a foreground color.
			Font headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerFont.setFontHeightInPoints((short)10);
			headerFont.setColor(IndexedColors.BLACK.index);
			headerFont.setFontName("Arial");

			//Create a CellStyle table header with the font
			CellStyle headerStyleCenter = workbook.createCellStyle();
			headerStyleCenter.setFont(headerFont);
			headerStyleCenter.setAlignment(HorizontalAlignment.CENTER);
			headerStyleCenter.setVerticalAlignment(VerticalAlignment.TOP);
			headerStyleCenter.setWrapText(true);
			//headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			//headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);

			CellStyle headerStyleLeft = workbook.createCellStyle();
			headerStyleLeft.setFont(headerFont);
			headerStyleLeft.setAlignment(HorizontalAlignment.LEFT);
			headerStyleLeft.setVerticalAlignment(VerticalAlignment.TOP);
			headerStyleLeft.setWrapText(true);
			//headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			//headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);

			//Create the table header row at row no 1
			Row mainHeaderRow2 = sh.createRow(1);
			Row mainHeaderRow3 = sh.createRow(2);
			Row mainHeaderRow4 = sh.createRow(3);
			Row mainHeaderRow7 = sh.createRow(6);

			//Iterate over the column headings to create columns
			for(int i=0;i<mainHeader2.length;i++) {
				Cell cell = mainHeaderRow2.createCell(i);
				cell.setCellValue(mainHeader2[i]);
				cell.setCellStyle(headerStyleCenter);
			}

			for(int i=0;i<mainHeader3.length;i++) {
				Cell cell = mainHeaderRow3.createCell(i);
				cell.setCellValue(mainHeader3[i]);
				cell.setCellStyle(headerStyleCenter);
			}

			for(int i=0;i<mainHeader4.length;i++) {
				Cell cell = mainHeaderRow4.createCell(i);
				cell.setCellValue(mainHeader4[i]);
				cell.setCellStyle(headerStyleCenter);
			}

			for(int i=0;i<mainHeader7.length;i++) {
				Cell cell = mainHeaderRow7.createCell(i);
				cell.setCellValue(mainHeader7[i]);
				cell.setCellStyle(headerStyleCenter);
			}

			//automerge main header column
			sh.addMergedRegion(new CellRangeAddress(1, 1, 3, 6));
			sh.addMergedRegion(new CellRangeAddress(3, 3, 3, 6));
			sh.addMergedRegion(new CellRangeAddress(6, 6, 3, 6));

			//--------------------
			//Create table header
			//--------------------
			String[] columnHeadings = {"","NO","USER ID","USER NAME","ID DESCRIPTION","USER STATUS", "LAST SIGNED ON DATE", "NO OF RESET SINCE CREATION", "BRANCH CODE"};

			//Create the table header row at row no 8
			Row headerRow = sh.createRow(8);
			headerRow.setHeight((short)600);

			CellStyle tableHeaderStyle = workbook.createCellStyle();
			tableHeaderStyle.setFont(headerFont);
			tableHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
			tableHeaderStyle.setVerticalAlignment(VerticalAlignment.TOP);
			tableHeaderStyle.setWrapText(true);
			tableHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			tableHeaderStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
			tableHeaderStyle.setBorderTop(BorderStyle.THIN);
			tableHeaderStyle.setBorderBottom(BorderStyle.THIN);
			tableHeaderStyle.setBorderLeft(BorderStyle.THIN);
			tableHeaderStyle.setBorderRight(BorderStyle.THIN);

			//Iterate over the column headings to create columns
			for(int i=1;i<columnHeadings.length;i++) {
				Cell cell = headerRow.createCell(i);

				cell.setCellValue(columnHeadings[i]);
				cell.setCellStyle(tableHeaderStyle);

			}

			//Freeze Header Row
			//sh.createFreezePane(0, 12);

			//----------
			//Fill data
			//----------
			CreationHelper creationHelper= workbook.getCreationHelper();
			CellStyle dateStyle = workbook.createCellStyle();
			dateStyle.setDataFormat(creationHelper.createDataFormat().getFormat("dd-MMM-yyyy"));
			dateStyle.setBorderTop(BorderStyle.THIN);
			dateStyle.setBorderBottom(BorderStyle.THIN);
			dateStyle.setBorderLeft(BorderStyle.THIN);
			dateStyle.setBorderRight(BorderStyle.THIN);

			Font dataFont = workbook.createFont();
			dataFont.setFontHeightInPoints((short)10);
			dataFont.setColor(IndexedColors.BLACK.index);
			dataFont.setFontName("Arial");

			//Create a CellStyle table data with the font
			CellStyle dataStyle = workbook.createCellStyle();
			dataStyle.setFont(dataFont);
			dataStyle.setWrapText(true);
			dataStyle.setBorderTop(BorderStyle.THIN);
			dataStyle.setBorderBottom(BorderStyle.THIN);
			dataStyle.setBorderLeft(BorderStyle.THIN);
			dataStyle.setBorderRight(BorderStyle.THIN);

			CellStyle dataStyleCenter = workbook.createCellStyle();
			dataStyleCenter.setFont(dataFont);
			dataStyleCenter.setAlignment(HorizontalAlignment.CENTER);
			dataStyleCenter.setWrapText(true);
			dataStyleCenter.setBorderTop(BorderStyle.THIN);
			dataStyleCenter.setBorderBottom(BorderStyle.THIN);
			dataStyleCenter.setBorderLeft(BorderStyle.THIN);
			dataStyleCenter.setBorderRight(BorderStyle.THIN);

			int rownum =9;
			for(Users i : a) {
				//System.out.println("rownum-before"+(rownum));
				Row row = sh.createRow(rownum++);
				//System.out.println("rownum-after"+(rownum));
				Cell dataCell = row.createCell(1);
				dataCell.setCellValue(i.getItemId());
				dataCell.setCellStyle(dataStyleCenter);

				Cell dataCell2 = row.createCell(2);
				dataCell2.setCellValue(i.getUserid());
				dataCell2.setCellStyle(dataStyle);

				Cell dataCell3 = row.createCell(3);
				dataCell3.setCellValue(i.getUsername());
				dataCell3.setCellStyle(dataStyle);

				Cell dataCell4 = row.createCell(4);
				dataCell4.setCellValue(i.getIdDescription());
				dataCell4.setCellStyle(dataStyle);

				Cell dataCell5 = row.createCell(5);
				dataCell5.setCellValue(i.getUserStatus());
				dataCell5.setCellStyle(dataStyle);

			/*	Cell dateCell = row.createCell(6);
				dateCell.setCellValue(i.getDateOfStatus());
				dateCell.setCellStyle(dateStyle);
*/
				Cell dataCell6 = row.createCell(6);
				dataCell6.setCellValue(i.getLastSignedOnDate());
				dataCell6.setCellStyle(dateStyle);

				Cell dataCell7 = row.createCell(7);
				dataCell7.setCellValue(i.getNoOfReset());
				dataCell7.setCellStyle(dataStyle);

				Cell dataCell8 = row.createCell(8);
				dataCell8.setCellValue(i.getBranchCode());
				dataCell8.setCellStyle(dataStyle);


			}

			//Autosize columns
			for(int i=0;i<columnHeadings.length;i++) {
				sh.autoSizeColumn(i);
			}

			//set table header width
			sh.setColumnWidth(3, 10000);
			sh.setColumnWidth(5, 3000);
			//sh.setColumnWidth(6, 5000);
			sh.setColumnWidth(6, 5000);
			sh.setColumnWidth(7, 5000);
			sh.setColumnWidth(8, 5000);

			//for printing setup
			sh.setAutobreaks(true);
			sh.setFitToPage(Boolean.TRUE);
			PrintSetup ps = sh.getPrintSetup();
			ps.setFitWidth( (short) 1);
			ps.setFitHeight( (short) 0);

			//Sheet sh2 = workbook.createSheet("Second");
			//Write the output to file
			FileOutputStream fileOut = new FileOutputStream(output + "/"+year+"-"+month+"-"+branchCodeConfig+".xlsx");

			//print page no at footer
			Footer footer = sh.getFooter();
			footer.setCenter( "Page " + HeaderFooter.page() + " of " + HeaderFooter.numPages() );

			//write content to excel file
			workbook.write(fileOut);

			fileOut.close();
			workbook.close();
			System.out.println("Report Created at "+ output + "/"+year+"-"+month+"-"+branchCodeConfig+".xlsx");


			
		}
		catch(Exception e) {
			e.printStackTrace();
		}

	}

	private static ArrayList<Users> getData(String year, String month, String branchCodeConfig){

		int j = 1;
		ArrayList<Users> a = new ArrayList();
		Connection conn = null;

		try {

			String dbURL = readConfigFile.main("app.dbURL");
			String user = readConfigFile.main("app.user");
			String pass = readConfigFile.main("app.password");

			conn = DriverManager.getConnection(dbURL, user, pass);

			if (conn != null) {
				DatabaseMetaData dm = (DatabaseMetaData) conn.getMetaData();

				//declare the statement object
				Statement sqlStatement = conn.createStatement();

				//declare the result set
				ResultSet rs = null;

				//Build the query string, making sure to use column aliases
				String queryString = "select *,\n" +
						"(select max(timestamp) from Auditlog c where c.USER_ID = a.loginid and action = '02') as lastlogin,\n" +
						"(select count(id) from Auditlog c where c.USER_ID = a.loginid and action = '05') as noOfReset \n" +
						"FROM Users a\n" +
						"where a.branchcode like 'CIMB%'\n" +
						"order by a.branchcode, a.loginid";

				//print the query string to the screen
				System.out.println("\nQuery string:");
				System.out.println(queryString);

				//execute the query
				rs=sqlStatement.executeQuery(queryString);

				//loop through the result set and call method to print the result set row
				while (rs.next())
				{
					//printResultSetRow(rs);
					String userid= rs.getString("loginid");
					String username= rs.getString("name");
					String idDescription= rs.getString("role");
					String userStatus= rs.getString("user_status");
					//String dateOfStatus= rs.getString("modify_dtm");
					String dateOfStatus= null;
					String lastSignedOnDate= rs.getString("lastlogin");
					Integer noOfReset= Integer.parseInt(rs.getString("noOfReset"));
					//Integer noOfReset=0;
					String branchCode= rs.getString("branchcode");

					if (userStatus.equals("E")){
						userStatus = "Enable";
					}else if (userStatus.equals("D")){
						userStatus = "Disable";
					}else if (userStatus.equals("H")){
						userStatus = "Deleted";
					}else {
						userStatus = "";
					}

					try {

						String dateOfStatusStr = "";
						String lastSignedOnDateStr = "";

						if(dateOfStatus != null && lastSignedOnDate != null){
							a.add(new Users(j, userid, username, idDescription, userStatus, new SimpleDateFormat("yyyy-MM-dd").parse(dateOfStatus), new SimpleDateFormat("yyyy-MM-dd").parse(lastSignedOnDate), noOfReset, branchCode));
						}else if(dateOfStatus == null && lastSignedOnDate != null){
							a.add(new Users(j, userid, username, idDescription, userStatus, null, new SimpleDateFormat("yyyy-MM-dd").parse(lastSignedOnDate), noOfReset, branchCode));
						}else if(dateOfStatus != null && lastSignedOnDate == null){
							a.add(new Users(j, userid, username, idDescription, userStatus, new SimpleDateFormat("yyyy-MM-dd").parse(dateOfStatus), null, noOfReset, branchCode));
						}else{
							a.add(new Users(j, userid, username, idDescription, userStatus, null, null, noOfReset, branchCode));
						}

						j  = j+1;

					}catch (Exception e) {
						e.printStackTrace();
					}

				}

				//close the result set
				rs.close();

				//close the database connection
				conn.close();
			}

		} catch (SQLException ex) {
			System.err.println("Error connecting to the database");
			ex.printStackTrace(System.err);
			System.exit(0);
		} finally {
			try {
				if (conn != null && !conn.isClosed()) {
					conn.close();
				}
			} catch (SQLException ex) {
				ex.printStackTrace();
			}
		}

		return a;
	}

}
