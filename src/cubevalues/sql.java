package cubevalues;

import java.sql.*;
import java.util.Properties;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;

import org.apache.poi.hpsf.Property;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.olap4j.*;

public class sql {
	
	static int a;
	
	static String b;
	
	static File filename=new File("C:\\Users\\RBS\\Desktop\\workbook1.xlsx");
	

	
	
public static void main(String[] args) throws ClassNotFoundException, SQLException, IOException {
	
 Class.forName("org.olap4j.driver.xmla.XmlaOlap4jDriver");
	 OlapConnection con =
(OlapConnection)DriverManager.getConnection("jdbc:xmla:Server=http://CDTSOLAP937D/olap/msmdpump.dll;Catalog=WHSVC_SI_M_IMS_1");
	 
	 System.out.println("connected");
	 
	 OlapWrapper wrapper = (OlapWrapper) con;
	 OlapConnection olapConnection = wrapper.unwrap(OlapConnection.class);
	 OlapStatement stmt = olapConnection.createStatement();
	 System.out.println("Connection successfull");
	 
	 System.out.println("Running query");
	 
	 
	 CellSet cellSet = stmt.executeOlapQuery("select {[Measures].[INDIRECT ex-WHS VALUE],"
	 		+ "[Measures].[INDIRECT ex-WHS VALUE ±% PP],"
	 		+ "[Measures].[INDIRECT ex-WHS VALUE ±% YAGO],"
	 		+ "[Measures].[INDIRECT ex-WHS VALUE EI],"
	 		+ "[Measures].[INDIRECT ex-WHS VALUE MS],"
	 		+ "[Measures].[INDIRECT ex-WHS VALUE MS ± YAGO]} on columns from [WHS SALES] "
	 		+ "where ([Wholesaler].[Wholesaler].&[31],[Period].[FYTD - Month].[FY YTD].&[FYTD 02])");
	 
	 System.out.println("printing values");
	 System.out.println("-------------------------------------");
	 
	 
	 // creating a excel worksheet and writing values in to the cell
	 
	 FileInputStream fis= new FileInputStream(filename);
	 XSSFWorkbook workbook = new XSSFWorkbook(fis);
	 XSSFSheet worksheet= workbook.createSheet("cube calculation");
	 
	 XSSFRow row1=null;
	 XSSFCell cell=null;
	 
	 row1=worksheet.createRow(0);
	 worksheet.getRow(0);
	 cell= row1.createCell(1);
	 cell.setCellValue("cube values");
	 
 for (int i=0; i<=5 ;i++){
		 
		 Cell a = cellSet.getCell(i);
		 
		 b =a.getFormattedValue();
		 
		System.out.println( b);
		
		 
			 row1=worksheet.createRow(i+1);
			 worksheet.getRow(i+1);
			 cell= row1.createCell(1);
			 cell.setCellValue(b);
			 FileOutputStream fos =new FileOutputStream(filename);
			 workbook.write(fos);
			 
	 }
 
// System.out.println("Printing complete, writing to excel started");
// 
// 	 FileInputStream fis= new FileInputStream(filename);
//	 XSSFWorkbook workbook = new XSSFWorkbook(fis);
//	 XSSFSheet worksheet= workbook.createSheet("cube calculation");
//	
//	 for (int i=0; i<=5;i++){
//		 XSSFRow row1=worksheet.createRow(i);
//		 worksheet.getRow(i);
//		 XSSFCell cell= row1.createCell(0);
//		 cell.setCellValue(b);
		 
	 }
//	 XSSFRow row1=worksheet.createRow(0);
//	 worksheet.getRow(0);
//	 XSSFCell cell= row1.createCell(0);
//	 cell.setCellValue(b);
	 
//	 fis.close();
//	 FileOutputStream fos =new FileOutputStream(filename);
//	 workbook.write(fos);
//	 System.out.println("done");
	 
	 
	 
//		 double roundvalue=Math.round((double) a.getValue());
//		 System.out.println(roundvalue); 
	 
/*	 Cell a = cellSet.getCell(1);
	 System.out.println(a.getValue());
*/
	 
	 
	
	 
//	 System.out.println(a);
//	 
//	 List<CellSetAxis> cellSetAxes = cellSet.getAxes();
//	 System.out.print("\t");
//     CellSetAxis columnsAxis = cellSetAxes.get(Axis.COLUMNS.axisOrdinal());
//     for (Position position : columnsAxis.getPositions()) {
//         Member measure = position.getMembers().get(0);
//         System.out.print(measure.getName());
//     }
//     
//     // Print rows.
//     CellSetAxis rowsAxis = cellSetAxes.get(Axis.ROWS.axisOrdinal());
//    // int cellOrdinal = 0;
//     for (Position rowPosition : rowsAxis.getPositions()) {
//         boolean first = true;
//         for (Member member : rowPosition.getMembers()) {
//             if (first) {
//                 first = false;
//             } else {
//                 System.out.print('\t');
//             }
//             System.out.print(member.getName());
//         }
         
//         for (Position columnPosition : columnsAxis.getPositions()) {
//             // Access the cell via its ordinal. The ordinal is kept in step
//             // because we increment the ordinal once for each row and
//             // column.
//             Cell cell = cellSet.getCell(cellOrdinal);
//
//             // Just for kicks, convert the ordinal to a list of coordinates.
//             // The list matches the row and column positions.
//             List<Integer> coordList =
//                 cellSet.ordinalToCoordinates(cellOrdinal);
//             assert coordList.get(0) == rowPosition.getOrdinal();
//             assert coordList.get(1) == columnPosition.getOrdinal();
//
//             ++cellOrdinal;
//
//             System.out.print('\t');
//             System.out.print(cell.getFormattedValue());
//         }
//         System.out.println();
//     }
	 
	 
	 
	 //System.out.println(cellSet.getRow());
	 
	 
//	 DataSet ds = new DataSet();
//		
//	 for (Position rowPos : cellSet.getAxes().get(1)) {
//	   ds.addRow();
//	   for (Position colPos : cellSet.getAxes().get(0)) {
//	 	test += Integer.toString(rowPos.getOrdinal()) + " : " + Integer.toString(colPos.getOrdinal());
//	 	Cell cell = cellSet.getCell(colPos, rowPos);
//	 	test += "Value: " + cell.getFormattedValue() + "
//	 ";
//	 			 
//	 	ds.addValue("column" + Integer.toString(colPos.getOrdinal()), cell.getFormattedValue());
//	   }
//	 }

	 
	 
	
}


