package importsperformance;
import java.sql.*;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.*;
import java.util.*;
import java.io.*;


public class ImportsPerformance {
    // Database Credentials
    public static final String DBURL = "jdbc:oracle:thin:@10.128.0.220:1521/METROBIP";
    public static final String DBUSER = "arcma";
    public static final String DBPASS = "arcma";
    
    
    // Create Excel Worbook
     static Workbook workbook ;
     static Sheet sheet;
     static CellStyle style1,style2;
     static Font font,font2;
     static Connection con;
     static ResultSet rs;
     static PreparedStatement pStatement;
     Cell cell;
     static String date;
     static String mdate;
     static String qdate;
     static String ydate;
    
    
    public static void main(String[] args) throws SQLException, IOException {
        try{
            System.out.println("IMPORTS PERFORMANCE");
            workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream("C:\\Documents and Settings\\Administrator\\My Documents\\NetBeansProjects\\ImportsPerformance\\template\\ImportsTemplate.xls")));
            sheet = workbook.getSheet("Department");
            font = workbook.createFont();
            font2 = workbook.createFont();
            font.setFontName("Calibri");
            font2.setFontName("Calibri");
            font2.setFontHeightInPoints((short)10);
            style1 = workbook.createCellStyle();
            style2 = workbook.createCellStyle();
            style1.setBorderBottom(CellStyle.BORDER_THIN);
            style1.setBorderRight(CellStyle.BORDER_THIN);
            style1.setBorderLeft(CellStyle.BORDER_THIN);
            style1.setBorderTop(CellStyle.BORDER_THIN);
            style1.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0.00"));
            style1.setFont(font);
            style2.setFont(font2);
            
            
        
            ImportsPerformance  importsPerf = new ImportsPerformance();
            
            
            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
            con = DriverManager.getConnection(DBURL, DBUSER, DBPASS);
            
            importsPerf.getDate();
            System.out.println("- "+date+" -");
            
            //MTD (TOTAL & IMPORTS)
            //importsPerf.totalSalesMTD();
            //importsPerf.importsTotalSalesMTD();
            //importsPerf.budgetTotalSalesMTD();
            //importsPerf.budgetImportsSalesMTD();
            
            //QTD (TOTAL & IMPORTS)
            //importsPerf.totalSalesQTD();
            //importsPerf.importsTotalSalesQTD();
            //importsPerf.budgetTotalSalesQTD();
            //importsPerf.budgetImportsSalesQTD();
            
            //QTD (TOTAL & IMPORTS)
            //importsPerf.totalSalesYTD();
            //importsPerf.importsTotalSalesYTD();
            //importsPerf.budgetTotalSalesYTD();
            //importsPerf.budgetImportsSalesYTD();
            
           
            rs.close();
            pStatement.close();
            con.close();
              
            System.out.println("Writing report..");
            HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
            workbook.setForceFormulaRecalculation(true);
            //System.out.println(date);
            FileOutputStream output = new FileOutputStream("C:\\Documents and Settings\\Administrator\\My Documents\\NetBeansProjects\\ImportsPerformance\\output\\Importation Performance (as of "+date+").xls");
            workbook.write(output);
            output.close();
           
            System.out.println("Report Completed.");
            
             //SEND EMAIL
            importsPerf.SendEmail();
            
         
        }
        catch (Exception e){
            System.out.println(e);
        }
       
           
    }
    
    private void getDate(){
        try{
            String query = "Select to_char(sysdate-1,'DD MON YYYY') as Sys_Date,"
                    + "to_char(trunc(sysdate-1, 'MM'),'DD MON YYYY') as mDate,"
                    + "to_char(trunc(sysdate-1, 'Q'),'DD MON YYYY') as qDate,"
                    + "to_char(trunc(sysdate-1, 'YEAR'),'DD MON YYYY') as yDate "
                    + "from DUAL";
            pStatement = con.prepareStatement(query);
            rs = pStatement.executeQuery();
            while (rs.next()){
                date = rs.getString("Sys_date");
                mdate = rs.getString("mdate");
                qdate = rs.getString("qdate");
                ydate = rs.getString("ydate");  
            }
         
            cell = sheet.getRow(1).createCell(0);
            cell.setCellValue("MTD: "+mdate+" - "+date);
            cell.setCellStyle(style2);
            
            cell = sheet.getRow(2).createCell(0);
            cell.setCellValue("QTD: "+qdate+" - "+date);
            cell.setCellStyle(style2);
            
            cell = sheet.getRow(3).createCell(0);
            cell.setCellValue("YTD: "+ydate+" - "+date);
            cell.setCellStyle(style2);
            
            cell = sheet.getRow(4).createCell(0);
            cell.setCellValue("As of "+date);
            cell.setCellStyle(style2);     
            
        }
        catch(Exception e){
            System.out.println(e);
        }
    }
    
    
    private void totalSalesMTD() throws SQLException{
 
       
        System.out.println("Fetching Total Sales(MTD)..");
        
   
   
           try{
                String query = "Select t1.department_code as DEPT_CODE, t2.NET_SALES from " +
                            "(select distinct department_code from dim_product where department_code in (3510,3520,3550,"+
                            "3560,3590,4510,4520,4540,5020,5030,5040,7010,7020,7030,7040,7050,7070," +
                            "7080,2510,2530,2540,2550,2560,2570,3010,3020,3030,3040,3050,3060,3080,6540,"+
                            "6550,6560,7550,9510,9520,9530,8010,8020,8030,1510,1520,1530,1540,1010,"+
                            "1020,1030,1040,1050,1060,2020,2030,2040,2050,5520,5530,8040,6010,"+
                            "8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030)) t1 "+
                            "left join (SELECT dp.department_code as DEPT_CODE, " +
                            "sum(NVL(ag.SALE_NET_VAL,0)-NVL(ag.SALE_TOT_TAX_VAL,0)-NVL(SALE_TOT_DISC_VAL,0)) "+
                            "as NET_SALES "+
                            "FROM agg_dly_str_prod ag "+
                            "LEFT JOIN dim_date d on ag.date_key = d.date_key "+
                            "LEFT JOIN dim_product dp on ag.product_code = dp.product_code "+
                            "where dp.department_code in (3510,3520,3550,3560,3590,4510,4520,4540,5020,5030,5040"+
                            ",7010,7020,7030,7040,7050,7070,7080,2510,2530,2540,2550,2560,2570,3010,3020,3030"+
                            ",3040,3050,3060,3080,6540,6550,6560,7550,9510,9520,9530,8010,8020,8030"+
                            ",1510,1520,1530,1540,1010,1020,1030,1040,1050,1060,2020,2030,2040,2050,5520"+
                            ",5530,8040,6010,8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030) "+
                            "and d.date_fld between trunc(sysdate-1,'MM') AND sysdate-1 "+
                            "and dp.latest = 1 and dp.active = 1 and dp.concession_flg = 'N' "+
                            "GROUP BY dp.department_code) "+
                            "t2 on t1.department_code = t2.DEPT_CODE "+
                            "order by case " +
                                  "when t1.department_code = 3510 then 1 "+
                                  "when t1.department_code = 3520 then 2 "+
                                  "when t1.department_code = 3550 then 3 "+
                                  "when t1.department_code = 3560 then 4 "+
                                  "when t1.department_code = 3590 then 5 "+
                                  "when t1.department_code = 4510 then 6 "+
                                  "when t1.department_code = 4520 then 7 "+
                                  "when t1.department_code = 4540 then 8 "+
                                  "when t1.department_code = 5020 then 9 "+
                                  "when t1.department_code = 5030 then 10 "+
                                  "when t1.department_code = 5040 then 11 "+
                                  "when t1.department_code = 7010 then 12 "+
                                  "when t1.department_code = 7020 then 13 "+
                                  "when t1.department_code = 7030 then 14 "+
                                  "when t1.department_code = 7040 then 15 "+
                                  "when t1.department_code = 7050 then 16 "+
                                  "when t1.department_code = 7070 then 17 "+
                                  "when t1.department_code = 7080 then 18 "+
                                  "when t1.department_code = 2510 then 19 "+
                                  "when t1.department_code = 2530 then 20 "+
                                  "when t1.department_code = 2540 then 21 "+
                                  "when t1.department_code = 2550 then 22 "+
                                  "when t1.department_code = 2560 then 23 "+
                                  "when t1.department_code = 2570 then 24 "+
                                  "when t1.department_code = 3010 then 25 "+
                                  "when t1.department_code = 3020 then 26 "+
                                  "when t1.department_code = 3030 then 27 "+
                                  "when t1.department_code = 3040 then 28 "+
                                  "when t1.department_code = 3050 then 29 "+
                                  "when t1.department_code = 3060 then 30 "+
                                  "when t1.department_code = 3080 then 31 "+
                                  "when t1.department_code = 6540 then 32 "+
                                  "when t1.department_code = 6550 then 33 "+
                                  "when t1.department_code = 6560 then 34 "+
                                  "when t1.department_code = 7550 then 35 "+
                                  "when t1.department_code = 9510 then 36 "+
                                  "when t1.department_code = 9520 then 37 "+
                                  "when t1.department_code = 9530 then 38 "+
                                  "when t1.department_code = 8010 then 39 "+
                                  "when t1.department_code = 8020 then 40 "+
                                  "when t1.department_code = 8030 then 41 "+
                                  "when t1.department_code = 1510 then 42 "+
                                  "when t1.department_code = 1520 then 43 "+
                                  "when t1.department_code = 1530 then 44 "+
                                  "when t1.department_code = 1540 then 45 "+
                                  "when t1.department_code = 1010 then 46 "+
                                  "when t1.department_code = 1020 then 47 "+
                                  "when t1.department_code = 1030 then 48 "+
                                  "when t1.department_code = 1040 then 49 "+
                                  "when t1.department_code = 1050 then 50 "+
                                  "when t1.department_code = 1060 then 51 "+
                                  "when t1.department_code = 2020 then 52 "+
                                  "when t1.department_code = 2030 then 53 "+
                                  "when t1.department_code = 2040 then 54 "+
                                  "when t1.department_code = 2050 then 55 "+
                                  "when t1.department_code = 5520 then 56 "+
                                  "when t1.department_code = 5530 then 57 "+
                                  "when t1.department_code = 8040 then 58 "+
                                  "when t1.department_code = 6010 then 59 "+
                                  "when t1.department_code = 8510 then 60 "+
                                  "when t1.department_code = 8520 then 61 "+
                                  "when t1.department_code = 8530 then 62 "+
                                  "when t1.department_code = 8540 then 63 "+
                                  "when t1.department_code = 8550 then 64 "+
                                  "when t1.department_code = 8560 then 65 "+
                                  "when t1.department_code = 8570 then 66 "+
                                  "when t1.department_code = 8590 then 67 "+
                                  "when t1.department_code = 9010 then 68 "+
                                  "when t1.department_code = 9020 then 69 "+
                                  "when t1.department_code = 9030 then 70 end";
                
                pStatement = con.prepareStatement(query);
                rs = pStatement.executeQuery();

                
                int row = 9;

                
                while(rs.next()){
                    if(row == 14 || row == 18 || row == 22 || row == 38 || row == 46 || row == 50 || row == 52){
                        row++;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(5);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 30 || row == 56 ||   row == 68 || row == 76 || row == 82 || row == 89 || row == 92 || row == 86 ){
                        row = row + 2;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(5);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 61 || row == 102 ){
                        row = row + 3;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(5);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else{
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(5);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    row ++;
                }
                
            }
            catch(Exception e){
                System.out.println(e);
            }
        } 
    private void totalSalesQTD() throws SQLException{
 
       
        System.out.println("Fetching Total Sales(QTD)..");
        
   
   
           try{
                String query = "Select t1.department_code as DEPT_CODE, t2.NET_SALES from " +
                            "(select distinct department_code from dim_product where department_code in (3510,3520,3550,"+
                            "3560,3590,4510,4520,4540,5020,5030,5040,7010,7020,7030,7040,7050,7070," +
                            "7080,2510,2530,2540,2550,2560,2570,3010,3020,3030,3040,3050,3060,3080,6540,"+
                            "6550,6560,7550,9510,9520,9530,8010,8020,8030,1510,1520,1530,1540,1010,"+
                            "1020,1030,1040,1050,1060,2020,2030,2040,2050,5520,5530,8040,6010,"+
                            "8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030)) t1 "+
                            "left join (SELECT dp.department_code as DEPT_CODE, " +
                            "sum(NVL(ag.SALE_NET_VAL,0)-NVL(ag.SALE_TOT_TAX_VAL,0)-NVL(SALE_TOT_DISC_VAL,0)) "+
                            "as NET_SALES "+
                            "FROM agg_dly_str_prod ag "+
                            "LEFT JOIN dim_date d on ag.date_key = d.date_key "+
                            "LEFT JOIN dim_product dp on ag.product_code = dp.product_code "+
                            "where dp.department_code in (3510,3520,3550,3560,3590,4510,4520,4540,5020,5030,5040"+
                            ",7010,7020,7030,7040,7050,7070,7080,2510,2530,2540,2550,2560,2570,3010,3020,3030"+
                            ",3040,3050,3060,3080,6540,6550,6560,7550,9510,9520,9530,8010,8020,8030"+
                            ",1510,1520,1530,1540,1010,1020,1030,1040,1050,1060,2020,2030,2040,2050,5520"+
                            ",5530,8040,6010,8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030) "+
                            "and d.date_fld between trunc(sysdate-1,'Q') AND sysdate-1 "+
                            "and dp.latest = 1 and dp.active = 1 and dp.concession_flg = 'N' "+
                            "GROUP BY dp.department_code) "+
                            "t2 on t1.department_code = t2.DEPT_CODE "+
                            "order by case " +
                                  "when t1.department_code = 3510 then 1 "+
                                  "when t1.department_code = 3520 then 2 "+
                                  "when t1.department_code = 3550 then 3 "+
                                  "when t1.department_code = 3560 then 4 "+
                                  "when t1.department_code = 3590 then 5 "+
                                  "when t1.department_code = 4510 then 6 "+
                                  "when t1.department_code = 4520 then 7 "+
                                  "when t1.department_code = 4540 then 8 "+
                                  "when t1.department_code = 5020 then 9 "+
                                  "when t1.department_code = 5030 then 10 "+
                                  "when t1.department_code = 5040 then 11 "+
                                  "when t1.department_code = 7010 then 12 "+
                                  "when t1.department_code = 7020 then 13 "+
                                  "when t1.department_code = 7030 then 14 "+
                                  "when t1.department_code = 7040 then 15 "+
                                  "when t1.department_code = 7050 then 16 "+
                                  "when t1.department_code = 7070 then 17 "+
                                  "when t1.department_code = 7080 then 18 "+
                                  "when t1.department_code = 2510 then 19 "+
                                  "when t1.department_code = 2530 then 20 "+
                                  "when t1.department_code = 2540 then 21 "+
                                  "when t1.department_code = 2550 then 22 "+
                                  "when t1.department_code = 2560 then 23 "+
                                  "when t1.department_code = 2570 then 24 "+
                                  "when t1.department_code = 3010 then 25 "+
                                  "when t1.department_code = 3020 then 26 "+
                                  "when t1.department_code = 3030 then 27 "+
                                  "when t1.department_code = 3040 then 28 "+
                                  "when t1.department_code = 3050 then 29 "+
                                  "when t1.department_code = 3060 then 30 "+
                                  "when t1.department_code = 3080 then 31 "+
                                  "when t1.department_code = 6540 then 32 "+
                                  "when t1.department_code = 6550 then 33 "+
                                  "when t1.department_code = 6560 then 34 "+
                                  "when t1.department_code = 7550 then 35 "+
                                  "when t1.department_code = 9510 then 36 "+
                                  "when t1.department_code = 9520 then 37 "+
                                  "when t1.department_code = 9530 then 38 "+
                                  "when t1.department_code = 8010 then 39 "+
                                  "when t1.department_code = 8020 then 40 "+
                                  "when t1.department_code = 8030 then 41 "+
                                  "when t1.department_code = 1510 then 42 "+
                                  "when t1.department_code = 1520 then 43 "+
                                  "when t1.department_code = 1530 then 44 "+
                                  "when t1.department_code = 1540 then 45 "+
                                  "when t1.department_code = 1010 then 46 "+
                                  "when t1.department_code = 1020 then 47 "+
                                  "when t1.department_code = 1030 then 48 "+
                                  "when t1.department_code = 1040 then 49 "+
                                  "when t1.department_code = 1050 then 50 "+
                                  "when t1.department_code = 1060 then 51 "+
                                  "when t1.department_code = 2020 then 52 "+
                                  "when t1.department_code = 2030 then 53 "+
                                  "when t1.department_code = 2040 then 54 "+
                                  "when t1.department_code = 2050 then 55 "+
                                  "when t1.department_code = 5520 then 56 "+
                                  "when t1.department_code = 5530 then 57 "+
                                  "when t1.department_code = 8040 then 58 "+
                                  "when t1.department_code = 6010 then 59 "+
                                  "when t1.department_code = 8510 then 60 "+
                                  "when t1.department_code = 8520 then 61 "+
                                  "when t1.department_code = 8530 then 62 "+
                                  "when t1.department_code = 8540 then 63 "+
                                  "when t1.department_code = 8550 then 64 "+
                                  "when t1.department_code = 8560 then 65 "+
                                  "when t1.department_code = 8570 then 66 "+
                                  "when t1.department_code = 8590 then 67 "+
                                  "when t1.department_code = 9010 then 68 "+
                                  "when t1.department_code = 9020 then 69 "+
                                  "when t1.department_code = 9030 then 70 end";
                
                pStatement = con.prepareStatement(query);
                rs = pStatement.executeQuery();

                
                 int row = 9;

                
                while(rs.next()){
                    if(row == 14 || row == 18 || row == 22 || row == 38 || row == 46 || row == 50 || row == 52){
                        row++;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(12);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 30 || row == 56 ||   row == 68 || row == 76 || row == 82 || row == 89 || row == 92 || row == 86 ){
                        row = row + 2;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(12);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 61 || row == 102 ){
                        row = row + 3;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(12);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else{
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(12);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    row ++;
                }
                
     
            }
            catch(Exception e){
                System.out.println(e);
            }
        } 
    private void totalSalesYTD() throws SQLException{
 
       
        System.out.println("Fetching Total Sales(YTD)..");
        
   
   
           try{
                String query = "Select t1.department_code as DEPT_CODE, t2.NET_SALES from " +
                            "(select distinct department_code from dim_product where department_code in (3510,3520,3550,"+
                            "3560,3590,4510,4520,4540,5020,5030,5040,7010,7020,7030,7040,7050,7070," +
                            "7080,2510,2530,2540,2550,2560,2570,3010,3020,3030,3040,3050,3060,3080,6540,"+
                            "6550,6560,7550,9510,9520,9530,8010,8020,8030,1510,1520,1530,1540,1010,"+
                            "1020,1030,1040,1050,1060,2020,2030,2040,2050,5520,5530,8040,6010,"+
                            "8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030)) t1 "+
                            "left join (SELECT dp.department_code as DEPT_CODE, " +
                            "sum(NVL(ag.SALE_NET_VAL,0)-NVL(ag.SALE_TOT_TAX_VAL,0)-NVL(SALE_TOT_DISC_VAL,0)) "+
                            "as NET_SALES "+
                            "FROM agg_dly_str_prod ag "+
                            "LEFT JOIN dim_date d on ag.date_key = d.date_key "+
                            "LEFT JOIN dim_product dp on ag.product_code = dp.product_code "+
                            "where dp.department_code in (3510,3520,3550,3560,3590,4510,4520,4540,5020,5030,5040"+
                            ",7010,7020,7030,7040,7050,7070,7080,2510,2530,2540,2550,2560,2570,3010,3020,3030"+
                            ",3040,3050,3060,3080,6540,6550,6560,7550,9510,9520,9530,8010,8020,8030"+
                            ",1510,1520,1530,1540,1010,1020,1030,1040,1050,1060,2020,2030,2040,2050,5520"+
                            ",5530,8040,6010,8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030) "+
                            "and d.date_fld between trunc(sysdate-1,'YEAR') AND sysdate-1 "+
                            "and dp.latest = 1 and dp.active = 1 and dp.concession_flg = 'N' "+
                            "GROUP BY dp.department_code) "+
                            "t2 on t1.department_code = t2.DEPT_CODE "+
                            "order by case " +
                                  "when t1.department_code = 3510 then 1 "+
                                  "when t1.department_code = 3520 then 2 "+
                                  "when t1.department_code = 3550 then 3 "+
                                  "when t1.department_code = 3560 then 4 "+
                                  "when t1.department_code = 3590 then 5 "+
                                  "when t1.department_code = 4510 then 6 "+
                                  "when t1.department_code = 4520 then 7 "+
                                  "when t1.department_code = 4540 then 8 "+
                                  "when t1.department_code = 5020 then 9 "+
                                  "when t1.department_code = 5030 then 10 "+
                                  "when t1.department_code = 5040 then 11 "+
                                  "when t1.department_code = 7010 then 12 "+
                                  "when t1.department_code = 7020 then 13 "+
                                  "when t1.department_code = 7030 then 14 "+
                                  "when t1.department_code = 7040 then 15 "+
                                  "when t1.department_code = 7050 then 16 "+
                                  "when t1.department_code = 7070 then 17 "+
                                  "when t1.department_code = 7080 then 18 "+
                                  "when t1.department_code = 2510 then 19 "+
                                  "when t1.department_code = 2530 then 20 "+
                                  "when t1.department_code = 2540 then 21 "+
                                  "when t1.department_code = 2550 then 22 "+
                                  "when t1.department_code = 2560 then 23 "+
                                  "when t1.department_code = 2570 then 24 "+
                                  "when t1.department_code = 3010 then 25 "+
                                  "when t1.department_code = 3020 then 26 "+
                                  "when t1.department_code = 3030 then 27 "+
                                  "when t1.department_code = 3040 then 28 "+
                                  "when t1.department_code = 3050 then 29 "+
                                  "when t1.department_code = 3060 then 30 "+
                                  "when t1.department_code = 3080 then 31 "+
                                  "when t1.department_code = 6540 then 32 "+
                                  "when t1.department_code = 6550 then 33 "+
                                  "when t1.department_code = 6560 then 34 "+
                                  "when t1.department_code = 7550 then 35 "+
                                  "when t1.department_code = 9510 then 36 "+
                                  "when t1.department_code = 9520 then 37 "+
                                  "when t1.department_code = 9530 then 38 "+
                                  "when t1.department_code = 8010 then 39 "+
                                  "when t1.department_code = 8020 then 40 "+
                                  "when t1.department_code = 8030 then 41 "+
                                  "when t1.department_code = 1510 then 42 "+
                                  "when t1.department_code = 1520 then 43 "+
                                  "when t1.department_code = 1530 then 44 "+
                                  "when t1.department_code = 1540 then 45 "+
                                  "when t1.department_code = 1010 then 46 "+
                                  "when t1.department_code = 1020 then 47 "+
                                  "when t1.department_code = 1030 then 48 "+
                                  "when t1.department_code = 1040 then 49 "+
                                  "when t1.department_code = 1050 then 50 "+
                                  "when t1.department_code = 1060 then 51 "+
                                  "when t1.department_code = 2020 then 52 "+
                                  "when t1.department_code = 2030 then 53 "+
                                  "when t1.department_code = 2040 then 54 "+
                                  "when t1.department_code = 2050 then 55 "+
                                  "when t1.department_code = 5520 then 56 "+
                                  "when t1.department_code = 5530 then 57 "+
                                  "when t1.department_code = 8040 then 58 "+
                                  "when t1.department_code = 6010 then 59 "+
                                  "when t1.department_code = 8510 then 60 "+
                                  "when t1.department_code = 8520 then 61 "+
                                  "when t1.department_code = 8530 then 62 "+
                                  "when t1.department_code = 8540 then 63 "+
                                  "when t1.department_code = 8550 then 64 "+
                                  "when t1.department_code = 8560 then 65 "+
                                  "when t1.department_code = 8570 then 66 "+
                                  "when t1.department_code = 8590 then 67 "+
                                  "when t1.department_code = 9010 then 68 "+
                                  "when t1.department_code = 9020 then 69 "+
                                  "when t1.department_code = 9030 then 70 end";
                pStatement = con.prepareStatement(query);
                rs = pStatement.executeQuery();

                
                int row = 9;

                while(rs.next()){
                    if(row == 14 || row == 18 || row == 22 || row == 38 || row == 46 || row == 50 || row == 52){
                        row++;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(19);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 30 || row == 56 ||   row == 68 || row == 76 || row == 82 || row == 89 || row == 92 || row == 86 ){
                        row = row + 2;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(19);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 61 || row == 102 ){
                        row = row + 3;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(19);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else{
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(19);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    row ++;
                }
         
            }
            catch(Exception e){
                System.out.println(e);
            }
        } 
    
     
    private void importsTotalSalesMTD() throws SQLException{
 
       
        System.out.println("Fetching Import Sales(MTD)..");
        
           try{
                String query = "Select t1.department_code as DEPT_CODE, t2.NET_SALES from " +
                            "(select distinct department_code from dim_product where department_code in (3510,3520,3550,"+
                            "3560,3590,4510,4520,4540,5020,5030,5040,7010,7020,7030,7040,7050,7070," +
                            "7080,2510,2530,2540,2550,2560,2570,3010,3020,3030,3040,3050,3060,3080,6540,"+
                            "6550,6560,7550,9510,9520,9530,8010,8020,8030,1510,1520,1530,1540,1010,"+
                            "1020,1030,1040,1050,1060,2020,2030,2040,2050,5520,5530,8040,6010,"+
                            "8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030)) t1 "+
                            "left join (SELECT dp.department_code as DEPT_CODE, " +
                            "sum(NVL(ag.SALE_NET_VAL,0)-NVL(ag.SALE_TOT_TAX_VAL,0)-NVL(SALE_TOT_DISC_VAL,0)) "+
                            "as NET_SALES "+
                            "FROM agg_dly_str_prod ag "+
                            "LEFT JOIN dim_date d on ag.date_key = d.date_key "+
                            "LEFT JOIN dim_product dp on ag.product_code = dp.product_code "+
                            "Left join uda_item_lov uda on ag.product_code = uda.item "+
                            "where dp.department_code in (3510,3520,3550,3560,3590,4510,4520,4540,5020,5030,5040"+
                            ",7010,7020,7030,7040,7050,7070,7080,2510,2530,2540,2550,2560,2570,3010,3020,3030"+
                            ",3040,3050,3060,3080,6540,6550,6560,7550,9510,9520,9530,8010,8020,8030"+
                            ",1510,1520,1530,1540,1010,1020,1030,1040,1050,1060,2020,2030,2040,2050,5520"+
                            ",5530,8040,6010,8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030) "+
                            "and d.date_fld between trunc(sysdate-1,'MM') AND sysdate-1 "+
                            "and dp.latest = 1 and dp.active = 1 and dp.concession_flg = 'N' "+
                            "and uda_id = 1204 "+
                            "GROUP BY dp.department_code) "+
                            "t2 on t1.department_code = t2.DEPT_CODE "+
                            "order by case " +
                                  "when t1.department_code = 3510 then 1 "+
                                  "when t1.department_code = 3520 then 2 "+
                                  "when t1.department_code = 3550 then 3 "+
                                  "when t1.department_code = 3560 then 4 "+
                                  "when t1.department_code = 3590 then 5 "+
                                  "when t1.department_code = 4510 then 6 "+
                                  "when t1.department_code = 4520 then 7 "+
                                  "when t1.department_code = 4540 then 8 "+
                                  "when t1.department_code = 5020 then 9 "+
                                  "when t1.department_code = 5030 then 10 "+
                                  "when t1.department_code = 5040 then 11 "+
                                  "when t1.department_code = 7010 then 12 "+
                                  "when t1.department_code = 7020 then 13 "+
                                  "when t1.department_code = 7030 then 14 "+
                                  "when t1.department_code = 7040 then 15 "+
                                  "when t1.department_code = 7050 then 16 "+
                                  "when t1.department_code = 7070 then 17 "+
                                  "when t1.department_code = 7080 then 18 "+
                                  "when t1.department_code = 2510 then 19 "+
                                  "when t1.department_code = 2530 then 20 "+
                                  "when t1.department_code = 2540 then 21 "+
                                  "when t1.department_code = 2550 then 22 "+
                                  "when t1.department_code = 2560 then 23 "+
                                  "when t1.department_code = 2570 then 24 "+
                                  "when t1.department_code = 3010 then 25 "+
                                  "when t1.department_code = 3020 then 26 "+
                                  "when t1.department_code = 3030 then 27 "+
                                  "when t1.department_code = 3040 then 28 "+
                                  "when t1.department_code = 3050 then 29 "+
                                  "when t1.department_code = 3060 then 30 "+
                                  "when t1.department_code = 3080 then 31 "+
                                  "when t1.department_code = 6540 then 32 "+
                                  "when t1.department_code = 6550 then 33 "+
                                  "when t1.department_code = 6560 then 34 "+
                                  "when t1.department_code = 7550 then 35 "+
                                  "when t1.department_code = 9510 then 36 "+
                                  "when t1.department_code = 9520 then 37 "+
                                  "when t1.department_code = 9530 then 38 "+
                                  "when t1.department_code = 8010 then 39 "+
                                  "when t1.department_code = 8020 then 40 "+
                                  "when t1.department_code = 8030 then 41 "+
                                  "when t1.department_code = 1510 then 42 "+
                                  "when t1.department_code = 1520 then 43 "+
                                  "when t1.department_code = 1530 then 44 "+
                                  "when t1.department_code = 1540 then 45 "+
                                  "when t1.department_code = 1010 then 46 "+
                                  "when t1.department_code = 1020 then 47 "+
                                  "when t1.department_code = 1030 then 48 "+
                                  "when t1.department_code = 1040 then 49 "+
                                  "when t1.department_code = 1050 then 50 "+
                                  "when t1.department_code = 1060 then 51 "+
                                  "when t1.department_code = 2020 then 52 "+
                                  "when t1.department_code = 2030 then 53 "+
                                  "when t1.department_code = 2040 then 54 "+
                                  "when t1.department_code = 2050 then 55 "+
                                  "when t1.department_code = 5520 then 56 "+
                                  "when t1.department_code = 5530 then 57 "+
                                  "when t1.department_code = 8040 then 58 "+
                                  "when t1.department_code = 6010 then 59 "+
                                  "when t1.department_code = 8510 then 60 "+
                                  "when t1.department_code = 8520 then 61 "+
                                  "when t1.department_code = 8530 then 62 "+
                                  "when t1.department_code = 8540 then 63 "+
                                  "when t1.department_code = 8550 then 64 "+
                                  "when t1.department_code = 8560 then 65 "+
                                  "when t1.department_code = 8570 then 66 "+
                                  "when t1.department_code = 8590 then 67 "+
                                  "when t1.department_code = 9010 then 68 "+
                                  "when t1.department_code = 9020 then 69 "+
                                  "when t1.department_code = 9030 then 70 end";
                
                pStatement = con.prepareStatement(query);
                rs = pStatement.executeQuery();

                
                 int row = 9;

                
                while(rs.next()){
                    if(row == 14 || row == 18 || row == 22 || row == 38 || row == 46 || row == 50 || row == 52){
                        row++;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(6);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 30 || row == 56 ||   row == 68 || row == 76 || row == 82 || row == 89 || row == 92 || row == 86 ){
                        row = row + 2;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(6);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 61 || row == 102 ){
                        row = row + 3;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(6);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else{
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(6);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    row ++;
                }
            }
            catch(Exception e){
                System.out.println(e);
            }
        }  
    private void importsTotalSalesQTD() throws SQLException{
 
       
        System.out.println("Fetching Import Sales(QTD)..");
        
           try{
                String query = "Select t1.department_code as DEPT_CODE, t2.NET_SALES from " +
                            "(select distinct department_code from dim_product where department_code in (3510,3520,3550,"+
                            "3560,3590,4510,4520,4540,5020,5030,5040,7010,7020,7030,7040,7050,7070," +
                            "7080,2510,2530,2540,2550,2560,2570,3010,3020,3030,3040,3050,3060,3080,6540,"+
                            "6550,6560,7550,9510,9520,9530,8010,8020,8030,1510,1520,1530,1540,1010,"+
                            "1020,1030,1040,1050,1060,2020,2030,2040,2050,5520,5530,8040,6010,"+
                            "8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030)) t1 "+
                            "left join (SELECT dp.department_code as DEPT_CODE, " +
                            "sum(NVL(ag.SALE_NET_VAL,0)-NVL(ag.SALE_TOT_TAX_VAL,0)-NVL(SALE_TOT_DISC_VAL,0)) "+
                            "as NET_SALES "+
                            "FROM agg_dly_str_prod ag "+
                            "LEFT JOIN dim_date d on ag.date_key = d.date_key "+
                            "LEFT JOIN dim_product dp on ag.product_code = dp.product_code "+
                            "Left join uda_item_lov uda on ag.product_code = uda.item "+
                            "where dp.department_code in (3510,3520,3550,3560,3590,4510,4520,4540,5020,5030,5040"+
                            ",7010,7020,7030,7040,7050,7070,7080,2510,2530,2540,2550,2560,2570,3010,3020,3030"+
                            ",3040,3050,3060,3080,6540,6550,6560,7550,9510,9520,9530,8010,8020,8030"+
                            ",1510,1520,1530,1540,1010,1020,1030,1040,1050,1060,2020,2030,2040,2050,5520"+
                            ",5530,8040,6010,8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030) "+
                            "and d.date_fld between trunc(sysdate-1,'Q') AND sysdate-1 "+
                            "and dp.latest = 1 and dp.active = 1 and dp.concession_flg = 'N' "+
                            "and uda_id = 1204 "+
                            "GROUP BY dp.department_code) "+
                            "t2 on t1.department_code = t2.DEPT_CODE "+
                            "order by case " +
                                  "when t1.department_code = 3510 then 1 "+
                                  "when t1.department_code = 3520 then 2 "+
                                  "when t1.department_code = 3550 then 3 "+
                                  "when t1.department_code = 3560 then 4 "+
                                  "when t1.department_code = 3590 then 5 "+
                                  "when t1.department_code = 4510 then 6 "+
                                  "when t1.department_code = 4520 then 7 "+
                                  "when t1.department_code = 4540 then 8 "+
                                  "when t1.department_code = 5020 then 9 "+
                                  "when t1.department_code = 5030 then 10 "+
                                  "when t1.department_code = 5040 then 11 "+
                                  "when t1.department_code = 7010 then 12 "+
                                  "when t1.department_code = 7020 then 13 "+
                                  "when t1.department_code = 7030 then 14 "+
                                  "when t1.department_code = 7040 then 15 "+
                                  "when t1.department_code = 7050 then 16 "+
                                  "when t1.department_code = 7070 then 17 "+
                                  "when t1.department_code = 7080 then 18 "+
                                  "when t1.department_code = 2510 then 19 "+
                                  "when t1.department_code = 2530 then 20 "+
                                  "when t1.department_code = 2540 then 21 "+
                                  "when t1.department_code = 2550 then 22 "+
                                  "when t1.department_code = 2560 then 23 "+
                                  "when t1.department_code = 2570 then 24 "+
                                  "when t1.department_code = 3010 then 25 "+
                                  "when t1.department_code = 3020 then 26 "+
                                  "when t1.department_code = 3030 then 27 "+
                                  "when t1.department_code = 3040 then 28 "+
                                  "when t1.department_code = 3050 then 29 "+
                                  "when t1.department_code = 3060 then 30 "+
                                  "when t1.department_code = 3080 then 31 "+
                                  "when t1.department_code = 6540 then 32 "+
                                  "when t1.department_code = 6550 then 33 "+
                                  "when t1.department_code = 6560 then 34 "+
                                  "when t1.department_code = 7550 then 35 "+
                                  "when t1.department_code = 9510 then 36 "+
                                  "when t1.department_code = 9520 then 37 "+
                                  "when t1.department_code = 9530 then 38 "+
                                  "when t1.department_code = 8010 then 39 "+
                                  "when t1.department_code = 8020 then 40 "+
                                  "when t1.department_code = 8030 then 41 "+
                                  "when t1.department_code = 1510 then 42 "+
                                  "when t1.department_code = 1520 then 43 "+
                                  "when t1.department_code = 1530 then 44 "+
                                  "when t1.department_code = 1540 then 45 "+
                                  "when t1.department_code = 1010 then 46 "+
                                  "when t1.department_code = 1020 then 47 "+
                                  "when t1.department_code = 1030 then 48 "+
                                  "when t1.department_code = 1040 then 49 "+
                                  "when t1.department_code = 1050 then 50 "+
                                  "when t1.department_code = 1060 then 51 "+
                                  "when t1.department_code = 2020 then 52 "+
                                  "when t1.department_code = 2030 then 53 "+
                                  "when t1.department_code = 2040 then 54 "+
                                  "when t1.department_code = 2050 then 55 "+
                                  "when t1.department_code = 5520 then 56 "+
                                  "when t1.department_code = 5530 then 57 "+
                                  "when t1.department_code = 8040 then 58 "+
                                  "when t1.department_code = 6010 then 59 "+
                                  "when t1.department_code = 8510 then 60 "+
                                  "when t1.department_code = 8520 then 61 "+
                                  "when t1.department_code = 8530 then 62 "+
                                  "when t1.department_code = 8540 then 63 "+
                                  "when t1.department_code = 8550 then 64 "+
                                  "when t1.department_code = 8560 then 65 "+
                                  "when t1.department_code = 8570 then 66 "+
                                  "when t1.department_code = 8590 then 67 "+
                                  "when t1.department_code = 9010 then 68 "+
                                  "when t1.department_code = 9020 then 69 "+
                                  "when t1.department_code = 9030 then 70 end";
                
                pStatement = con.prepareStatement(query);
                rs = pStatement.executeQuery();

                
                 int row = 9;

                
                while(rs.next()){
                    if(row == 14 || row == 18 || row == 22 || row == 38 || row == 46 || row == 50 || row == 52){
                        row++;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(13);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 30 || row == 56 ||   row == 68 || row == 76 || row == 82 || row == 89 || row == 92 || row == 86 ){
                        row = row + 2;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(13);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 61 || row == 102 ){
                        row = row + 3;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(13);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else{
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(13);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    row ++;
                }
     
            }
            catch(Exception e){
                System.out.println(e);
            }
        } 
    private void importsTotalSalesYTD() throws SQLException{
 
       
        System.out.println("Fetching Import Sales(YTD)..");
        
           try{
                String query = "Select t1.department_code as DEPT_CODE, t2.NET_SALES from " +
                            "(select distinct department_code from dim_product where department_code in (3510,3520,3550,"+
                            "3560,3590,4510,4520,4540,5020,5030,5040,7010,7020,7030,7040,7050,7070," +
                            "7080,2510,2530,2540,2550,2560,2570,3010,3020,3030,3040,3050,3060,3080,6540,"+
                            "6550,6560,7550,9510,9520,9530,8010,8020,8030,1510,1520,1530,1540,1010,"+
                            "1020,1030,1040,1050,1060,2020,2030,2040,2050,5520,5530,8040,6010,"+
                            "8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030)) t1 "+
                            "left join (SELECT dp.department_code as DEPT_CODE, " +
                            "sum(NVL(ag.SALE_NET_VAL,0)-NVL(ag.SALE_TOT_TAX_VAL,0)-NVL(SALE_TOT_DISC_VAL,0)) "+
                            "as NET_SALES "+
                            "FROM agg_dly_str_prod ag "+
                            "LEFT JOIN dim_date d on ag.date_key = d.date_key "+
                            "LEFT JOIN dim_product dp on ag.product_code = dp.product_code "+
                            "Left join uda_item_lov uda on ag.product_code = uda.item "+
                            "where dp.department_code in (3510,3520,3550,3560,3590,4510,4520,4540,5020,5030,5040"+
                            ",7010,7020,7030,7040,7050,7070,7080,2510,2530,2540,2550,2560,2570,3010,3020,3030"+
                            ",3040,3050,3060,3080,6540,6550,6560,7550,9510,9520,9530,8010,8020,8030"+
                            ",1510,1520,1530,1540,1010,1020,1030,1040,1050,1060,2020,2030,2040,2050,5520"+
                            ",5530,8040,6010,8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030) "+
                            "and d.date_fld between trunc(sysdate-1,'YEAR') AND sysdate-1 "+
                            "and dp.latest = 1 and dp.active = 1 and dp.concession_flg = 'N' "+
                            "and uda_id = 1204 "+
                            "GROUP BY dp.department_code) "+
                            "t2 on t1.department_code = t2.DEPT_CODE "+
                            "order by case " +
                                  "when t1.department_code = 3510 then 1 "+
                                  "when t1.department_code = 3520 then 2 "+
                                  "when t1.department_code = 3550 then 3 "+
                                  "when t1.department_code = 3560 then 4 "+
                                  "when t1.department_code = 3590 then 5 "+
                                  "when t1.department_code = 4510 then 6 "+
                                  "when t1.department_code = 4520 then 7 "+
                                  "when t1.department_code = 4540 then 8 "+
                                  "when t1.department_code = 5020 then 9 "+
                                  "when t1.department_code = 5030 then 10 "+
                                  "when t1.department_code = 5040 then 11 "+
                                  "when t1.department_code = 7010 then 12 "+
                                  "when t1.department_code = 7020 then 13 "+
                                  "when t1.department_code = 7030 then 14 "+
                                  "when t1.department_code = 7040 then 15 "+
                                  "when t1.department_code = 7050 then 16 "+
                                  "when t1.department_code = 7070 then 17 "+
                                  "when t1.department_code = 7080 then 18 "+
                                  "when t1.department_code = 2510 then 19 "+
                                  "when t1.department_code = 2530 then 20 "+
                                  "when t1.department_code = 2540 then 21 "+
                                  "when t1.department_code = 2550 then 22 "+
                                  "when t1.department_code = 2560 then 23 "+
                                  "when t1.department_code = 2570 then 24 "+
                                  "when t1.department_code = 3010 then 25 "+
                                  "when t1.department_code = 3020 then 26 "+
                                  "when t1.department_code = 3030 then 27 "+
                                  "when t1.department_code = 3040 then 28 "+
                                  "when t1.department_code = 3050 then 29 "+
                                  "when t1.department_code = 3060 then 30 "+
                                  "when t1.department_code = 3080 then 31 "+
                                  "when t1.department_code = 6540 then 32 "+
                                  "when t1.department_code = 6550 then 33 "+
                                  "when t1.department_code = 6560 then 34 "+
                                  "when t1.department_code = 7550 then 35 "+
                                  "when t1.department_code = 9510 then 36 "+
                                  "when t1.department_code = 9520 then 37 "+
                                  "when t1.department_code = 9530 then 38 "+
                                  "when t1.department_code = 8010 then 39 "+
                                  "when t1.department_code = 8020 then 40 "+
                                  "when t1.department_code = 8030 then 41 "+
                                  "when t1.department_code = 1510 then 42 "+
                                  "when t1.department_code = 1520 then 43 "+
                                  "when t1.department_code = 1530 then 44 "+
                                  "when t1.department_code = 1540 then 45 "+
                                  "when t1.department_code = 1010 then 46 "+
                                  "when t1.department_code = 1020 then 47 "+
                                  "when t1.department_code = 1030 then 48 "+
                                  "when t1.department_code = 1040 then 49 "+
                                  "when t1.department_code = 1050 then 50 "+
                                  "when t1.department_code = 1060 then 51 "+
                                  "when t1.department_code = 2020 then 52 "+
                                  "when t1.department_code = 2030 then 53 "+
                                  "when t1.department_code = 2040 then 54 "+
                                  "when t1.department_code = 2050 then 55 "+
                                  "when t1.department_code = 5520 then 56 "+
                                  "when t1.department_code = 5530 then 57 "+
                                  "when t1.department_code = 8040 then 58 "+
                                  "when t1.department_code = 6010 then 59 "+
                                  "when t1.department_code = 8510 then 60 "+
                                  "when t1.department_code = 8520 then 61 "+
                                  "when t1.department_code = 8530 then 62 "+
                                  "when t1.department_code = 8540 then 63 "+
                                  "when t1.department_code = 8550 then 64 "+
                                  "when t1.department_code = 8560 then 65 "+
                                  "when t1.department_code = 8570 then 66 "+
                                  "when t1.department_code = 8590 then 67 "+
                                  "when t1.department_code = 9010 then 68 "+
                                  "when t1.department_code = 9020 then 69 "+
                                  "when t1.department_code = 9030 then 70 end";
                
                pStatement = con.prepareStatement(query);
                rs = pStatement.executeQuery();

                
                 int row = 9;

                
                while(rs.next()){
                    if(row == 14 || row == 18 || row == 22 || row == 38 || row == 46 || row == 50 || row == 52){
                        row++;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(20);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 30 || row == 56 ||   row == 68 || row == 76 || row == 82 || row == 89 || row == 92 || row == 86 ){
                        row = row + 2;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(20);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 61 || row == 102 ){
                        row = row + 3;
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(20);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else{
                        double sales = rs.getDouble("NET_SALES");
                        cell = sheet.getRow(row).createCell(20);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    row ++;
                }
     
            }
            catch(Exception e){
                System.out.println(e);
            }
        }
    
    
    private void budgetTotalSalesMTD() throws SQLException{
 
       
        System.out.println("Fetching Budget Sales(MTD)..");
        
           try{
                String query = "Select t1.department_code as DEPT_CODE, t2.TARGET_SALES from "+
                            "(select distinct department_code from dim_product where department_code in (3510,3520,3550,"+
                            "3560,3590,4510,4520,4540,5020,5030,5040,7010,7020,7030,7040,7050,7070,"+
                            "7080,2510,2530,2540,2550,2560,2570,3010,3020,3030,3040,3050,3060,3080,6540,"+
                            "6550,6560,7550,9510,9520,9530,8010,8020,8030,1510,1520,1530,1540,1010,"+
                            "1020,1030,1040,1050,1060,2020,2030,2040,2050,5520,5530,8040,6010,"+
                            "8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030)) t1 LEFT JOIN "+
                            "(SELECT f.department_code as DEPT_CODE,sum(NVL(f.target_sale_val,0)) "+
                            "as TARGET_SALES FROM fct_target f LEFT JOIN dim_date d on f.date_key = d.date_key "+
                            "where f.department_code in (3510,3520,3550,3560,3590,4510,4520,4540,5020,5030,5040"+
                            ",7010,7020,7030,7040,7050,7070,7080,2510,2530,2540,2550,2560,2570,3010,3020,3030"+
                            ",3040,3050,3060,3080,6540,6550,6560,7550,9510,9520,9530,8010,8020,8030"+
                            ",1510,1520,1530,1540,1010,1020,1030,1040,1050,1060,2020,2030,2040,2050,5520"+
                            ",5530,8040,6010,8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030) "+
                            "and d.date_fld between trunc(sysdate-1,'MM') AND sysdate-1 "+
                            "GROUP BY f.department_code) t2 ON t1.department_code = t2.DEPT_CODE "+
                            "order by case "+
                                "when t1.department_code = 3510 then 1 "+
                                "when t1.department_code = 3520 then 2 "+
                                "when t1.department_code = 3550 then 3 "+
                                "when t1.department_code = 3560 then 4 "+
                                "when t1.department_code = 3590 then 5 "+
                                "when t1.department_code = 4510 then 6 "+
                                "when t1.department_code = 4520 then 7 "+
                                "when t1.department_code = 4540 then 8 "+
                                "when t1.department_code = 5020 then 9 "+
                                "when t1.department_code = 5030 then 10 "+
                                "when t1.department_code = 5040 then 11 "+
                                "when t1.department_code = 7010 then 12 "+
                                "when t1.department_code = 7020 then 13 "+
                                "when t1.department_code = 7030 then 14 "+
                                "when t1.department_code = 7040 then 15 "+
                                "when t1.department_code = 7050 then 16 "+
                                "when t1.department_code = 7070 then 17 "+
                                "when t1.department_code = 7080 then 18 "+
                                "when t1.department_code = 2510 then 19 "+
                                "when t1.department_code = 2530 then 20 "+
                                "when t1.department_code = 2540 then 21 "+
                                "when t1.department_code = 2550 then 22 "+
                                "when t1.department_code = 2560 then 23 "+
                                "when t1.department_code = 2570 then 24 "+
                                "when t1.department_code = 3010 then 25 "+
                                "when t1.department_code = 3020 then 26 "+
                                "when t1.department_code = 3030 then 27 "+
                                "when t1.department_code = 3040 then 28 "+
                                "when t1.department_code = 3050 then 29 "+
                                "when t1.department_code = 3060 then 30 "+
                                "when t1.department_code = 3080 then 31 "+
                                "when t1.department_code = 6540 then 32 "+
                                "when t1.department_code = 6550 then 33 "+
                                "when t1.department_code = 6560 then 34 "+
                                "when t1.department_code = 7550 then 35 "+
                                "when t1.department_code = 9510 then 36 "+
                                "when t1.department_code = 9520 then 37 "+
                                "when t1.department_code = 9530 then 38 "+
                                "when t1.department_code = 8010 then 39 "+
                                "when t1.department_code = 8020 then 40 "+
                                "when t1.department_code = 8030 then 41 "+
                                "when t1.department_code = 1510 then 42 "+
                                "when t1.department_code = 1520 then 43 "+
                                "when t1.department_code = 1530 then 44 "+
                                "when t1.department_code = 1540 then 45 "+
                                "when t1.department_code = 1010 then 46 "+
                                "when t1.department_code = 1020 then 47 "+
                                "when t1.department_code = 1030 then 48 "+
                                "when t1.department_code = 1040 then 49 "+
                                "when t1.department_code = 1050 then 50 "+
                                "when t1.department_code = 1060 then 51 "+
                                "when t1.department_code = 2020 then 52 "+
                                "when t1.department_code = 2030 then 53 "+
                                "when t1.department_code = 2040 then 54 "+
                                "when t1.department_code = 2050 then 55 "+
                                "when t1.department_code = 5520 then 56 "+
                                "when t1.department_code = 5530 then 57 "+
                                "when t1.department_code = 8040 then 58 "+
                                "when t1.department_code = 6010 then 59 "+
                                "when t1.department_code = 8510 then 60 "+
                                "when t1.department_code = 8520 then 61 "+
                                "when t1.department_code = 8530 then 62 "+
                                "when t1.department_code = 8540 then 63 "+
                                "when t1.department_code = 8550 then 64 "+
                                "when t1.department_code = 8560 then 65 "+
                                "when t1.department_code = 8570 then 66 "+
                                "when t1.department_code = 8590 then 67 "+
                                "when t1.department_code = 9010 then 68 "+
                                "when t1.department_code = 9020 then 69 "+
                                "when t1.department_code = 9030 then 70 end ";
                
                pStatement = con.prepareStatement(query);
                rs = pStatement.executeQuery();

                
                int row = 9;

                
                while(rs.next()){
                    if(row == 14 || row == 18 || row == 22 || row == 38 || row == 46 || row == 50 || row == 52){
                        row++;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(8);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 30 || row == 56 ||   row == 68 || row == 76 || row == 82 || row == 89 || row == 92 || row == 86 ){
                        row = row + 2;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(8);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 61 || row == 102 ){
                        row = row + 3;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(8);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else{
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(8);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    row ++;
                }
     
            }
            catch(Exception e){
                System.out.println(e);
            }
        }
    private void budgetTotalSalesQTD() throws SQLException{
 
       
        System.out.println("Fetching Budget Sales(QTD)..");
        
           try{
                String query = "Select t1.department_code as DEPT_CODE, t2.TARGET_SALES from "+
                            "(select distinct department_code from dim_product where department_code in (3510,3520,3550,"+
                            "3560,3590,4510,4520,4540,5020,5030,5040,7010,7020,7030,7040,7050,7070,"+
                            "7080,2510,2530,2540,2550,2560,2570,3010,3020,3030,3040,3050,3060,3080,6540,"+
                            "6550,6560,7550,9510,9520,9530,8010,8020,8030,1510,1520,1530,1540,1010,"+
                            "1020,1030,1040,1050,1060,2020,2030,2040,2050,5520,5530,8040,6010,"+
                            "8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030)) t1 LEFT JOIN "+
                            "(SELECT f.department_code as DEPT_CODE,sum(NVL(f.target_sale_val,0)) "+
                            "as TARGET_SALES FROM fct_target f LEFT JOIN dim_date d on f.date_key = d.date_key "+
                            "where f.department_code in (3510,3520,3550,3560,3590,4510,4520,4540,5020,5030,5040"+
                            ",7010,7020,7030,7040,7050,7070,7080,2510,2530,2540,2550,2560,2570,3010,3020,3030"+
                            ",3040,3050,3060,3080,6540,6550,6560,7550,9510,9520,9530,8010,8020,8030"+
                            ",1510,1520,1530,1540,1010,1020,1030,1040,1050,1060,2020,2030,2040,2050,5520"+
                            ",5530,8040,6010,8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030) "+
                            "and d.date_fld between trunc(sysdate-1,'Q') AND sysdate-1 "+
                            "GROUP BY f.department_code) t2 ON t1.department_code = t2.DEPT_CODE "+
                            "order by case "+
                                "when t1.department_code = 3510 then 1 "+
                                "when t1.department_code = 3520 then 2 "+
                                "when t1.department_code = 3550 then 3 "+
                                "when t1.department_code = 3560 then 4 "+
                                "when t1.department_code = 3590 then 5 "+
                                "when t1.department_code = 4510 then 6 "+
                                "when t1.department_code = 4520 then 7 "+
                                "when t1.department_code = 4540 then 8 "+
                                "when t1.department_code = 5020 then 9 "+
                                "when t1.department_code = 5030 then 10 "+
                                "when t1.department_code = 5040 then 11 "+
                                "when t1.department_code = 7010 then 12 "+
                                "when t1.department_code = 7020 then 13 "+
                                "when t1.department_code = 7030 then 14 "+
                                "when t1.department_code = 7040 then 15 "+
                                "when t1.department_code = 7050 then 16 "+
                                "when t1.department_code = 7070 then 17 "+
                                "when t1.department_code = 7080 then 18 "+
                                "when t1.department_code = 2510 then 19 "+
                                "when t1.department_code = 2530 then 20 "+
                                "when t1.department_code = 2540 then 21 "+
                                "when t1.department_code = 2550 then 22 "+
                                "when t1.department_code = 2560 then 23 "+
                                "when t1.department_code = 2570 then 24 "+
                                "when t1.department_code = 3010 then 25 "+
                                "when t1.department_code = 3020 then 26 "+
                                "when t1.department_code = 3030 then 27 "+
                                "when t1.department_code = 3040 then 28 "+
                                "when t1.department_code = 3050 then 29 "+
                                "when t1.department_code = 3060 then 30 "+
                                "when t1.department_code = 3080 then 31 "+
                                "when t1.department_code = 6540 then 32 "+
                                "when t1.department_code = 6550 then 33 "+
                                "when t1.department_code = 6560 then 34 "+
                                "when t1.department_code = 7550 then 35 "+
                                "when t1.department_code = 9510 then 36 "+
                                "when t1.department_code = 9520 then 37 "+
                                "when t1.department_code = 9530 then 38 "+
                                "when t1.department_code = 8010 then 39 "+
                                "when t1.department_code = 8020 then 40 "+
                                "when t1.department_code = 8030 then 41 "+
                                "when t1.department_code = 1510 then 42 "+
                                "when t1.department_code = 1520 then 43 "+
                                "when t1.department_code = 1530 then 44 "+
                                "when t1.department_code = 1540 then 45 "+
                                "when t1.department_code = 1010 then 46 "+
                                "when t1.department_code = 1020 then 47 "+
                                "when t1.department_code = 1030 then 48 "+
                                "when t1.department_code = 1040 then 49 "+
                                "when t1.department_code = 1050 then 50 "+
                                "when t1.department_code = 1060 then 51 "+
                                "when t1.department_code = 2020 then 52 "+
                                "when t1.department_code = 2030 then 53 "+
                                "when t1.department_code = 2040 then 54 "+
                                "when t1.department_code = 2050 then 55 "+
                                "when t1.department_code = 5520 then 56 "+
                                "when t1.department_code = 5530 then 57 "+
                                "when t1.department_code = 8040 then 58 "+
                                "when t1.department_code = 6010 then 59 "+
                                "when t1.department_code = 8510 then 60 "+
                                "when t1.department_code = 8520 then 61 "+
                                "when t1.department_code = 8530 then 62 "+
                                "when t1.department_code = 8540 then 63 "+
                                "when t1.department_code = 8550 then 64 "+
                                "when t1.department_code = 8560 then 65 "+
                                "when t1.department_code = 8570 then 66 "+
                                "when t1.department_code = 8590 then 67 "+
                                "when t1.department_code = 9010 then 68 "+
                                "when t1.department_code = 9020 then 69 "+
                                "when t1.department_code = 9030 then 70 end ";
                
                pStatement = con.prepareStatement(query);
                rs = pStatement.executeQuery();

                
                int row = 9;

                
                while(rs.next()){
                    if(row == 14 || row == 18 || row == 22 || row == 38 || row == 46 || row == 50 || row == 52){
                        row++;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(15);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 30 || row == 56 ||   row == 68 || row == 76 || row == 82 || row == 89 || row == 92 || row == 86 ){
                        row = row + 2;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(15);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 61 || row == 102 ){
                        row = row + 3;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(15);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else{
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(15);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    row ++;
                }
     
            }
            catch(Exception e){
                System.out.println(e);
            }
        }
    private void budgetTotalSalesYTD() throws SQLException{
 
       
        System.out.println("Fetching Budget Sales(QTD)..");
        
           try{
                String query = "Select t1.department_code as DEPT_CODE, t2.TARGET_SALES from "+
                            "(select distinct department_code from dim_product where department_code in (3510,3520,3550,"+
                            "3560,3590,4510,4520,4540,5020,5030,5040,7010,7020,7030,7040,7050,7070,"+
                            "7080,2510,2530,2540,2550,2560,2570,3010,3020,3030,3040,3050,3060,3080,6540,"+
                            "6550,6560,7550,9510,9520,9530,8010,8020,8030,1510,1520,1530,1540,1010,"+
                            "1020,1030,1040,1050,1060,2020,2030,2040,2050,5520,5530,8040,6010,"+
                            "8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030)) t1 LEFT JOIN "+
                            "(SELECT f.department_code as DEPT_CODE,sum(NVL(f.target_sale_val,0)) "+
                            "as TARGET_SALES FROM fct_target f LEFT JOIN dim_date d on f.date_key = d.date_key "+
                            "where f.department_code in (3510,3520,3550,3560,3590,4510,4520,4540,5020,5030,5040"+
                            ",7010,7020,7030,7040,7050,7070,7080,2510,2530,2540,2550,2560,2570,3010,3020,3030"+
                            ",3040,3050,3060,3080,6540,6550,6560,7550,9510,9520,9530,8010,8020,8030"+
                            ",1510,1520,1530,1540,1010,1020,1030,1040,1050,1060,2020,2030,2040,2050,5520"+
                            ",5530,8040,6010,8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030) "+
                            "and d.date_fld between trunc(sysdate-1,'Y') AND sysdate-1 "+
                            "GROUP BY f.department_code) t2 ON t1.department_code = t2.DEPT_CODE "+
                            "order by case "+
                                "when t1.department_code = 3510 then 1 "+
                                "when t1.department_code = 3520 then 2 "+
                                "when t1.department_code = 3550 then 3 "+
                                "when t1.department_code = 3560 then 4 "+
                                "when t1.department_code = 3590 then 5 "+
                                "when t1.department_code = 4510 then 6 "+
                                "when t1.department_code = 4520 then 7 "+
                                "when t1.department_code = 4540 then 8 "+
                                "when t1.department_code = 5020 then 9 "+
                                "when t1.department_code = 5030 then 10 "+
                                "when t1.department_code = 5040 then 11 "+
                                "when t1.department_code = 7010 then 12 "+
                                "when t1.department_code = 7020 then 13 "+
                                "when t1.department_code = 7030 then 14 "+
                                "when t1.department_code = 7040 then 15 "+
                                "when t1.department_code = 7050 then 16 "+
                                "when t1.department_code = 7070 then 17 "+
                                "when t1.department_code = 7080 then 18 "+
                                "when t1.department_code = 2510 then 19 "+
                                "when t1.department_code = 2530 then 20 "+
                                "when t1.department_code = 2540 then 21 "+
                                "when t1.department_code = 2550 then 22 "+
                                "when t1.department_code = 2560 then 23 "+
                                "when t1.department_code = 2570 then 24 "+
                                "when t1.department_code = 3010 then 25 "+
                                "when t1.department_code = 3020 then 26 "+
                                "when t1.department_code = 3030 then 27 "+
                                "when t1.department_code = 3040 then 28 "+
                                "when t1.department_code = 3050 then 29 "+
                                "when t1.department_code = 3060 then 30 "+
                                "when t1.department_code = 3080 then 31 "+
                                "when t1.department_code = 6540 then 32 "+
                                "when t1.department_code = 6550 then 33 "+
                                "when t1.department_code = 6560 then 34 "+
                                "when t1.department_code = 7550 then 35 "+
                                "when t1.department_code = 9510 then 36 "+
                                "when t1.department_code = 9520 then 37 "+
                                "when t1.department_code = 9530 then 38 "+
                                "when t1.department_code = 8010 then 39 "+
                                "when t1.department_code = 8020 then 40 "+
                                "when t1.department_code = 8030 then 41 "+
                                "when t1.department_code = 1510 then 42 "+
                                "when t1.department_code = 1520 then 43 "+
                                "when t1.department_code = 1530 then 44 "+
                                "when t1.department_code = 1540 then 45 "+
                                "when t1.department_code = 1010 then 46 "+
                                "when t1.department_code = 1020 then 47 "+
                                "when t1.department_code = 1030 then 48 "+
                                "when t1.department_code = 1040 then 49 "+
                                "when t1.department_code = 1050 then 50 "+
                                "when t1.department_code = 1060 then 51 "+
                                "when t1.department_code = 2020 then 52 "+
                                "when t1.department_code = 2030 then 53 "+
                                "when t1.department_code = 2040 then 54 "+
                                "when t1.department_code = 2050 then 55 "+
                                "when t1.department_code = 5520 then 56 "+
                                "when t1.department_code = 5530 then 57 "+
                                "when t1.department_code = 8040 then 58 "+
                                "when t1.department_code = 6010 then 59 "+
                                "when t1.department_code = 8510 then 60 "+
                                "when t1.department_code = 8520 then 61 "+
                                "when t1.department_code = 8530 then 62 "+
                                "when t1.department_code = 8540 then 63 "+
                                "when t1.department_code = 8550 then 64 "+
                                "when t1.department_code = 8560 then 65 "+
                                "when t1.department_code = 8570 then 66 "+
                                "when t1.department_code = 8590 then 67 "+
                                "when t1.department_code = 9010 then 68 "+
                                "when t1.department_code = 9020 then 69 "+
                                "when t1.department_code = 9030 then 70 end ";
                
                pStatement = con.prepareStatement(query);
                rs = pStatement.executeQuery();

                
                int row = 9;

                
                while(rs.next()){
                    if(row == 14 || row == 18 || row == 22 || row == 38 || row == 46 || row == 50 || row == 52){
                        row++;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(22);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 30 || row == 56 ||   row == 68 || row == 76 || row == 82 || row == 89 || row == 92 || row == 86 ){
                        row = row + 2;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(22);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 61 || row == 102 ){
                        row = row + 3;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(22);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else{
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(22);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    row ++;
                }
     
            }
            catch(Exception e){
                System.out.println(e);
            }
        }
    
    
    private void budgetImportsSalesMTD() throws SQLException{
 
       
        System.out.println("Fetching Budget Imports Sales(MTD)..");
        
           try{
                String query = "Select t1.department_code as DEPT_CODE, t2.TARGET_SALES from "+
                            "(select distinct department_code from dim_product where department_code in (3510,3520,3550,"+
                            "3560,3590,4510,4520,4540,5020,5030,5040,7010,7020,7030,7040,7050,7070,"+
                            "7080,2510,2530,2540,2550,2560,2570,3010,3020,3030,3040,3050,3060,3080,6540,"+
                            "6550,6560,7550,9510,9520,9530,8010,8020,8030,1510,1520,1530,1540,1010,"+
                            "1020,1030,1040,1050,1060,2020,2030,2040,2050,5520,5530,8040,6010,"+
                            "8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030)) t1 LEFT JOIN "+
                            "(SELECT f.department_code as DEPT_CODE,sum(NVL(f.target_sale_val,0)) "+
                            "as TARGET_SALES FROM fct_target_imports f LEFT JOIN dim_date d on f.date_key = d.date_key "+
                            "where f.department_code in (3510,3520,3550,3560,3590,4510,4520,4540,5020,5030,5040"+
                            ",7010,7020,7030,7040,7050,7070,7080,2510,2530,2540,2550,2560,2570,3010,3020,3030"+
                            ",3040,3050,3060,3080,6540,6550,6560,7550,9510,9520,9530,8010,8020,8030"+
                            ",1510,1520,1530,1540,1010,1020,1030,1040,1050,1060,2020,2030,2040,2050,5520"+
                            ",5530,8040,6010,8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030) "+
                            "and d.date_fld between trunc(sysdate-1,'MM') AND sysdate-1 "+
                            "GROUP BY f.department_code) t2 ON t1.department_code = t2.DEPT_CODE "+
                            "order by case "+
                                "when t1.department_code = 3510 then 1 "+
                                "when t1.department_code = 3520 then 2 "+
                                "when t1.department_code = 3550 then 3 "+
                                "when t1.department_code = 3560 then 4 "+
                                "when t1.department_code = 3590 then 5 "+
                                "when t1.department_code = 4510 then 6 "+
                                "when t1.department_code = 4520 then 7 "+
                                "when t1.department_code = 4540 then 8 "+
                                "when t1.department_code = 5020 then 9 "+
                                "when t1.department_code = 5030 then 10 "+
                                "when t1.department_code = 5040 then 11 "+
                                "when t1.department_code = 7010 then 12 "+
                                "when t1.department_code = 7020 then 13 "+
                                "when t1.department_code = 7030 then 14 "+
                                "when t1.department_code = 7040 then 15 "+
                                "when t1.department_code = 7050 then 16 "+
                                "when t1.department_code = 7070 then 17 "+
                                "when t1.department_code = 7080 then 18 "+
                                "when t1.department_code = 2510 then 19 "+
                                "when t1.department_code = 2530 then 20 "+
                                "when t1.department_code = 2540 then 21 "+
                                "when t1.department_code = 2550 then 22 "+
                                "when t1.department_code = 2560 then 23 "+
                                "when t1.department_code = 2570 then 24 "+
                                "when t1.department_code = 3010 then 25 "+
                                "when t1.department_code = 3020 then 26 "+
                                "when t1.department_code = 3030 then 27 "+
                                "when t1.department_code = 3040 then 28 "+
                                "when t1.department_code = 3050 then 29 "+
                                "when t1.department_code = 3060 then 30 "+
                                "when t1.department_code = 3080 then 31 "+
                                "when t1.department_code = 6540 then 32 "+
                                "when t1.department_code = 6550 then 33 "+
                                "when t1.department_code = 6560 then 34 "+
                                "when t1.department_code = 7550 then 35 "+
                                "when t1.department_code = 9510 then 36 "+
                                "when t1.department_code = 9520 then 37 "+
                                "when t1.department_code = 9530 then 38 "+
                                "when t1.department_code = 8010 then 39 "+
                                "when t1.department_code = 8020 then 40 "+
                                "when t1.department_code = 8030 then 41 "+
                                "when t1.department_code = 1510 then 42 "+
                                "when t1.department_code = 1520 then 43 "+
                                "when t1.department_code = 1530 then 44 "+
                                "when t1.department_code = 1540 then 45 "+
                                "when t1.department_code = 1010 then 46 "+
                                "when t1.department_code = 1020 then 47 "+
                                "when t1.department_code = 1030 then 48 "+
                                "when t1.department_code = 1040 then 49 "+
                                "when t1.department_code = 1050 then 50 "+
                                "when t1.department_code = 1060 then 51 "+
                                "when t1.department_code = 2020 then 52 "+
                                "when t1.department_code = 2030 then 53 "+
                                "when t1.department_code = 2040 then 54 "+
                                "when t1.department_code = 2050 then 55 "+
                                "when t1.department_code = 5520 then 56 "+
                                "when t1.department_code = 5530 then 57 "+
                                "when t1.department_code = 8040 then 58 "+
                                "when t1.department_code = 6010 then 59 "+
                                "when t1.department_code = 8510 then 60 "+
                                "when t1.department_code = 8520 then 61 "+
                                "when t1.department_code = 8530 then 62 "+
                                "when t1.department_code = 8540 then 63 "+
                                "when t1.department_code = 8550 then 64 "+
                                "when t1.department_code = 8560 then 65 "+
                                "when t1.department_code = 8570 then 66 "+
                                "when t1.department_code = 8590 then 67 "+
                                "when t1.department_code = 9010 then 68 "+
                                "when t1.department_code = 9020 then 69 "+
                                "when t1.department_code = 9030 then 70 end ";
                
                pStatement = con.prepareStatement(query);
                rs = pStatement.executeQuery();

                
                int row = 9;

                
                while(rs.next()){
                    if(row == 14 || row == 18 || row == 22 || row == 38 || row == 46 || row == 50 || row == 52){
                        row++;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(9);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 30 || row == 56 ||   row == 68 || row == 76 || row == 82 || row == 89 || row == 92 || row == 86 ){
                        row = row + 2;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(9);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 61 || row == 102 ){
                        row = row + 3;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(9);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else{
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(9);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    row ++;
                }
     
            }
            catch(Exception e){
                System.out.println(e);
            }
        }
    private void budgetImportsSalesQTD() throws SQLException{
 
       
        System.out.println("Fetching Budget Imports Sales(QTD)..");
        
           try{
                String query = "Select t1.department_code as DEPT_CODE, t2.TARGET_SALES from "+
                            "(select distinct department_code from dim_product where department_code in (3510,3520,3550,"+
                            "3560,3590,4510,4520,4540,5020,5030,5040,7010,7020,7030,7040,7050,7070,"+
                            "7080,2510,2530,2540,2550,2560,2570,3010,3020,3030,3040,3050,3060,3080,6540,"+
                            "6550,6560,7550,9510,9520,9530,8010,8020,8030,1510,1520,1530,1540,1010,"+
                            "1020,1030,1040,1050,1060,2020,2030,2040,2050,5520,5530,8040,6010,"+
                            "8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030)) t1 LEFT JOIN "+
                            "(SELECT f.department_code as DEPT_CODE,sum(NVL(f.target_sale_val,0)) "+
                            "as TARGET_SALES FROM fct_target_imports f LEFT JOIN dim_date d on f.date_key = d.date_key "+
                            "where f.department_code in (3510,3520,3550,3560,3590,4510,4520,4540,5020,5030,5040"+
                            ",7010,7020,7030,7040,7050,7070,7080,2510,2530,2540,2550,2560,2570,3010,3020,3030"+
                            ",3040,3050,3060,3080,6540,6550,6560,7550,9510,9520,9530,8010,8020,8030"+
                            ",1510,1520,1530,1540,1010,1020,1030,1040,1050,1060,2020,2030,2040,2050,5520"+
                            ",5530,8040,6010,8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030) "+
                            "and d.date_fld between trunc(sysdate-1,'Q') AND sysdate-1 "+
                            "GROUP BY f.department_code) t2 ON t1.department_code = t2.DEPT_CODE "+
                            "order by case "+
                                "when t1.department_code = 3510 then 1 "+
                                "when t1.department_code = 3520 then 2 "+
                                "when t1.department_code = 3550 then 3 "+
                                "when t1.department_code = 3560 then 4 "+
                                "when t1.department_code = 3590 then 5 "+
                                "when t1.department_code = 4510 then 6 "+
                                "when t1.department_code = 4520 then 7 "+
                                "when t1.department_code = 4540 then 8 "+
                                "when t1.department_code = 5020 then 9 "+
                                "when t1.department_code = 5030 then 10 "+
                                "when t1.department_code = 5040 then 11 "+
                                "when t1.department_code = 7010 then 12 "+
                                "when t1.department_code = 7020 then 13 "+
                                "when t1.department_code = 7030 then 14 "+
                                "when t1.department_code = 7040 then 15 "+
                                "when t1.department_code = 7050 then 16 "+
                                "when t1.department_code = 7070 then 17 "+
                                "when t1.department_code = 7080 then 18 "+
                                "when t1.department_code = 2510 then 19 "+
                                "when t1.department_code = 2530 then 20 "+
                                "when t1.department_code = 2540 then 21 "+
                                "when t1.department_code = 2550 then 22 "+
                                "when t1.department_code = 2560 then 23 "+
                                "when t1.department_code = 2570 then 24 "+
                                "when t1.department_code = 3010 then 25 "+
                                "when t1.department_code = 3020 then 26 "+
                                "when t1.department_code = 3030 then 27 "+
                                "when t1.department_code = 3040 then 28 "+
                                "when t1.department_code = 3050 then 29 "+
                                "when t1.department_code = 3060 then 30 "+
                                "when t1.department_code = 3080 then 31 "+
                                "when t1.department_code = 6540 then 32 "+
                                "when t1.department_code = 6550 then 33 "+
                                "when t1.department_code = 6560 then 34 "+
                                "when t1.department_code = 7550 then 35 "+
                                "when t1.department_code = 9510 then 36 "+
                                "when t1.department_code = 9520 then 37 "+
                                "when t1.department_code = 9530 then 38 "+
                                "when t1.department_code = 8010 then 39 "+
                                "when t1.department_code = 8020 then 40 "+
                                "when t1.department_code = 8030 then 41 "+
                                "when t1.department_code = 1510 then 42 "+
                                "when t1.department_code = 1520 then 43 "+
                                "when t1.department_code = 1530 then 44 "+
                                "when t1.department_code = 1540 then 45 "+
                                "when t1.department_code = 1010 then 46 "+
                                "when t1.department_code = 1020 then 47 "+
                                "when t1.department_code = 1030 then 48 "+
                                "when t1.department_code = 1040 then 49 "+
                                "when t1.department_code = 1050 then 50 "+
                                "when t1.department_code = 1060 then 51 "+
                                "when t1.department_code = 2020 then 52 "+
                                "when t1.department_code = 2030 then 53 "+
                                "when t1.department_code = 2040 then 54 "+
                                "when t1.department_code = 2050 then 55 "+
                                "when t1.department_code = 5520 then 56 "+
                                "when t1.department_code = 5530 then 57 "+
                                "when t1.department_code = 8040 then 58 "+
                                "when t1.department_code = 6010 then 59 "+
                                "when t1.department_code = 8510 then 60 "+
                                "when t1.department_code = 8520 then 61 "+
                                "when t1.department_code = 8530 then 62 "+
                                "when t1.department_code = 8540 then 63 "+
                                "when t1.department_code = 8550 then 64 "+
                                "when t1.department_code = 8560 then 65 "+
                                "when t1.department_code = 8570 then 66 "+
                                "when t1.department_code = 8590 then 67 "+
                                "when t1.department_code = 9010 then 68 "+
                                "when t1.department_code = 9020 then 69 "+
                                "when t1.department_code = 9030 then 70 end ";
                
                pStatement = con.prepareStatement(query);
                rs = pStatement.executeQuery();

                
                int row = 9;

                
                while(rs.next()){
                    if(row == 14 || row == 18 || row == 22 || row == 38 || row == 46 || row == 50 || row == 52){
                        row++;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(16);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 30 || row == 56 ||   row == 68 || row == 76 || row == 82 || row == 89 || row == 92 || row == 86 ){
                        row = row + 2;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(16);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 61 || row == 102 ){
                        row = row + 3;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(16);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else{
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(16);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    row ++;
                }
     
            }
            catch(Exception e){
                System.out.println(e);
            }
        }
    private void budgetImportsSalesYTD() throws SQLException{
 
       
        System.out.println("Fetching Budget Imports Sales(YTD)..");
        
           try{
                String query = "Select t1.department_code as DEPT_CODE, t2.TARGET_SALES from "+
                            "(select distinct department_code from dim_product where department_code in (3510,3520,3550,"+
                            "3560,3590,4510,4520,4540,5020,5030,5040,7010,7020,7030,7040,7050,7070,"+
                            "7080,2510,2530,2540,2550,2560,2570,3010,3020,3030,3040,3050,3060,3080,6540,"+
                            "6550,6560,7550,9510,9520,9530,8010,8020,8030,1510,1520,1530,1540,1010,"+
                            "1020,1030,1040,1050,1060,2020,2030,2040,2050,5520,5530,8040,6010,"+
                            "8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030)) t1 LEFT JOIN "+
                            "(SELECT f.department_code as DEPT_CODE,sum(NVL(f.target_sale_val,0)) "+
                            "as TARGET_SALES FROM fct_target_imports f LEFT JOIN dim_date d on f.date_key = d.date_key "+
                            "where f.department_code in (3510,3520,3550,3560,3590,4510,4520,4540,5020,5030,5040"+
                            ",7010,7020,7030,7040,7050,7070,7080,2510,2530,2540,2550,2560,2570,3010,3020,3030"+
                            ",3040,3050,3060,3080,6540,6550,6560,7550,9510,9520,9530,8010,8020,8030"+
                            ",1510,1520,1530,1540,1010,1020,1030,1040,1050,1060,2020,2030,2040,2050,5520"+
                            ",5530,8040,6010,8510,8520,8530,8540,8550,8560,8570,8590,9010,9020,9030) "+
                            "and d.date_fld between trunc(sysdate-1,'Y') AND sysdate-1 "+
                            "GROUP BY f.department_code) t2 ON t1.department_code = t2.DEPT_CODE "+
                            "order by case "+
                                "when t1.department_code = 3510 then 1 "+
                                "when t1.department_code = 3520 then 2 "+
                                "when t1.department_code = 3550 then 3 "+
                                "when t1.department_code = 3560 then 4 "+
                                "when t1.department_code = 3590 then 5 "+
                                "when t1.department_code = 4510 then 6 "+
                                "when t1.department_code = 4520 then 7 "+
                                "when t1.department_code = 4540 then 8 "+
                                "when t1.department_code = 5020 then 9 "+
                                "when t1.department_code = 5030 then 10 "+
                                "when t1.department_code = 5040 then 11 "+
                                "when t1.department_code = 7010 then 12 "+
                                "when t1.department_code = 7020 then 13 "+
                                "when t1.department_code = 7030 then 14 "+
                                "when t1.department_code = 7040 then 15 "+
                                "when t1.department_code = 7050 then 16 "+
                                "when t1.department_code = 7070 then 17 "+
                                "when t1.department_code = 7080 then 18 "+
                                "when t1.department_code = 2510 then 19 "+
                                "when t1.department_code = 2530 then 20 "+
                                "when t1.department_code = 2540 then 21 "+
                                "when t1.department_code = 2550 then 22 "+
                                "when t1.department_code = 2560 then 23 "+
                                "when t1.department_code = 2570 then 24 "+
                                "when t1.department_code = 3010 then 25 "+
                                "when t1.department_code = 3020 then 26 "+
                                "when t1.department_code = 3030 then 27 "+
                                "when t1.department_code = 3040 then 28 "+
                                "when t1.department_code = 3050 then 29 "+
                                "when t1.department_code = 3060 then 30 "+
                                "when t1.department_code = 3080 then 31 "+
                                "when t1.department_code = 6540 then 32 "+
                                "when t1.department_code = 6550 then 33 "+
                                "when t1.department_code = 6560 then 34 "+
                                "when t1.department_code = 7550 then 35 "+
                                "when t1.department_code = 9510 then 36 "+
                                "when t1.department_code = 9520 then 37 "+
                                "when t1.department_code = 9530 then 38 "+
                                "when t1.department_code = 8010 then 39 "+
                                "when t1.department_code = 8020 then 40 "+
                                "when t1.department_code = 8030 then 41 "+
                                "when t1.department_code = 1510 then 42 "+
                                "when t1.department_code = 1520 then 43 "+
                                "when t1.department_code = 1530 then 44 "+
                                "when t1.department_code = 1540 then 45 "+
                                "when t1.department_code = 1010 then 46 "+
                                "when t1.department_code = 1020 then 47 "+
                                "when t1.department_code = 1030 then 48 "+
                                "when t1.department_code = 1040 then 49 "+
                                "when t1.department_code = 1050 then 50 "+
                                "when t1.department_code = 1060 then 51 "+
                                "when t1.department_code = 2020 then 52 "+
                                "when t1.department_code = 2030 then 53 "+
                                "when t1.department_code = 2040 then 54 "+
                                "when t1.department_code = 2050 then 55 "+
                                "when t1.department_code = 5520 then 56 "+
                                "when t1.department_code = 5530 then 57 "+
                                "when t1.department_code = 8040 then 58 "+
                                "when t1.department_code = 6010 then 59 "+
                                "when t1.department_code = 8510 then 60 "+
                                "when t1.department_code = 8520 then 61 "+
                                "when t1.department_code = 8530 then 62 "+
                                "when t1.department_code = 8540 then 63 "+
                                "when t1.department_code = 8550 then 64 "+
                                "when t1.department_code = 8560 then 65 "+
                                "when t1.department_code = 8570 then 66 "+
                                "when t1.department_code = 8590 then 67 "+
                                "when t1.department_code = 9010 then 68 "+
                                "when t1.department_code = 9020 then 69 "+
                                "when t1.department_code = 9030 then 70 end ";
                
                pStatement = con.prepareStatement(query);
                rs = pStatement.executeQuery();

                
                int row = 9;

                
                while(rs.next()){
                    if(row == 14 || row == 18 || row == 22 || row == 38 || row == 46 || row == 50 || row == 52){
                        row++;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(23);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 30 || row == 56 ||   row == 68 || row == 76 || row == 82 || row == 89 || row == 92 || row == 86 ){
                        row = row + 2;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(23);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else if (row == 61 || row == 102 ){
                        row = row + 3;
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(23);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    else{
                        double sales = rs.getDouble("TARGET_SALES");
                        cell = sheet.getRow(row).createCell(23);
                        cell.setCellValue(sales);
                        cell.setCellStyle(style1);
                    }
                    row ++;
                }
     
            }
            catch(Exception e){
                System.out.println(e);
            }
        }
    
     private void SendEmail(){
       String from = "Report Mailer<report.mailer@metroretail.com.ph>";
       String host = "mymail.metrogaisano.com";
      
       Properties mailproperties = new Properties();
      
     
       mailproperties.put("mail.transport.protocol", "smtp");
       mailproperties.put("mail.host", host);
       mailproperties.put("mail.smtp.auth", "false");
       mailproperties.put("mail.smtp.port", "25");
       mailproperties.put("mail.smtp.ssl.enable", "false");
       mailproperties.put("mail.smtp.starttls.enable", "false");
       
 
       Session session = Session.getInstance(mailproperties);
      

       try{
       String[] to = {"eric.molina@metroretail.com.ph"};
       String[] cc = {"lloydpatrick.flores@metroretail.com.ph"};
       InternetAddress[] addressTo = new InternetAddress[to.length];
       InternetAddress[] addressCc = new InternetAddress[cc.length];
        for (int i = 0; i < to.length; i++)
            {
                addressTo[i] = new InternetAddress(to[i]);
            }
        for (int i = 0; i < cc.length; i++)
            {
                addressCc[i] = new InternetAddress(cc[i]);
            }
          // Create a default MimeMessage object.
          MimeMessage message = new MimeMessage(session);
          
          // Set From: header field of the header.
         message.setFrom(new InternetAddress(from));

         // Set To: header field of the header.
         //message.addRecipient(Message.RecipientType.TO, new InternetAddress(to2));
         message.addRecipients(Message.RecipientType.TO, addressTo);
         message.addRecipients(Message.RecipientType.CC, addressCc);

         // Set Subject: header field
         message.setSubject("Imports Performance (as of "+date+")");

         // Adding HTML Measures
         BodyPart messageBodyPart = new MimeBodyPart();
         String htmlText ="<html>"+
                            "Dear Users, <br> <br>"+
                            "Please refer below for details of the enclosed report: <br> <br>"+
                            "<table border = 1>"+
                                    "<tr>"+
                                            "<th bgcolor = #FDE9D9 align='center'><b>Measure</b></th>"+
                                            "<th bgcolor = #FDE9D9 align='center'><b>Definition</b></th>"+
                                    "</tr>"+
                                    "<tr>"+
                                            "<td>Total Sales Actual</td>"+
                                            "<td>Net Sales Retail - Outright (Actual)</td>"+
                                    "</tr>"+
                                    "<tr>"+
                                            "<td>Imports Sales Actual</td>"+
                                            "<td>Net Sales Retail - Imports (Actual)</td>"+
                                    "</tr>"+
                                    "<tr>"+
                                            "<td>Imports STB Actual</td>"+
                                            "<td>Share of actual imports sales to total outright sales</td>"+
                                    "</tr>"+
                                    "<tr>"+
                                            "<td>Total Sales Budget</td>"+
                                            "<td>Net Sales Retail - Outright (Budget)</td>"+
                                    "</tr>"+
                                    "<tr>"+
                                            "<td>Imports Sales Budget</td>"+
                                            "<td>Net Sales Retail - Imports (Budget)</td>"+
                                    "</tr>"+
                                    "<tr>"+
                                            "<td>Imports STB Budget</td>"+
                                            "<td>Share of imports sales budget to total outright sales budget</td>"+
                                    "</tr>"+
                                    "<tr>"+
                                            "<td>STB Variance</td>"+
                                            "<td>Difference between the actual and budget imports share </td>"+
                                    "</tr>"+
                            "</table>"+
                            "<br>"+
                            "If you need assistance, kindly email arcbi.support&#64;metrogaisano.com. <br> <br>"+
                            "Regards, <br>"+
                            "ARC BI Support <p>"+
                            "</html>";
         messageBodyPart.setContent(htmlText, "text/html");
         Multipart multipart = new MimeMultipart();
         multipart.addBodyPart(messageBodyPart);
        
         
          // Adding attachment
         messageBodyPart = new MimeBodyPart();
         FileDataSource source = new FileDataSource("C:/Documents and Settings/Administrator/My Documents/NetBeansProjects/ImportsPerformance/output/Importation Performance (as of "+date+").xls");
         messageBodyPart.setDataHandler(new DataHandler(source));
         messageBodyPart.setFileName("Importation Performance (as of "+date+").xls");
         multipart.addBodyPart(messageBodyPart);

         // Send the complete message parts
         message.setContent(multipart);

         
         
         
         System.out.println("Sending Report..");
         // Send message
         Transport transport = session.getTransport("smtp");
         Transport.send(message);
         transport.close();
         System.out.println("Report sent successfully.");
         System.out.println("Recipients:");
         for (int i = 0; i < to.length; i++)
            {
                System.out.println(to[i]);
            }
        for (int i = 0; i < cc.length; i++)
            {
                System.out.println(cc[i]);
            }
      }catch (MessagingException mex) {
          System.out.print(mex);
      }
   }
    
    
    
}
