/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package docx;

import java.io.*;
import java.util.List;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.*;

/**
 *
 * @author huyhoang
 */
public class CreateDocx {

    public void CreateOneTable() {
        int numRow = 3; //Tháng/năm, Nghạch/bậc, Hệ số lương
        int numCol = 29; //all DataRow from DB
        int maxCol = 9; // max number of columns that fit page 

        try {
            XWPFDocument doc = new XWPFDocument(new FileInputStream("TestDoc.docx"));
            List<XWPFTable> tables = doc.getTables();
            XWPFTable old_table = tables.get(9); //position of old table
            XmlCursor cursor = old_table.getCTTbl().newCursor(); //make cursor to point at old_table
            int position = doc.getPosOfTable(old_table); //get element_number of old_table
            doc.removeBodyElement(position); //remove old_table
            XWPFTable new_table = doc.insertNewTbl(cursor); //insert new table
            new_table.setWidth("100%");

            //fill new_table
            int set = 0;
            while (numCol >= maxCol) {
                for (int rowIndex = 0 + set * numRow; rowIndex < numRow + set * numRow; rowIndex++) //create 3 row each Set
                {
                    XWPFTableRow row = new_table.getRow(rowIndex);
                    if (row == null) {
                        row = new_table.createRow();
                    }

                    for (int colIndex = 0; colIndex <= maxCol; colIndex++) //1st column not used to fill data
                    {
                        XWPFTableCell cell = row.getCell(colIndex);
                        if (cell == null) {
                            cell = row.createCell();
                        }

                        if ((rowIndex == 0 + set * numRow) && (colIndex == 0)) {
                            cell.setText("Tháng/năm");
                        } else if ((rowIndex == 1 + set * numRow) && (colIndex == 0)) {
                            cell.setText("Ngạch/bậc");
                        } else if ((rowIndex == 2 + set * numRow) && (colIndex == 0)) {
                            cell.setText("Hệ số lương");
                        } else {
                            //set Data here
                            cell.setText(String.format("data %d", set * maxCol + colIndex - 1));
                        }

                    }
                }
                set = set + 1;
                numCol -= maxCol;
            }
            //write last set
            for (int rowIndex = 0 + set * numRow; rowIndex < numRow + set * numRow; rowIndex++) {
                XWPFTableRow row = new_table.getRow(rowIndex);
                if (row == null) {
                    row = new_table.createRow();
                }
                for (int colIndex = 0; colIndex <= numCol; colIndex++) {
                    XWPFTableCell cell = row.getCell(colIndex);
                    if (cell == null) {
                        cell = row.createCell();
                    }

                    if ((rowIndex == 0 + set * numRow) && (colIndex == 0)) {
                        cell.setText("Tháng/năm");
                    } else if ((rowIndex == 1 + set * numRow) && (colIndex == 0)) {
                        cell.setText("Ngạch/bậc");
                    } else if ((rowIndex == 2 + set * numRow) && (colIndex == 0)) {
                        cell.setText("Hệ số lương");
                    } else {
                        //set Data here
                        cell.setText(String.format("data %d", set * maxCol + colIndex - 1));
                    }
                }
            }

            FileOutputStream out = new FileOutputStream(new File("TestDoc.docx"));
            doc.write(out);
            doc.close();

        } catch (FileNotFoundException e) {
            System.out.println("File not found.");
        } catch (IOException ex) {
            System.out.println("IOException while processing file");
        }

    }

    public void CreateMultiTable() {
        int numRow = 3; //Tháng/năm, Nghạch/bậc, Hệ số lương
        int numCol = 29; //all DataRow from DB
        int maxCol = 9; // max number of columns that fit page 

        try {
            XWPFDocument doc = new XWPFDocument(new FileInputStream("TestDoc.docx"));
            List<XWPFTable> tables = doc.getTables();
            XWPFTable old_table = tables.get(9); //position of old table
            XmlCursor cursor = old_table.getCTTbl().newCursor(); //make cursor to point at old_table
            int position = doc.getPosOfTable(old_table); //get element_number of old_table
            doc.removeBodyElement(position); //remove old_table

            int set = 0;
            while (numCol >= maxCol) {
                XWPFTable new_table = doc.insertNewTbl(cursor); //insert new table
                //new_table.setWidth("100%");
                for (int rowIndex = 0; rowIndex < numRow; rowIndex++) //create 3 row each Set
                {
                    XWPFTableRow row = new_table.getRow(rowIndex);
                    if (row == null) {
                        row = new_table.createRow();
                    }

                    for (int colIndex = 0; colIndex <= maxCol; colIndex++) //1st column not used to fill data
                    {
                        XWPFTableCell cell = row.getCell(colIndex);
                        if (cell == null) {
                            cell = row.createCell();
                        }

                        if ((rowIndex == 0) && (colIndex == 0)) {
                            cell.setText("Tháng/năm");
                        } else if ((rowIndex == 1) && (colIndex == 0)) {
                            cell.setText("Ngạch/bậc");
                        } else if ((rowIndex == 2) && (colIndex == 0)) {
                            cell.setText("Hệ số lương");
                        } else {
                            //set Data here
                            cell.setText(String.format("data %d", set * maxCol + colIndex - 1));
                        }

                    }
                }
                    
                //move cursor to the end of the table 
                cursor.toEndToken();
                while (cursor.hasNextToken() && cursor.toNextToken() != org.apache.xmlbeans.XmlCursor.TokenType.START);
                
                //add break line
                XWPFParagraph newParagraph = doc.insertNewParagraph(cursor);
                XWPFRun run = newParagraph.createRun();
                run.addBreak();
                cursor.toEndToken();
                while (cursor.hasNextToken() && cursor.toNextToken() != org.apache.xmlbeans.XmlCursor.TokenType.START);
                
                set = set + 1;
                numCol -= maxCol;
            }

            //last table
            XWPFTable new_table = doc.insertNewTbl(cursor); //insert new table
            //new_table.setWidth("100%");
            for (int rowIndex = 0; rowIndex < numRow; rowIndex++) {
                XWPFTableRow row = new_table.getRow(rowIndex);
                if (row == null) {
                    row = new_table.createRow();
                }
                for (int colIndex = 0; colIndex <= numCol; colIndex++) {
                    XWPFTableCell cell = row.getCell(colIndex);
                    if (cell == null) {
                        cell = row.createCell();
                    }

                    if ((rowIndex == 0) && (colIndex == 0)) {
                        cell.setText("Tháng/năm");
                    } else if ((rowIndex == 1) && (colIndex == 0)) {
                        cell.setText("Ngạch/bậc");
                    } else if ((rowIndex == 2) && (colIndex == 0)) {
                        cell.setText("Hệ số lương");
                    } else {
                        //set Data here
                        cell.setText(String.format("data %d", set * maxCol + colIndex - 1));
                    }
                }
            }

            FileOutputStream out = new FileOutputStream(new File("TestDoc.docx"));
            doc.write(out);
            doc.close();

        } catch (FileNotFoundException e) {
            System.out.println("File not found.");
        } catch (IOException ex) {
            System.out.println("IOException while processing file");
        }

    }

}
