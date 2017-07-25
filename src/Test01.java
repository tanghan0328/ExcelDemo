import java.io.File;
import java.io.FileInputStream;
import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.*;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;

import java.io.FileOutputStream;

import java.io.*;


public class Test01 {


    public static void main(String[] args) throws Exception {
        ArrayList<String> strArray = new ArrayList();
        ArrayList<String> enArray = new ArrayList();
        ArrayList<String> jaArray = new ArrayList();
        ArrayList<String> spArray = new ArrayList();
        int number = 0;
        //创建输入流
        FileInputStream fis = new FileInputStream(new File("/Users/tangtang/Desktop/1.xlsx"));
        //通过构造函数传参
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        //获取工作表
        XSSFSheet sheet = workbook.getSheetAt(0);
        //获取行,行号作为参数传递给getRow方法,第一行从0开始计算
        XSSFRow row = sheet.getRow(3);
        //总行数
        int trLength = sheet.getLastRowNum();
        //总列数
        int tdLength = row.getLastCellNum();
        for(int index =2;index<trLength;index++){
            for(int indexY=1;indexY<5;indexY++){

                XSSFRow row1 = sheet.getRow(index);
                //得到Excel工作表指定行的单元格
                XSSFCell cellColumn1  = row1.getCell(1);
                if(cellColumn1==null||cellColumn1.equals("")||cellColumn1.getCellType() ==XSSFCell.CELL_TYPE_BLANK){
                    break;
                }
                XSSFCell cell = row1.getCell(indexY);
                String cellValue= "";
                if(cell==null||cell.equals("")||cell.getCellType() ==XSSFCell.CELL_TYPE_BLANK){
                    cellValue = "";
                }else{
                    cellValue = cell.getStringCellValue();
                }
                System.out.println("值是多少===》"+cellValue);
                if(indexY == 1){
                    strArray.add(cellValue);
                }
                if(indexY == 2){
                    enArray.add(cellValue);
                }
                if(indexY == 3){
                    jaArray.add(cellValue);
                }
                if(indexY == 4){
                    spArray.add(cellValue);
                }
            }
            number++;
        }
        fis.close();

        FileOutputStream out = null;

        FileOutputStream outSTr = null;

        BufferedOutputStream Buff=null;

        FileWriter fw = null;

        int count=1000;//写文件行数
        try{
            out = new FileOutputStream(new File("/Users/tangtang/Desktop/zn.text"));
            System.out.println("======="+number);
            for (int i = 0; i < enArray.size(); i++) {
                 out.write(("\""+strArray.get(i) + "\" = \"" + enArray.get(i)+"\";\r\n").getBytes());
             }

             out = new FileOutputStream(new File("/Users/tangtang/Desktop/ja.text"));

            System.out.println("======="+number);
            for (int i = 0; i < jaArray.size(); i++) {
                out.write(("\""+strArray.get(i) + "\" = \"" + jaArray.get(i)+"\";\r\n").getBytes());
             }

             out = new FileOutputStream(new File("/Users/tangtang/Desktop/sp.text"));

             System.out.println("======="+number);
             for (int i = 0; i < spArray.size(); i++) {
                out.write(("\""+strArray.get(i) + "\" = \"" + spArray.get(i)+"\";\r\n").getBytes());
             }
            out.close();
        }catch (Exception e){

            e.printStackTrace();
        }

    }


}
