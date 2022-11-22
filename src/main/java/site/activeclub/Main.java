package site.activeclub;

import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;


public class Main {

    // 设置变量
    static HashMap<String, Integer> typeAndCount = new HashMap<>();
    static HashMap<String, Integer> nameAndCount = new HashMap<>();
    static StringBuilder stringBuilder = new StringBuilder();
    static String msg = "123";

    static Integer rolNum = 6;

    public static void main(String[] args) {

        // 读取文件路径
        String parentDirPath = "F:\\2工作记录\\";
        String dirPath = "F:\\2工作记录\\事件记录";
        List<File> fileList = getFileList(dirPath);


        // 读取文件内容
        for (File file : fileList) {
            readExcel(file);
        }

        // 数据排序
        List<KeyCounter> typeAndCountList = new ArrayList<>();
        List<KeyCounter> nameAndCountList = new ArrayList<>();

        for (Map.Entry<String, Integer> entry : typeAndCount.entrySet()) {
            String key = entry.getKey();
            Integer count = entry.getValue();
            typeAndCountList.add(new KeyCounter(key,count));
        }

        for (Map.Entry<String, Integer> entry : nameAndCount.entrySet()) {
            String key = entry.getKey();
            Integer count = entry.getValue();
            nameAndCountList.add(new KeyCounter(key,count));
        }

        Collections.sort(typeAndCountList);
        Collections.sort(nameAndCountList);


        // 文件输出
        writeExcel(typeAndCountList,parentDirPath+"typeAndCount.xls");
        writeExcel(nameAndCountList,parentDirPath+"nameAndCount.xls");

        System.out.println();

    }

    private static List<File> getFileList(String dirPath) {
        List<File> fileList = new ArrayList<>();

        File file = new File(dirPath);
        String absolutePath = file.getAbsolutePath();
        if (!file.isDirectory()) {
            fileList.add(file);
            return fileList;
        } else if (file.isDirectory()) {
            for (String childPath : file.list()) {
                List<File> childfileList = getFileList(absolutePath + "\\" +childPath);
                fileList.addAll(childfileList);
            }
        }
        return fileList;
    }

    @Data
    static class KeyCounter implements Comparable<KeyCounter>{
        private String key;
        private Integer counter;

        public KeyCounter(String key, Integer counter) {
            this.key = key;
            this.counter = counter;
        }

        @Override
        public int compareTo(KeyCounter o) {
            return o.getCounter() - this.counter;
        }
    }

    private static void writeExcel(List<KeyCounter> list,String filePath){


        Workbook wb = new HSSFWorkbook();
        int rowSize = 0;
        Sheet sheet = wb.createSheet();
        Row row = sheet.createRow(rowSize);

        // titil
        row.createCell(0).setCellValue("对象名");
        row.createCell(1).setCellValue("次数");

        try {
            for (int x = 0; x < list.size(); x++) {

                KeyCounter keyCounter = list.get(x);
                rowSize = 1;
                Row rowNew = sheet.createRow(rowSize + x);
                rowNew.createCell(0).setCellValue(keyCounter.getKey());
                rowNew.createCell(1).setCellValue(keyCounter.getCounter());
            }
        } catch (Exception e) {

        }
        OutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(filePath);
            wb.write(outputStream);
        } catch (Exception e) {

        } finally {
            try {
                if (outputStream != null) {
                    outputStream.flush();
                    outputStream.close();
                }
            } catch (Exception e) {

            }
            try{
                if(wb != null){
                    wb.close();
                }
            } catch (Exception e){

            }
        }


    }

    private static void readExcel(File file) {
        try {
            //创建工作簿
            XSSFWorkbook hssfWorkbook = new XSSFWorkbook(new FileInputStream(file));
            //获取工作簿下sheet的个数
            int sheetNum = hssfWorkbook.getNumberOfSheets();


            System.out.println("该excel文件中总共有：" + sheetNum + "个sheet");

            //读取第1个工作表
            int i = 0;
            System.out.println("读取第" + (i + 1) + "个sheet");
            XSSFSheet sheet = hssfWorkbook.getSheetAt(i);
            //获取最后一行的num，即总行数。此处从0开始
            int maxRow = sheet.getLastRowNum();
            System.out.println(String.format("文件%s的行数为%d", file.getPath(), maxRow));


            for (int row = 1; row <= maxRow; row++) { // 跳过第一行

                //获取最后单元格num，即总单元格数 ***注意：此处从1开始计数***
                int maxRol = sheet.getRow(row).getLastCellNum();
                System.out.println("--------第" + row + "行的数据如下--------");


                for (Integer rol = 1; rol <= 5; rol++) { // 只需要读取 1~5的数据即可
                    try {
                        msg =  sheet.getRow(row).getCell(rol).toString();
                    }catch (Exception e){
                        continue;
                    }

                    String result = msg.replace(" ", "");

                    if (result == null || result.length() == 0) {
                        continue;
                    }

                    switch (rol){
                        case 1:// 记录类型
                            typeAndCountAdd(typeAndCount, result);
                            break;

                        case 2 : // 标题
                            stringBuilder.append(result);
                            break;

                        case 3:// 对接人
                            String[] split = result.split("[\\\\/,.、]");
                            for (String s : split) {
                                typeAndCountAdd(nameAndCount, s);
                            }
                            break;

                        case 4 : // 备注
                            stringBuilder.append(result);
                            break;

                        case 5 :// 原因
                            stringBuilder.append(result);
                            break;

                        default:
                            break;
                    }
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void typeAndCountAdd(Map<String,Integer> map, String key){
        if(map.containsKey(key)){
            Integer count = map.get(key);
            map.put(key,count +1);
        }else {
            map.put(key,1);
        }
    }

}