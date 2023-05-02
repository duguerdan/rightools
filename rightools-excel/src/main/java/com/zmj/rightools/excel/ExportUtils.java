package com.zmj.rightools.excel;

import com.zmj.rightools.excel.entity.Personnel;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author zhangmingjian
 * @date 2023/5/1
 */
@Slf4j
public class ExportUtils {
    
    public static void export() throws FileNotFoundException {
        FileOutputStream outputStream = new FileOutputStream("C:/Users/admin/Desktop/excel/personnel.xlsx");
        String[] titles = {"ID", "姓名", "生日", "性别", "国家", "行业", "备注"};
        Personnel personnel = new Personnel();
        personnel.setId(1);
        personnel.setName("张三");
        personnel.setBirthday(new Date());
        personnel.setGender(1);
        personnel.setCountry("中国");
        personnel.setIndustry("IT");
        ArrayList<Personnel> personnels = new ArrayList<>();
        personnels.add(personnel);
        createExcel(titles, personnels, outputStream);
    }
    
    public static void createExcel(String[] titles, List<Personnel> personnels, OutputStream outputStream) {
        log.info("创建excel");
        try (
                XSSFWorkbook workbook = new XSSFWorkbook()
        ) {
            // 创建sheet
            XSSFSheet sheet = workbook.createSheet("人员信息");
            // 设置列宽
            sheet.setColumnWidth(2, 256 * 20);
            // 创建表头
            XSSFRow titleRow = sheet.createRow(0);
            for (int i = 0; i < titles.length; i++) {
                XSSFCell cell = titleRow.createCell(i);
                cell.setCellValue(titles[i]);
            }
            // 时间格式
            XSSFDataFormat dataFormat = workbook.createDataFormat();
            XSSFCellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setDataFormat(dataFormat.getFormat("yyyy-MM-dd HH:mm:ss"));
            // 填充数据
            for (int i = 0; i < personnels.size(); i++) {
                Personnel personnel = personnels.get(i);
                XSSFRow row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(personnel.getId());
                row.createCell(1).setCellValue(personnel.getName());
                XSSFCell cell2 = row.createCell(2);
                cell2.setCellValue(personnel.getBirthday());
                cell2.setCellStyle(cellStyle);
                row.createCell(3).setCellValue(personnel.getGender() == 1 ? "男" : "女");
                row.createCell(4).setCellValue(personnel.getCountry());
                row.createCell(5).setCellValue(personnel.getIndustry());
                row.createCell(6).setCellValue(personnel.getComment());
            }
            workbook.write(outputStream);
        } catch (Exception e) {
            log.error("create excel error");
            throw new RuntimeException(e);
        }
        log.info("excel创建完成");
    }
    
}
