package com.ls.util;

import com.ls.pojo.Employee;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import org.apache.commons.lang.time.DateFormatUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Date;
import java.util.List;

/**
 * 数据写入表格
 */
public class WriteExcel {
    public void importToExcel(List<Employee> list, HttpServletResponse response){
        String fileName = DateFormatUtils.format(new Date(), "yyyyMMddhhmmss");
        response.reset();
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xls");
        WritableWorkbook wb;

        String sheetName = "First";

        try {
            OutputStream os = response.getOutputStream();
            wb = Workbook.createWorkbook(os);
            WritableSheet sheet = wb.createSheet(sheetName,0);
            Label label = new Label(0, 0, "员工编号");
            sheet.addCell(label);
            Label label1 = new Label(1, 0, "姓名");
            sheet.addCell(label1);
            Label label2 = new Label(2, 0, "性别");
            sheet.addCell(label2);
            Label label3 = new Label(3, 0, "邮箱");
            sheet.addCell(label3);
            Label label4 = new Label(4, 0, "部门");
            sheet.addCell(label4);
            int k=1;
            for(int i=0;i<list.size();i++){
                sheet.addCell(new Label(0, k, list.get(i).getId().toString()));
                sheet.addCell(new Label(1, k, list.get(i).getLastName()));
                sheet.addCell(new Label(2, k, list.get(i).getGender().toString()));
                sheet.addCell(new Label(3, k, list.get(i).getEmail()));
                sheet.addCell(new Label(4, k, list.get(i).getDepartment()));
                k++;
            }
            wb.write();
            wb.close();

        } catch (IOException e) {
            e.printStackTrace();
        } catch (RowsExceededException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }
}
