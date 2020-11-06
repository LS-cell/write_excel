package com.ls.servlet;

import com.ls.exception.ExcelException;
import com.ls.pojo.Employee;
import com.ls.util.WriteExcel;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


public class ExcelServlet extends HttpServlet {
    public void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException{
        List<Employee> employees = new ArrayList<Employee>();
        Employee employee1 = new Employee();
        employee1.setId(1589);
        employee1.setLastName("张军");
        employee1.setGender(1);
        employee1.setEmail("1589@qq.com");
        employee1.setDepartment("java开发");
        employees.add(employee1);

        Employee employee2 = new Employee();
        employee2.setId(1596);
        employee2.setLastName("向梅");
        employee2.setGender(0);
        employee2.setEmail("1596@qq.com");
        employee2.setDepartment("web开发");
        employees.add(employee2);

        Employee employee3 = new Employee();
        employee3.setId(1645);
        employee3.setLastName("李哲");
        employee3.setGender(1);
        employee3.setEmail("1645@qq.com");
        employee3.setDepartment("大数据开发");
        employees.add(employee3);

        WriteExcel writeExcel = new WriteExcel();
        writeExcel.importToExcel(employees,response);
    }
}

