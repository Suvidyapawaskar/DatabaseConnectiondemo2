package com.model;

import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import com.mysql.jdbc.Driver;

public class MercuryTsDbConnectionTest {

	public static void main(String[] args) throws ClassNotFoundException, SQLException, IOException 
	{
		int rowno=0;
		int colno=0;
		
		Class.forName("com.mysql.jdbc.Driver");
		Connection con=DriverManager.getConnection("jdbc:mysql://localhost:3306/test", "root", "root");
		Statement st=con.createStatement();
		ResultSet res=st.executeQuery("select * from employee");
		
		while(res.next())
		{
			String empnm=res.getString("Emp_name");
			System.out.println(empnm);
			
			WriteExcelData obj=new WriteExcelData();
			obj.writeExcel("F:\\Suvidya_data\\Java_Selenium\\DatabaseConnectDemo\\data.xlsx", "Sheet1", rowno++, colno, empnm);
			
			
		}
		
		

	}

}
