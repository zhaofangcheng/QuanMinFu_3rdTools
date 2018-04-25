package com.zfc.biz;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang.StringUtils;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 * 银联全民付邮件通知，交易金额累加工具
 * 业余民间版本，不承担任何资金风险。
 * @author zfc
 * zhaofangcheng@126.com
 */
public class EmailParse {
	
	static double sum=0;
	/**
	 * @param args
	 * @throws IOException 
	 * @throws BiffException 
	 */
	public static void main(String[] args) throws BiffException, IOException {
		String excelPath="E:\\全民付";
		sum=getFiles(excelPath,sum);
		System.out.println("合计金额---------"+sum);

	}
	/**
	 * 获取目录下的文件
	 * @param excelPath
	 * @param sum
	 * @return
	 * @throws BiffException
	 * @throws IOException
	 */
	public static double getFiles(String excelPath,double sum) throws BiffException, IOException{
		List<File> filelist=new ArrayList();
		File dir = new File(excelPath);
		File[] files = dir.listFiles();  
		if (files != null) {
		    for (int i = 0; i < files.length; i++) {
		        String fileName = files[i].getName();
	            String strFileName = files[i].getAbsolutePath();
	            sum=parseExcel(files[i],sum);
	            filelist.add(files[i]);
		    }
		}
		return sum;
	}
	public static void test() throws BiffException, IOException{
		double sum=0;
		parseExcel( new File("E:\\全民付\\5_20170821_9.xls"),sum);
	}
	/**
	 * 解析Excel内容，获取交易金额累加数据
	 * @param excelFile
	 * @param sum
	 * @return
	 */
	public static double parseExcel(File excelFile,Double sum){
		try {
			Workbook wb=Workbook.getWorkbook(excelFile);
			Sheet sheet=wb.getSheet(0);
			int x=5;
			getXaxis(excelFile.getName(),sheet,x);
			int rows=sheet.getRows();
			for(int i=3;i<rows;i++){
				String cellText=sheet.getCell(x, i).getContents();
				if(StringUtils.isNotBlank(cellText)&&parseDouble(cellText)!=0){
					sum=CalculateUtil.add(sum, parseDouble(cellText));
				}
			}
			return sum;
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("异常了-----------"+excelFile.getName() +e);
		} 
		return sum;
	}
	/**
	 * 递归得到交易金额所在的x坐标
	 * @param sheet
	 * @param x
	 */
	public static void getXaxis(String fileName,Sheet sheet,int x){
		Cell cell=sheet.getCell(x,2);
		if(!"交易金额".equals(cell.getContents())){
			x=x+1;
			getXaxis(fileName,sheet,x);
		}
	}
	/**
	 * @param cellText
	 * @return
	 */
	public static double parseDouble(String cellText){
		try {
			return Double.parseDouble(cellText);
		} catch (NumberFormatException e) {
			return 0;
		}
	}
}
