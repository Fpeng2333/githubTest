package com.springboot.collection.util;

import java.io.FileOutputStream;
import java.util.*;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;


public class DownExcel {

	/**
	 * 导出
	 *
	 * @param model
	 * @param response
	 * @return
	 * @throws Exception
	 */
	@SuppressWarnings({ "unchecked", "resource", "deprecation" })
	public static void exportExcel(Map<String, Object> model,
								   HttpServletResponse response) throws Exception {

		//整理传过来的数据
		String title = model.get("title").toString();
		List<Map<String, Object>> dataList = (List<Map<String, Object>>) model.get("dataList");
		String[] tableHeads = model.get("tableHead").toString().split(",");

		// 第一步，创建一个webbook，对应一个Excel文件  
		HSSFWorkbook wb = new HSSFWorkbook();

		// 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
		HSSFSheet sheet = wb.createSheet(title);

		// 第四步，创建单元格，并设置值表头 设置表头居中
		HSSFCellStyle style = wb.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式
		style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER); // 垂直居中

		//合并单元格
		CellRangeAddress one = new CellRangeAddress(0,0,0,tableHeads.length-1);
		sheet.addMergedRegion(one);

		// 第一行
		HSSFRow row = sheet.createRow(0);
		row.setHeight((short) 900);
		HSSFCell cell = row.createCell(0);
		cell.setCellValue(title);
		cell.setCellStyle(style);

		//第二行
		row = sheet.createRow(1);
		row.setHeight((short) 600);
		for(int i = 0; i < tableHeads.length; i++){
			cell = row.createCell(i);
			cell.setCellValue(tableHeads[i].substring(0,tableHeads[i].indexOf("-")).replace(" ", ""));
			cell.setCellStyle(style);
		}

		//第三行
		row = sheet.createRow(2);
		row.setHeight((short) 600);
		for(int i = 0; i < tableHeads.length; i++){
			int s = tableHeads[i].substring(tableHeads[i].indexOf("-")+1).replace(" ", "").split("").length;
			sheet.setColumnWidth(i, 2000*s);
			cell = row.createCell(i);
			cell.setCellValue(tableHeads[i].substring(tableHeads[i].indexOf("-")+1).replace(" ", ""));
			cell.setCellStyle(style);
		}

		//第四行向下的数据
		if ("true".equals(model.get("source").toString())) {
			for (int i = 0; i < dataList.size(); i++) {
				row = sheet.createRow(i+3);
				row.setHeight((short) 500);
				for (int j = 0; j < tableHeads.length; j++) {
					cell = row.createCell(j);
					cell.setCellValue(dataList.get(i).get(tableHeads[j].substring(0,tableHeads[j].indexOf("-")).replace(" ", "")).toString());
					cell.setCellStyle(style);
				}
			}
		} else if ("false".equals(model.get("source").toString())) {

		}

		//存储路径 记得要改
		FileOutputStream fileOut = new FileOutputStream(model.get("downPath").toString()+title+".xls");
		wb.write(fileOut);
		fileOut.close();

		System.out.println("yes");

	}

	/**
	 * 导入整理数据
	 *
	 * 整理出格式例如
	 *
	 * List<Map<String, Object>>
	 *
	 * [{id=8, created=2018-05-15 10:39:16.0, name=aa}, {id=8, created=2018-05-15 10:39:16.0, name=aa}]
	 *
	 * @param file
	 * @return
	 * @throws Exception
	 */
	@SuppressWarnings("resource")
	public static List<Map<String, Object>> importExcel(MultipartFile file) throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook(file.getInputStream());
		// 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
		XSSFSheet sheet = wb.getSheetAt(0);
		//获得总列数
		int coloumNum = 14;
		//获取总行数
		int rowNum = sheet.getLastRowNum();
		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		XSSFRow row1 = sheet.getRow(0);
		XSSFRow row;
		XSSFCell cell;
		for (int i = 0; i < rowNum; i++) {
			row = sheet.getRow(i+1);
			if (row != null) {
				Map<String, Object> map = new HashMap<String, Object>();
				for (int j = 0; j < coloumNum; j++) {
					cell = row.getCell(j);
					String value = "";
					if (cell != null) {
						switch (cell.getCellType()) {
							case XSSFCell.CELL_TYPE_FORMULA:
								break;
							case XSSFCell.CELL_TYPE_NUMERIC:
								cell.setCellType(Cell.CELL_TYPE_STRING);
								// 防止把1 取成1.0
								value = cell.getStringCellValue();
								break;
							case XSSFCell.CELL_TYPE_STRING:
								value = cell.getRichStringCellValue().getString();
								break;
							default:
								value = "";
								break;
						}
					} else {
						value = "";
					}
					if(row1.getCell(j).toString().equals("姓名")){
						map.put("lawyerName", value);
					}
					if(row1.getCell(j).toString().equals("城市")){
						map.put("city", value);
					}
					if(row1.getCell(j).toString().equals("律师事务所")){
						map.put("company", value);
					}
					if(row1.getCell(j).toString().equals("照片")){
						map.put("picture", value);
					}
					if(row1.getCell(j).toString().equals("年龄")){
						map.put("age", value);
					}
					if(row1.getCell(j).toString().equals("执业证号")){
						map.put("licenseNo", value);
					}
					if(row1.getCell(j).toString().equals("律师电话")){
						map.put("mobilephone1", value);
					}
					if(row1.getCell(j).toString().equals("主管司法局")){
						map.put("chargeJudicialBureau", value);
					}
					if(row1.getCell(j).toString().equals("地址1")){
						map.put("address", value);
					}
					if(row1.getCell(j).toString().equals("统一信用代码")){
						map.put("creditCode", value);
					}
					if(row1.getCell(j).toString().equals("单位电话")){
						map.put("phone", value);
					}
					if(row1.getCell(j).toString().equals("负责人")){
						if(value.equals("否"))
							map.put("chargePerson", 0);
						else{
							map.put("chargePerson",1);
						}
					}
				}
				list.add(map);
			}
		}
		return list;
	}

	public static List<Map<String, Object>> importExcel1(MultipartFile file) throws Exception {
		HSSFWorkbook wb = new HSSFWorkbook(file.getInputStream());
		// 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
		HSSFSheet sheet = wb.getSheetAt(0);
		//获得总列数
		int coloumNum = 14;
		//获取总行数
		int rowNum = sheet.getLastRowNum();
		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		HSSFRow row1 = sheet.getRow(0);
		HSSFRow row;
		HSSFCell cell;
		for (int i = 0; i < rowNum; i++) {
			row = sheet.getRow(i+1);
			if (row != null) {
				Map<String, Object> map = new HashMap<String, Object>();
				for (int j = 0; j < coloumNum; j++) {
					cell = row.getCell(j);
					String value = "";
					if (cell != null) {
						switch (cell.getCellType()) {
							case HSSFCell.CELL_TYPE_FORMULA:
								break;
							case HSSFCell.CELL_TYPE_NUMERIC:
								cell.setCellType(Cell.CELL_TYPE_STRING);
								// 防止把1 取成1.0
								value = cell.getStringCellValue();
								break;
							case HSSFCell.CELL_TYPE_STRING:
								value = cell.getRichStringCellValue().getString();
								break;
							default:
								value = "";
								break;
						}
					} else {
						value = "";
					}
					if(row1.getCell(j).toString().equals("姓名")){
						map.put("lawyerName", value);
					}
					if(row1.getCell(j).toString().equals("城市")){
						map.put("city", value);
					}
					if(row1.getCell(j).toString().equals("律师事务所")){
						map.put("company", value);
					}
					if(row1.getCell(j).toString().equals("照片")){
						map.put("picture", value);
					}
					if(row1.getCell(j).toString().equals("年龄")){
						map.put("age", value);
					}
					if(row1.getCell(j).toString().equals("执业证号")){
						map.put("licenseNo", value);
					}
					if(row1.getCell(j).toString().equals("律师电话")){
						map.put("mobilephone1", value);
					}
					if(row1.getCell(j).toString().equals("主管司法局")){
						map.put("chargeJudicialBureau", value);
					}
					if(row1.getCell(j).toString().equals("地址1")){
						map.put("address", value);
					}
					if(row1.getCell(j).toString().equals("统一信用代码")){
						map.put("creditCode", value);
					}
					if(row1.getCell(j).toString().equals("单位电话")){
						map.put("phone", value);
					}
					if(row1.getCell(j).toString().equals("负责人")){
						if(value=="否")
							map.put("chargePerson", 0);
						else{
							map.put("chargePerson",1);
						}
					}
				}
				list.add(map);
			}
		}
		return list;
	}

	public static List<Map<String, Object>> importPictureExcel(MultipartFile file) throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook(file.getInputStream());
		// 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
		XSSFSheet sheet = wb.getSheetAt(0);
		//获得总列数
		int coloumNum = sheet.getRow(1).getPhysicalNumberOfCells();
		//获取总行数
		int rowNum = sheet.getLastRowNum();
		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		XSSFRow row1 = sheet.getRow(0);
		XSSFRow row;
		XSSFCell cell;
		for (int i = 0; i < rowNum; i++) {
			row = sheet.getRow(i+1);
			if (row != null) {
				Map<String, Object> map = new HashMap<String, Object>();
				for (int j = 0; j < coloumNum; j++) {
					cell = row.getCell(j);
					String value = "";
					if (cell != null) {
						switch (cell.getCellType()) {
							case XSSFCell.CELL_TYPE_FORMULA:
								break;
							case XSSFCell.CELL_TYPE_NUMERIC:
								cell.setCellType(Cell.CELL_TYPE_STRING);
								// 防止把1 取成1.0
								value = cell.getStringCellValue();
								break;
							case XSSFCell.CELL_TYPE_STRING:
								value = cell.getRichStringCellValue().getString();
								break;
							default:
								value = "";
								break;
						}
					} else {
						value = "";
					}
					if(row1.getCell(j).toString().equals("执业证号")){
						map.put("licenseNo", value);
					}
					if(row1.getCell(j).toString().equals("照片")){
						map.put("picture", value);
					}
				}
				list.add(map);
			}
		}
		return list;
	}

	public static List<Map<String, Object>> importPictureExcel1(MultipartFile file) throws Exception {
		HSSFWorkbook wb = new HSSFWorkbook(file.getInputStream());
		// 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
		HSSFSheet sheet = wb.getSheetAt(0);
		//获得总列数
		int coloumNum = sheet.getRow(1).getPhysicalNumberOfCells();
		//获取总行数
		int rowNum = sheet.getLastRowNum();
		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		HSSFRow row1 = sheet.getRow(0);
		HSSFRow row;
		HSSFCell cell;
		for (int i = 0; i < rowNum; i++) {
			row = sheet.getRow(i+1);
			if (row != null) {
				Map<String, Object> map = new HashMap<String, Object>();
				for (int j = 0; j < coloumNum; j++) {
					cell = row.getCell(j);
					String value = "";
					if (cell != null) {
						switch (cell.getCellType()) {
							case HSSFCell.CELL_TYPE_FORMULA:
								break;
							case HSSFCell.CELL_TYPE_NUMERIC:
								cell.setCellType(Cell.CELL_TYPE_STRING);
								// 防止把1 取成1.0
								value = cell.getStringCellValue();
								break;
							case HSSFCell.CELL_TYPE_STRING:
								value = cell.getRichStringCellValue().getString();
								break;
							default:
								value = "";
								break;
						}
					} else {
						value = "";
					}
					if(row1.getCell(j).toString().equals("执业证号")){
						map.put("licenseNo", value);
					}
					if(row1.getCell(j).toString().equals("照片")){
						map.put("picture", value);
					}
				}
				list.add(map);
			}
		}
		return list;
	}

	public static List<Map<String, Object>> importCaseExcel(MultipartFile file) throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook(file.getInputStream());
		// 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
		XSSFSheet sheet = wb.getSheetAt(0);
		//获得总列数
		int coloumNum = sheet.getRow(0).getPhysicalNumberOfCells();
		//获取总行数
		int rowNum = sheet.getLastRowNum();
		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		XSSFRow row1 = sheet.getRow(0);
		XSSFRow row;
		XSSFCell cell;
		for (int i = 0; i < rowNum; i++) {
			row = sheet.getRow(i+1);
			if (row != null) {
				Map<String, Object> map = new HashMap<String, Object>();
				for (int j = 0; j < coloumNum; j++) {
					cell = row.getCell(j);
					String value = "";
					if (cell != null) {
						switch (cell.getCellType()) {
							case XSSFCell.CELL_TYPE_FORMULA:
								break;
							case XSSFCell.CELL_TYPE_NUMERIC:
								cell.setCellType(Cell.CELL_TYPE_STRING);
								// 防止把1 取成1.0
								value = cell.getStringCellValue();
								break;
							case XSSFCell.CELL_TYPE_STRING:
								value = cell.getRichStringCellValue().getString();
								break;
							default:
								value = "";
								break;
						}
					} else {
						value = "";
					}
					if (row.getCell(0)!=null){
						if (row.getCell(0).getRichStringCellValue().getString()==""){
							if(row1.getCell(j).toString().equals("申请执行人委托代理人")){
								Object obj = list.get(i-1).get("executionApplicantAgent");
								list.get(i-1).put("executionApplicantAgent", obj + "/" + value);
							}
							if(row1.getCell(j).toString().equals("执业证号")){
								Object obj = list.get(i-1).get("licenseNo");
								list.get(i-1).put("licenseNo", obj + "/" + value);
							}
						} else {
							if(row1.getCell(j).toString().equals("申请执行人姓名")){
								map.put("executionApplicantName", value);
							}
							if(row1.getCell(j).toString().equals("申请执行人出生年月")){
								map.put("executionApplicantBirth", value);
							}
							if(row1.getCell(j).toString().equals("申请执行人证件号码")){
								map.put("executionApplicantNo", value);
							}
							if(row1.getCell(j).toString().equals("申请执行人手机号码")){
								map.put("executionApplicantPhone", value);
							}
							if(row1.getCell(j).toString().equals("申请执行人委托代理人")){
								map.put("executionApplicantAgent", value);
							}
							if(row1.getCell(j).toString().equals("执业证号")){
								map.put("licenseNo", value);
							}
							if(row1.getCell(j).toString().equals("被执行人姓名")){
								map.put("executeeName", value);
							}
							if(row1.getCell(j).toString().equals("证件号码")){
								map.put("executeeNo", value);
							}
							if(row1.getCell(j).toString().equals("执行案号")){
								map.put("executionCaseNo", value);
							}
							if(row1.getCell(j).toString().equals("执行标的")){
								map.put("subjectMatterExecution", value);
							}
							if(row1.getCell(j).toString().equals("被执行人手机号码")){
								map.put("executeePhone",value);
							}
							if(row1.getCell(j).toString().equals("被执行人联系电话")){
								if (value.equals("NULL"))
									map.put("executeeFixedTelephone","");
								else
									map.put("executeeFixedTelephone",value);
							}
							if(row1.getCell(j).toString().equals("执行法院")){
								map.put("courtExecution",value);
							}
						}
					}
				}
				map.put("feedbackContent","");
				list.add(map);
			}
		}
		return list;
	}

	public static List<Map<String, Object>> importCaseExcel1(MultipartFile file) throws Exception {
		HSSFWorkbook wb = new HSSFWorkbook(file.getInputStream());
		// 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
		HSSFSheet sheet = wb.getSheetAt(0);
		//获得总列数
		int coloumNum = sheet.getRow(1).getPhysicalNumberOfCells();
		//获取总行数
		int rowNum = sheet.getLastRowNum();
		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		HSSFRow row1 = sheet.getRow(0);
		HSSFRow row;
		HSSFCell cell;
		for (int i = 0; i < rowNum; i++) {
			row = sheet.getRow(i + 1);
			if (row != null) {
				Map<String, Object> map = new HashMap<String, Object>();
				for (int j = 0; j < coloumNum; j++) {
					cell = row.getCell(j);
					String value = "";
					if (cell != null) {
						switch (cell.getCellType()) {
							case HSSFCell.CELL_TYPE_FORMULA:
								break;
							case HSSFCell.CELL_TYPE_NUMERIC:
								cell.setCellType(Cell.CELL_TYPE_STRING);
								// 防止把1 取成1.0
								value = cell.getStringCellValue();
								break;
							case HSSFCell.CELL_TYPE_STRING:
								value = cell.getRichStringCellValue().getString();
								break;
							default:
								value = "";
								break;
						}
					} else {
						value = "";
					}
					if (row.getCell(0) != null) {
						if (row.getCell(0).getRichStringCellValue().getString() == "") {
							if (row1.getCell(j).toString().equals("申请执行人委托代理人")) {
								Object obj = list.get(i - 1).get("executionApplicantAgent");
								list.get(i - 1).put("executionApplicantAgent", obj + "/" + value);
							}
							if (row1.getCell(j).toString().equals("执业证号")) {
								Object obj = list.get(i - 1).get("licenseNo");
								list.get(i - 1).put("licenseNo", obj + "/" + value);
							}
						} else {
							if (row1.getCell(j).toString().equals("申请执行人姓名")) {
								map.put("executionApplicantName", value);
							}
							if (row1.getCell(j).toString().equals("申请执行人出生年月")) {
								map.put("executionApplicantBirth", value);
							}
							if (row1.getCell(j).toString().equals("申请执行人证件号码")) {
								map.put("executionApplicantNo", value);
							}
							if (row1.getCell(j).toString().equals("申请执行人手机号码")) {
								map.put("executionApplicantPhone", value);
							}
							if (row1.getCell(j).toString().equals("申请执行人委托代理人")) {
								map.put("executionApplicantAgent", value);
							}
							if (row1.getCell(j).toString().equals("执业证号")) {
								map.put("licenseNo", value);
							}
							if (row1.getCell(j).toString().equals("被执行人姓名")) {
								map.put("executeeName", value);
							}
							if (row1.getCell(j).toString().equals("证件号码")) {
								map.put("executeeNo", value);
							}
							if (row1.getCell(j).toString().equals("执行案号")) {
								map.put("executionCaseNo", value);
							}
							if (row1.getCell(j).toString().equals("执行标的")) {
								map.put("subjectMatterExecution", value);
							}
							if (row1.getCell(j).toString().equals("被执行人手机号码")) {
								map.put("executeePhone", value);
							}
							if (row1.getCell(j).toString().equals("被执行人联系电话")) {
								if (value.equals("NULL"))
									map.put("executeeFixedTelephone","");
								else
									map.put("executeeFixedTelephone",value);
							}
							if (row1.getCell(j).toString().equals("执行法院")) {
								map.put("courtExecution", value);
							}
						}
					}
					map.put("feedbackContent","");
					list.add(map);
				}
			}
		}
		return list;
	}
	
	
	
	
	/*FileInputStream inp = new FileInputStream("E:\\WEIAN.xls"); 
	HSSFWorkbook wb = new HSSFWorkbook(inp);
	HSSFSheet sheet = wb.getSheetAt(2); // 获得第三个工作薄(2008工作薄)
	// 填充上面的表格,数据需要从数据库查询
	HSSFRow row5 = sheet.getRow(4); // 获得工作薄的第五行
	HSSFCell cell54 = row5.getCell(3);// 获得第五行的第四个单元格
	cell54.setCellValue("测试纳税人名称");// 给单元格赋值
	//获得总列数
	int coloumNum=sheet.getRow(0).getPhysicalNumberOfCells();
	int rowNum=sheet.getLastRowNum();//获得总行数
*/
}
