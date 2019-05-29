package com.itrms.barcode;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

@Controller
public class BarcodeController {
	
	@Autowired
	private JdbcTemplate jdbcTemplate;
	
	@ResponseBody
	@RequestMapping("/submitFile")
	public void submitFile(@RequestParam("file") MultipartFile file,HttpServletRequest request, HttpServletResponse response) {
		//存储所有数据
		List<Map<String, Object>> barcodeFile = new ArrayList<>();
		Sheet sheet = null;
        Row row = null;
        
		try {
			InputStream inputStream = file.getInputStream();
			
	        //获取Excel工作薄
	        Workbook work = this.getWorkbook(inputStream, file.getOriginalFilename());
	        if (null == work) {
	            throw new Exception("创建Excel工作薄为空！");
	        }
	        //获取第一个标签页
            sheet = work.getSheetAt(0);
            int firstRowNum = sheet.getFirstRowNum(),lastRowNum = sheet.getLastRowNum();
            System.out.println("firstRowNum: " + firstRowNum);
            System.out.println("lastRowNum: " + lastRowNum);
            for (int j = firstRowNum; j <= lastRowNum; j++) {
            	row = sheet.getRow(j);
            	if(row == null) {
            		continue;
            	}else {
            		Map<String, Object> barcodeDetail = new HashMap<>();
            		//条码
            		if(row.getCell(4) != null && !String.valueOf(row.getCell(4)).startsWith("Barcode")
            				&& !String.valueOf(row.getCell(4)).startsWith("null") 
            				&& !"".equals(String.valueOf(row.getCell(4)))) {
            			row.getCell(4).setCellType(HSSFCell.CELL_TYPE_STRING);
            			String barcode = String.valueOf(row.getCell(4));
            			
            			//条码以M开头的全部替换成0，S结尾的删去
            			barcode = barcode.replaceAll("M", "0");
            			barcode = barcode.replaceAll("S", "");
            			
            			barcodeDetail.put("barcode", barcode);
            		}else {
            			continue;
            		}
            		//商品名称
            		if(row.getCell(3) != null && !String.valueOf(row.getCell(3)).startsWith("Product")) {
            			row.getCell(3).setCellType(HSSFCell.CELL_TYPE_STRING);
            			barcodeDetail.put("product", String.valueOf(row.getCell(3)));
            		}
            		//商家信息清单
            		if(row.getCell(2) != null && !String.valueOf(row.getCell(2)).startsWith("Brand")) {
            			row.getCell(2).setCellType(HSSFCell.CELL_TYPE_STRING);
            			barcodeDetail.put("brand", String.valueOf(row.getCell(2)));
            		}
            		//容量
            		if(row.getCell(5) != null && !String.valueOf(row.getCell(5)).startsWith("Volume")) {
            			row.getCell(5).setCellType(HSSFCell.CELL_TYPE_STRING);
            			String vol = excludeSpecial(String.valueOf(row.getCell(5)));
            			if("".equals(vol)) vol = "500";
            			barcodeDetail.put("volume", vol);
            		}
            		//重量
            		if(row.getCell(8) != null && !String.valueOf(row.getCell(8)).startsWith("Weight")) {
            			row.getCell(8).setCellType(HSSFCell.CELL_TYPE_STRING);
            			String weight = excludeSpecial(String.valueOf(row.getCell(8)));
            			if("".equals(weight)) weight = "30";
            			barcodeDetail.put("weight", weight);
            		}
            		//材质
            		if(row.getCell(7) != null && !String.valueOf(row.getCell(7)).startsWith("Plastic")) {
            			row.getCell(7).setCellType(HSSFCell.CELL_TYPE_STRING);
            			String pla = String.valueOf(row.getCell(7));
            			if(!"PET".equals(pla) && Integer.parseInt(String.valueOf(barcodeDetail.get("volume"))) <= 330) {
            				barcodeDetail.put("plastic", "非PET");
            			}else {
            				barcodeDetail.put("plastic", "PET");
            			}
            		}
            		//产品系列 &&瓶型组
            		if("PET".equals(String.valueOf(barcodeDetail.get("plastic")))) {
            			barcodeDetail.put("cpxl", "飲料");
            			barcodeDetail.put("pxz", "100(PET)");
            		}else {
            			barcodeDetail.put("cpxl", "豆漿");
            			barcodeDetail.put("pxz", "300(非PET)");
            		}
            		barcodeFile.add(barcodeDetail);
            	}
            }
            work.close();
            
            //现已获取所有数据，开始生成excel
    		List<String> title = new ArrayList<String>();
    		title.add("條碼");
    		title.add("商品名稱");
    		title.add("瓶型組ID");
    		title.add("商家信息清單");
    		title.add("產品系列");
    		title.add("容量(ml)");
    		title.add("材質");
    		title.add("重量(g)");
    		HSSFWorkbook work1 = this.downloadModelExcel(title,barcodeFile);
    		
    		try {
                FileOutputStream fout = new FileOutputStream("E:/Members.xls");
                work1.write(fout);
                fout.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private Workbook getWorkbook(InputStream inStr, String fileName) throws Exception {
        Workbook workbook = null;
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        if (".xls".equals(fileType)) {
            workbook = new HSSFWorkbook(inStr);
        } else if (".xlsx".equals(fileType)) {
            workbook = new XSSFWorkbook(inStr);
        } else {
            throw new Exception("请上传excel文件！");
        }
        return workbook;
    }

	/*去掉前後空格和特殊字符*/
	private String excludeSpecial(String str) {
		Pattern p = Pattern.compile("[&\\|\\\\\\*^%$#@\\-!！……￥?？{}《》~]");//去除特殊字符
		Matcher m = p.matcher(str);
		str = m.replaceAll("").trim().replace("\\", "");//将匹配的特殊字符转变为空
		return str;
	}
	
	private HSSFWorkbook downloadModelExcel(List<String> title,List<Map<String, Object>> list) {
		//创建workbook对象
		HSSFWorkbook workBook = new HSSFWorkbook();
		//创建工作薄页1
		HSSFSheet sheet = workBook.createSheet("Sheet1");
		//创建行对象
		HSSFRow row = sheet.createRow(0);
		//设置行高
		row.setHeight((short)500);
		//创建每个单元格的样式对象;
		HSSFCellStyle cellStyle = workBook.createCellStyle();
		//居中
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		//创建字体对象
		HSSFFont font = workBook.createFont();
		//将字体变为粗体
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		//将字体颜色变红色
		font.setColor(HSSFColor.RED.index);
		//字体大小
		font.setFontHeightInPoints((short)16);
		//修改样式字体
		cellStyle.setFont(font);
		//列总数
		int lie = title.size();
		for(int i = 0; i < lie; i++) {
			//创建单元格对象
			HSSFCell cell = row.createCell(i);
			//设置单元格的类型
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			//设置每列的宽度；
			sheet.setColumnWidth(i, 6000);
			//修改单元格样式
			cell.setCellStyle(cellStyle);
			//输入单元格值
			cell.setCellValue(title.get(i));
		}
		
		for(int i = 1,len = list.size(); i <= len; i ++ ) {
			//创建行对象
			HSSFRow r = sheet.createRow(i);
			//设置行高
			r.setHeight((short)250);
			
			HSSFCell cell = r.createCell(0);
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			sheet.setColumnWidth(0, 6000);
			cell.setCellValue(String.valueOf(list.get(i-1).get("barcode")));
			
			cell = r.createCell(1);
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			sheet.setColumnWidth(1, 6000);
			cell.setCellValue(String.valueOf(list.get(i-1).get("product")));
			
			cell = r.createCell(2);
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			sheet.setColumnWidth(2, 6000);
			cell.setCellValue(String.valueOf(list.get(i-1).get("pxz")));
			
			cell = r.createCell(3);
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			sheet.setColumnWidth(3, 6000);
			cell.setCellValue(String.valueOf(list.get(i-1).get("brand")));
			
			cell = r.createCell(4);
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			sheet.setColumnWidth(4, 6000);
			cell.setCellValue(String.valueOf(list.get(i-1).get("cpxl")));
			
			cell = r.createCell(5);
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			sheet.setColumnWidth(5, 6000);
			cell.setCellValue(String.valueOf(list.get(i-1).get("volume")));
			
			cell = r.createCell(6);
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			sheet.setColumnWidth(6, 6000);
			cell.setCellValue(String.valueOf(list.get(i-1).get("plastic")));
			
			cell = r.createCell(7);
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			sheet.setColumnWidth(7, 6000);
			cell.setCellValue(String.valueOf(list.get(i-1).get("weight")));
		}
		return workBook;
	}
	
	@ResponseBody
	@RequestMapping("/getSql")
	public void getSql(@RequestParam("file") MultipartFile file,HttpServletRequest request, HttpServletResponse response) {
		//存储所有数据
		List<Map<String, Object>> barcodeFile = new ArrayList<>();
		Sheet sheet = null;
	    Row row = null;
	    
		try {
			InputStream inputStream = file.getInputStream();
			
	        //获取Excel工作薄
	        Workbook work = this.getWorkbook(inputStream, file.getOriginalFilename());
	        if (null == work) {
	            throw new Exception("创建Excel工作薄为空！");
	        }
	        //获取第一个标签页
	        sheet = work.getSheetAt(0);
	        int firstRowNum = sheet.getFirstRowNum(),lastRowNum = sheet.getLastRowNum();
	        for (int j = firstRowNum; j <= lastRowNum; j++) {
	        	row = sheet.getRow(j);
	        	if(row == null) {
	        		continue;
	        	}else {
	        		Map<String, Object> barcodeDetail = new HashMap<>();
	        		//条码
	        		if(row.getCell(0) != null && !String.valueOf(row.getCell(0)).startsWith("條碼")
	        				&& !String.valueOf(row.getCell(0)).startsWith("null") 
	        				&& !"".equals(String.valueOf(row.getCell(0)))) {
	        			row.getCell(0).setCellType(HSSFCell.CELL_TYPE_STRING);
	        			String barcode = String.valueOf(row.getCell(0));
	        			
	        			String sql = "select * from res_bar_code where BAR_CODE = '" +barcode+ "' ";
	        			Map<String, Object> map = jdbcTemplate.queryForMap(sql);
	        			if(map == null) {
	        				continue;
	        			}else {
	        				//查看其他信息
	        			}
	        			
	        		}else {
	        			continue;
	        		}
	        		//商品名称
	        		if(row.getCell(3) != null && !String.valueOf(row.getCell(3)).startsWith("Product")) {
	        			row.getCell(3).setCellType(HSSFCell.CELL_TYPE_STRING);
	        			barcodeDetail.put("product", String.valueOf(row.getCell(3)));
	        		}
	        		//商家信息清单
	        		if(row.getCell(2) != null && !String.valueOf(row.getCell(2)).startsWith("Brand")) {
	        			row.getCell(2).setCellType(HSSFCell.CELL_TYPE_STRING);
	        			barcodeDetail.put("brand", String.valueOf(row.getCell(2)));
	        		}
	        		//容量
	        		if(row.getCell(5) != null && !String.valueOf(row.getCell(5)).startsWith("Volume")) {
	        			row.getCell(5).setCellType(HSSFCell.CELL_TYPE_STRING);
	        			String vol = excludeSpecial(String.valueOf(row.getCell(5)));
	        			if("".equals(vol)) vol = "500";
	        			barcodeDetail.put("volume", vol);
	        		}
	        		//重量
	        		if(row.getCell(8) != null && !String.valueOf(row.getCell(8)).startsWith("Weight")) {
	        			row.getCell(8).setCellType(HSSFCell.CELL_TYPE_STRING);
	        			String weight = excludeSpecial(String.valueOf(row.getCell(8)));
	        			if("".equals(weight)) weight = "30";
	        			barcodeDetail.put("weight", weight);
	        		}
	        		//材质
	        		if(row.getCell(7) != null && !String.valueOf(row.getCell(7)).startsWith("Plastic")) {
	        			row.getCell(7).setCellType(HSSFCell.CELL_TYPE_STRING);
	        			String pla = String.valueOf(row.getCell(7));
	        			if(!"PET".equals(pla) && Integer.parseInt(String.valueOf(barcodeDetail.get("volume"))) <= 330) {
	        				barcodeDetail.put("plastic", "非PET");
	        			}else {
	        				barcodeDetail.put("plastic", "PET");
	        			}
	        		}
	        		//产品系列 &&瓶型组
	        		if("PET".equals(String.valueOf(barcodeDetail.get("plastic")))) {
	        			barcodeDetail.put("cpxl", "飲料");
	        			barcodeDetail.put("pxz", "100(PET)");
	        		}else {
	        			barcodeDetail.put("cpxl", "豆漿");
	        			barcodeDetail.put("pxz", "300(非PET)");
	        		}
	        		barcodeFile.add(barcodeDetail);
	        	}
	        }
	        work.close();
	        
	        //现已获取所有数据，开始生成excel
			List<String> title = new ArrayList<String>();
			title.add("條碼");
			title.add("商品名稱");
			title.add("瓶型組ID");
			title.add("商家信息清單");
			title.add("產品系列");
			title.add("容量(ml)");
			title.add("材質");
			title.add("重量(g)");
			HSSFWorkbook work1 = this.downloadModelExcel(title,barcodeFile);
			
			try {
	            FileOutputStream fout = new FileOutputStream("E:/Members.xls");
	            work1.write(fout);
	            fout.close();
	        } catch (Exception e) {
	            e.printStackTrace();
	        }
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}