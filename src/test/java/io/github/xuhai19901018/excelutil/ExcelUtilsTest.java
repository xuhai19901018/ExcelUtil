package io.github.xuhai19901018.excelutil;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.imageio.ImageIO;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;


public class ExcelUtilsTest {
	@Test
	public void t1() {
		
		try {
			Map<String, Object> model = new HashMap();

			model.put("no", "C160414-1");
			model.put("year", "2022");
			model.put("month", "4");
			model.put("day", "11");

			model.put("out", 1111);
			model.put("in", 222);

			ExcelUtils.addValue("model", model);

			
			List<Map<String, Object>> rows = new ArrayList<Map<String,Object>>();
			for(int i = 0; i< 30; i++) {
				Map<String, Object> map = new HashMap<String, Object>();
				map.put("mcode", Math.random());
				rows.add(map);
			}
			

			for (int i = 0; i < rows.size(); i++) {
				rows.get(i).put("index", i + 1);
			}

			List<Object> pages = new ArrayList<Object>();

			for (int i = 0; i < rows.size() / 12 + 1; i++) {

				List<Map<String, Object>> page = new ArrayList<Map<String, Object>>();

				for (int j = i * 12; j < i * 12 + 12; j++) {
					if(j< rows.size()) {
						page.add(rows.get(j));
					}
					else {
						page.add(new HashMap());
					}
					
				}
				pages.add(page);
			}

			ExcelUtils.addValue("pages", pages);

			OutputStream out = new FileOutputStream(new File("D:\\home\\t1.xlsm"));
			// 输出Excel
			ExcelUtils.export("D:\\projects\\CWT\\STS\\gwts\\src\\main\\resources\\templates\\" + "原料出库通知单.xlsm", out);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}
	
	@Test
	public void t2() {
		
		try {
				

			ExcelUtils.addValue("pages", new int[] {1,2,3});

			// 输出Excel
			ExcelUtils.export("D:\\projects\\CWT\\STS\\gwts\\src\\main\\resources\\templates\\s2.xlsm" , new FileOutputStream(new File("D:\\home\\t1.xlsm")));
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}
	
	
	@Test
	public void t3() {
		
		try {
			ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
			BufferedImage bufferImg = ImageIO.read(new File("D:\\home\\日油.jpg"));
			ImageIO.write(bufferImg, "jpg", byteArrayOut);
			
			
			Workbook wb = WorkbookUtils.openWorkbook("D:\\projects\\CWT\\STS\\gwts\\src\\main\\resources\\templates\\s3.xlsm" );
			Sheet sheet = wb.getSheetAt(0);
			
			Drawing patriarch  = sheet.createDrawingPatriarch();
			
			XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0,(short) 1, 1, (short) 2, 2);
			
			
	        patriarch.createPicture(anchor, wb.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));

			
			WorkbookUtils.copyRange(sheet, 0,0,2,100,3,0);
			
			WorkbookUtils.SaveWorkbook(wb, new FileOutputStream(new File("D:\\home\\t3.xlsm")));
			
			
			

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}
	
}
