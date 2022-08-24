package io.github.xuhai19901018.excelutil.tags;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.StringTokenizer;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import io.github.xuhai19901018.excelutil.ExcelParser;

/***
 * 2022年6月22日
 * 图片插入
 * @author xuhai
 *
 */
public class ImgTag implements ITag {

	public static final String KEY_IMG = "#img";

	public int[] parseTag(Object context, Sheet sheet, Row curRow, Cell curCell) throws Exception {
		curCell.setCellValue("");
		String expr = "";
		String img = curCell.getStringCellValue();
		StringTokenizer st = new StringTokenizer(img, " ");
		int width = 0;
		int height = 0;
		int dx1 = 0;
		int dy1 = 0;
		int dx2 = 0;
		int dy2 = 0;

		int pos = 0;
		while (st.hasMoreTokens()) {
			String str = st.nextToken();
			if (pos == 1) {
				expr = str;
			}
			if (pos == 2) {
				width = Integer.parseInt(str);
			}
			if (pos == 3) {
				height = Integer.parseInt(str);
			}
			if (pos == 4) {
				dx1 = Integer.parseInt(str);
			}
			if (pos == 5) {
				dy1 = Integer.parseInt(str);
			}
			if (pos == 6) {
				dx2 = Integer.parseInt(str);
			}
			if (pos == 7) {
				dy2 = Integer.parseInt(str);
			}
			pos++;
		}

		String imgPath = (String) ExcelParser.parseExpr(context, expr);

		if (null == imgPath)
			return new int[] { 0, 0, 0 };

		ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
		BufferedImage bufferImg = ImageIO.read(new File(imgPath));
		ImageIO.write(bufferImg, "png", byteArrayOut);

		Drawing patriarch = sheet.createDrawingPatriarch();

		XSSFClientAnchor anchor = new XSSFClientAnchor(dx1*XSSFShape.EMU_PER_POINT, dy1*XSSFShape.EMU_PER_POINT, dx2*XSSFShape.EMU_PER_POINT, dy2*XSSFShape.EMU_PER_POINT, (short) curCell.getColumnIndex(), curCell.getRowIndex(), (short) (curCell.getColumnIndex() + width), curCell.getRowIndex() + height);

		patriarch.createPicture(anchor, sheet.getWorkbook().addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_PNG));

		curCell.setCellValue("");
		
		return new int[] { 0, 0, 0 };
	}

	public String getTagName() {
		return KEY_IMG;
	}

	public boolean hasEndTag() {
		return false;
	}
}
