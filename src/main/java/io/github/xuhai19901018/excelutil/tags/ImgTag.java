package io.github.xuhai19901018.excelutil.tags;

import java.awt.*;
import java.awt.geom.AffineTransform;
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

		String expr = "";
		String img = curCell.getStringCellValue();
		StringTokenizer st = new StringTokenizer(img, " ");
		int width = 0;
		int height = 0;
		int dx1 = 0;
		int dy1 = 0;
		int dx2 = 0;
		int dy2 = 0;
		int rotateAngle = 0;

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
			if (pos == 8) {
				rotateAngle = Integer.parseInt(str);
			}
			pos++;
		}

		String imgPath = (String) ExcelParser.parseExpr(context, expr);

		if (null == imgPath)
		{
			curCell.setCellValue("");
			return new int[] { 0, 0, 0 };
		}

		// 读取图片
		BufferedImage bufferImg = ImageIO.read(new File(imgPath));
		ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
		if(rotateAngle!=0){
			// 旋转图片，不要裁剪, 使用透明色填充
			// Convert degrees to radians
			double radians = Math.toRadians(rotateAngle);

			// Calculate the new image dimensions
			int widthi = bufferImg.getWidth();
			int heighti = bufferImg.getHeight();
			double sin = Math.abs(Math.sin(radians));
			double cos = Math.abs(Math.cos(radians));
			int newWidth = (int) Math.floor(widthi * cos + heighti * sin);
			int newHeight = (int) Math.floor(heighti * cos + widthi * sin);

			// Create a new image with ARGB type to support transparency
			BufferedImage rotatedImage = new BufferedImage(newWidth, newHeight, BufferedImage.TYPE_INT_ARGB);

			// Create a graphics context and apply the rotation transformation
			Graphics2D g2d = rotatedImage.createGraphics();
			// Enable anti-aliasing for smoother edges
			g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
			// Set the background to transparent
			g2d.setComposite(AlphaComposite.Src);
			g2d.setColor(new Color(255, 255, 255, 0)); // Transparent color
			g2d.fillRect(0, 0, newWidth, newHeight);

			// Apply the rotation transformation
			AffineTransform transform = new AffineTransform();
			transform.translate((newWidth - widthi) / 2, (newHeight - heighti) / 2);
			transform.rotate(radians, widthi / 2, heighti / 2);
			g2d.setTransform(transform);
			g2d.drawImage(bufferImg, 0, 0, null);
			g2d.dispose();

//			String tempFile = System.getProperty("java.io.tmpdir") + File.separator +new File(imgPath).getName() + ".png";
//			ImageIO.write(rotatedImage, "png", new File(tempFile));
//			ImageIO.write(rotatedImage, "png", new File("D:\\home\\logo0.png"));

			ImageIO.write(rotatedImage, "png", byteArrayOut);
		}
		else{
			//		BufferedImage bufferImg = ImageIO.read(new File(imgPath));
			ImageIO.write(bufferImg, "png", byteArrayOut);
		}

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
