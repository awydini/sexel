package net.aydini.sexel.reader;

import java.math.BigDecimal;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import net.aydini.sexel.exception.SexelException;
import org.apache.poi.ss.usermodel.DateUtil;

/**
 * 
 * @author <a href="mailto:hi@aydini.net">Aydin Nasrollahpour </a>
 *
 */
public class CellValueExtractor {

	public static Object extractCellValue(Class<?> fieldType, Cell cell) {
		try {
			return extractValue(fieldType, cell);
		} catch (Exception e) {
			
			throw new SexelException(e.getMessage(),e);
		}
	}

	 private static Object extractValue(Class<?> fieldType, Cell cell) {

	        if(cell.getCellType() == CellType.BOOLEAN)
	            return cell.getBooleanCellValue();
	        if(cell.getCellType() == CellType.BLANK)
	            return null;
	        if(cell.getCellType() == CellType.STRING)
	            return cell.getStringCellValue();
	        if (DateUtil.isCellDateFormatted(cell))
	            return cell.getDateCellValue();
	        if(cell.getCellType() == CellType.NUMERIC)
	            return convertNumeric(cell.getNumericCellValue(),fieldType);
	        return cell.getStringCellValue();

	    }
	  private static Object convertNumeric(Double cellValue,Class<?> fieldClass)
	    {
	        if(cellValue == null)
	            return null;
	        if(fieldClass.equals(Double.class))
	            return cellValue;
	        if(fieldClass.equals(Long.class))
	            return cellValue.longValue();
	        if(fieldClass.equals(Integer.class))
	            return cellValue.intValue();
	        if(fieldClass.equals(Float.class))
	            return cellValue.floatValue();
	        if(fieldClass.equals(BigDecimal.class))
	            return new BigDecimal(cellValue);
	        return String.valueOf(cellValue.longValue());
	    }
}
