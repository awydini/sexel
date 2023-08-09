package net.aydini.sexel.writer;

import java.lang.reflect.Field;
import java.util.Set;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import net.aydini.mom.util.reflection.ReflectionUtil;
import net.aydini.sexel.annotation.SexelField;
import net.aydini.sexel.configuration.ConfigurationProperty;
import net.aydini.sexel.exception.SexelException;
import net.aydini.sexel.workbook.WorkbookHolder;

/**
 * 
 * @author <a href="mailto:hi@aydini.net">Aydin Nasrollahpour </a>
 *
 */
public class RowWriter extends AbstractWriter {

	private final Row row;

	private Object rowData;

	private Set<Field> fields;

	private boolean isHeader = false;

	private WorkbookHolder workBookHolder;

	RowWriter(ConfigurationProperty configurationProperty, Row row) {
		super(configurationProperty);
		this.row = row;
	}

	public RowWriter setWorkBookHolder(WorkbookHolder workBookHolder) {
		this.workBookHolder = workBookHolder;
		return this;
	}

	public RowWriter setIsheader(boolean isheader) {
		this.isHeader = isheader;
		return this;
	}

	public RowWriter setFields(Set<Field> fields) {
		this.fields = fields;
		return this;
	}

	public RowWriter setRowData(Object rowData) {
		this.rowData = rowData;
		return this;
	}

	protected void doWrite() throws Exception {
		fields.parallelStream().forEach(field ->{
			SexelField sexelField = field.getAnnotation(SexelField.class);
			if (sexelField != null) {
				Cell cell = row.createCell(sexelField.columnIndex());
				new CellWriter(getConfigurationProperty(), cell).setWorkBookHolder(workBookHolder)
						.setField(field).setIsHeader(isHeader)
						.setCellValue(getCellValue(field, rowData, isHeader))
						.write();
			}
		});
	}

	private Object getCellValue(Field field, Object rowData, boolean isHeader) {
		if (isHeader) {
			String headerTitle = field.getAnnotation(SexelField.class).headerTitle();
			return StringUtils.isNotEmpty(headerTitle) ? headerTitle : field.getName();
		}
		try {
			return ReflectionUtil.getFieldValueFromObject(field, rowData);
		} catch (Exception e) {
			throw new SexelException(e.getMessage(), e);
		}
	}

}
