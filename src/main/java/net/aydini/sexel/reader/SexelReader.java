package net.aydini.sexel.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import net.aydini.sexel.configuration.ConfigurationProperty;
import net.aydini.sexel.constant.Constants;
import net.aydini.sexel.exception.SexelException;

/**
 * 
 * @author <a href="mailto:hi@aydini.net">Aydin Nasrollahpour </a>
 *
 */
public class SexelReader extends AbstractReader<List<Object>> {

	private String filePath;

	private Class<?> outputClass = Object.class;

	private final Set<String> sheetsNames;

	private Workbook workbook;

	public SexelReader() {
		this(new ConfigurationProperty().setStartRow(1));
	}

	public SexelReader(ConfigurationProperty configurationProperty) {
		super(configurationProperty);
		sheetsNames = new HashSet<>();
	}

	public SexelReader setFilePath(String filePath) {
		this.filePath = filePath;
		return this;
	}

	public SexelReader setSheetName(String sheetName) {
		if (StringUtils.isAllEmpty(sheetName))
			throw new RuntimeException("invalid sheetName");
		sheetsNames.add(sheetName);
		return this;
	}

	public SexelReader setOutputClass(Class<?> outputClass) {
		if (outputClass != null)
			this.outputClass = outputClass;
		return this;
	}

	private void validate() {
		if (StringUtils.isAllEmpty(filePath))
			throw new SexelException("invalid file path");
		if (!FilenameUtils.getExtension(filePath).toLowerCase().equals(Constants.ACCEPTABLE_EXCEL_TYPE))
			throw new SexelException("only .xlsx is supported");
	}

	public List<Object> doRead() {
		validate();
		initWorkBook();
		List<Object> list = new ArrayList<>();
		sheetsNames.forEach(item -> list.addAll(doRead(item)));
		return list;
	}

	private List<Object> doRead(String sheetName) {

		Sheet sheet = workbook.getSheet(sheetName);
		return (List<Object>) new SheetReader(sheet, getConfigurationProperty()).setOutputClass(outputClass).read();
	}

	private void initWorkBook() {
		try {
			workbook = new HSSFWorkbook(new FileInputStream(new File(filePath)));
		} catch (IOException e) {
			throw new SexelException(e.getMessage(), e);
		}
	}

}