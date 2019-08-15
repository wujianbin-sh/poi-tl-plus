package com.poitlplus;

import java.io.File;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

import com.deepoove.poi.data.DocxRenderData;

import lombok.Getter;

/**
 * 
 * @author WuJianbin
 */
public class DocxTableRenderData extends DocxRenderData {
	@Getter
	Function<Object, Map<String, Object>> functionToGetTableData;

	public DocxTableRenderData(String fileName, List<?> dataListForPages,
			Function<Object, Map<String, Object>> functionToGetTableData) {
		super(new File(fileName), dataListForPages);
		this.functionToGetTableData = functionToGetTableData;
	}

	public DocxTableRenderData(File docxFile, List<?> dataListForPages,
			Function<Object, Map<String, Object>> functionToGetTableData) {
		super(docxFile, dataListForPages);
		this.functionToGetTableData = functionToGetTableData;
	}
	
	public DocxTableRenderData(InputStream docxInputStream, List<?> dataListForPages,
			Function<Object, Map<String, Object>> functionToGetTableData) {
		super(docxInputStream, dataListForPages);
		this.functionToGetTableData = functionToGetTableData;
	}
}
