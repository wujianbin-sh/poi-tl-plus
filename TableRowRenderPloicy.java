package com.poitlplus;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTParaRPr;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.style.Style;
import com.deepoove.poi.policy.DynamicTableRenderPolicy;
import com.deepoove.poi.policy.MiniTableRenderPolicy;

import lombok.Getter;
import lombok.Setter;

/**
 * 
 * @author WuJianbin
 */
public class TableRowRenderPloicy extends DynamicTableRenderPolicy {

	private static final Logger logger = LoggerFactory.getLogger(TableRowRenderPloicy.class);

	private static final String FIELD_ROW_NUM = "ROW_NUM";
	private static final String FIELD_PREFIX = "==";
	private static int FIELD_PREFIX_LENGTH = FIELD_PREFIX.length();

	@Setter
	Collection<?> dataList;
	public String tableKey;
	@Getter
	String rowTag;

	int startRow;

	List<String> cellFields = new ArrayList<>();
	List<Style> cellStyles = new ArrayList<>();;

	public TableRowRenderPloicy(String tableKey, int startRow) {
		this(tableKey, startRow, null, null);
	}

	public TableRowRenderPloicy(String tableKey, int startRow, TableColumnRenderPloicy columnPolicy) {
		this(tableKey, startRow, null, columnPolicy);
	}

	public TableRowRenderPloicy(String tableKey, int startRow, List<?> dataList, TableColumnRenderPloicy columnPolicy) {
		this.tableKey = tableKey;
		this.startRow = startRow;
		this.dataList = dataList;
		this.rowTag = tableKey + "ROW";
	}

	private boolean isStringNullOrEmpty(String str) {
		return str == null || "".contentEquals(str.trim());

	}

	protected Style parseCellStyle(XWPFTableCell cell) {

		XWPFParagraph paragraph = cell.getParagraphs().get(0);

		try {
			CTParaRPr pr = paragraph.getCTP().getPPr().getRPr();

			Style style = new Style();
			style.setBold(pr.isSetB());
			style.setItalic(pr.isSetI());
			style.setUnderLine(pr.isSetU());
			style.setStrike(pr.isSetStrike());
			if (pr.isSetRFonts())
				style.setFontFamily(pr.getRFonts().getAscii());
			if (pr.isSetSz())
				style.setFontSize(Integer.parseInt((pr.getSz().getVal()).toString()) / 2);
			if (pr.isSetColor())
				style.setColor(pr.getColor().xgetVal().getStringValue());
			return style;
		} catch (Exception ex) {
			logger.error("faield to get style, " + cell.getText());
			logger.error(ex.getMessage());
			return null;
		}
	}

	protected void parseCellFieldsAndStyles(XWPFTableRow templateRow) {
		if (cellFields.size() > 0)
			return; // fields already parsed, so return immediately

		templateRow.getTableCells().forEach(cell -> {
			String cellText = cell.getText();
			String fieldName = null;
			Style style = null;
			if (!isStringNullOrEmpty(cellText)) {
				if (cellText.startsWith(FIELD_PREFIX)) {
					fieldName = cellText.substring(FIELD_PREFIX_LENGTH);
					style = parseCellStyle(cell);
				}
			}

			cellFields.add(fieldName);
			cellStyles.add(style);
		});
	}

	protected TextRenderData buildCellData(int columnIndex, String cellValue) {
		Style style = cellStyles.get(columnIndex);
		return (style == null) ? new TextRenderData(cellValue) : new TextRenderData(cellValue, style);
	}

	protected TextRenderData[] buildRowData(Object rowData, int rowIndex) {
		List<TextRenderData> fieldValues = new ArrayList<>();

		for (int i = 0; i < cellFields.size(); i++) {
			String fieldName = cellFields.get(i);

			if (fieldName == null) {
				fieldValues.add(buildCellData(i, ""));
			} else if (FIELD_ROW_NUM.equals(fieldName)) {
				fieldValues.add(buildCellData(i, Integer.toString(rowIndex - startRow + 1)));
			} else {
				try {
					fieldValues.add(buildCellData(i, BeanUtils.getProperty(rowData, fieldName)));
				} catch (Exception ex) {
					fieldValues.add(buildCellData(i, ""));
					logger.error("faield to get fieldValue, " + ex.getMessage());
				}
			}
		}

		return fieldValues.toArray(new TextRenderData[fieldValues.size()]);
	}

	public void cloneCell(XWPFTableRow newRow, XWPFTableRow templateRow) {
		templateRow.getTableCells().forEach(cell -> {
			XWPFTableCell newCell = newRow.createCell();
			newCell.setColor(cell.getColor());

			if (cell.getVerticalAlignment() != null) {
				newCell.setVerticalAlignment(XWPFVertAlign.valueOf(cell.getVerticalAlignment().name()));
			}

			XWPFParagraph newPara = newCell.getParagraphs().get(0);
			cell.getParagraphs().forEach(para -> {
				if (!isStringNullOrEmpty(cell.getText())) {
					newPara.getCTP().setPPr(para.getCTP().getPPr());
				}
			});
		});
	}

	@SuppressWarnings("unchecked")
	@Override
	public void render(XWPFTable table, Object data) {
		if (data == null) {
			logger.error("table data not set, tableKey=" + tableKey);
			return;
		}
		if (!(data instanceof Collection)) {
			logger.error("table data is not a Collection, tableKey=" + tableKey);
			return;
		}

		XWPFTableRow templateTabelRow = table.getRow(startRow);
		parseCellFieldsAndStyles(templateTabelRow);

		dataList = (Collection<Object>) data;
		int rowIndex = startRow;

		Iterator<?> it = dataList.iterator();
		while (it.hasNext()) {
			Object rowData = it.next();
			RowRenderData row = RowRenderData.build(buildRowData(rowData, rowIndex));
			XWPFTableRow newTableRow = table.insertNewTableRow(rowIndex + 1);

			cloneCell(newTableRow, templateTabelRow);

			MiniTableRenderPolicy.Helper.renderRow(table, rowIndex + 1, row);
			++rowIndex;
		}
		table.removeRow(startRow);
	}

}
