package com.poitlplus;

import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.deepoove.poi.NiceXWPFDocument;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.Configure.ConfigureBuilder;
import com.deepoove.poi.data.DocxRenderData;
import com.deepoove.poi.policy.DocxRenderPolicy;
import com.deepoove.poi.template.run.RunTemplate;

/**
 * 
 * @author WuJianbin
 */
public class DocxTableRenderPolicy extends DocxRenderPolicy{

	@Override
	public void doRender(RunTemplate runTemplate, DocxRenderData data, XWPFTemplate template) throws Exception {
		NiceXWPFDocument doc = template.getXWPFDocument();
		XWPFRun run = runTemplate.getRun();
		List<NiceXWPFDocument> docMerges = getMergedDocxs(data, template.getConfig());
		doc = doc.merge(docMerges, run);
		template.reload(doc);
	}

	protected List<NiceXWPFDocument> getMergedDocxs(DocxRenderData docData, Configure configure) throws IOException {
		DocxTableRenderData data = (DocxTableRenderData)docData;

		List<NiceXWPFDocument> docs = new ArrayList<NiceXWPFDocument>();
		byte[] docx = data.getDocx();
		List<?> dataList = data.getRenderDatas();
		if (null == dataList || dataList.isEmpty()) {
			docs.add(new NiceXWPFDocument(new ByteArrayInputStream(docx)));
		} else {
			for (int i = 0; i < dataList.size(); i++) {
				XWPFTemplate temp = XWPFTemplate.compile(new ByteArrayInputStream(docx), configure);

				Map<String, Object> tableData = data.functionToGetTableData.apply(dataList.get(i));

				temp.render(tableData);
				docs.add(temp.getXWPFDocument());
			}
		}
		return docs;
	}

	public static DocxTableRenderPolicy docxTableRenderPolicy;   	

	protected static final List<String> emptyList = new ArrayList<>();

	public static Configure getConfigure(TableRowRenderPloicy... tableRowRenderPloicies) {
		docxTableRenderPolicy = new DocxTableRenderPolicy();

		ConfigureBuilder builder = Configure.newBuilder()
				.addPlugin('+', docxTableRenderPolicy);	

		for(int i=0; i<tableRowRenderPloicies.length; i++) {
			builder.customPolicy(tableRowRenderPloicies[i].tableKey, tableRowRenderPloicies[i]);  //// for each inner Table, process row data
		}

		return builder.build();
	}

	public static void renderDocx(String masterTemplateFile, String outputFile, Object data, TableRowRenderPloicy... tableRowRenderPloicies) throws Exception {
		Configure config = DocxTableRenderPolicy.getConfigure(tableRowRenderPloicies);

		XWPFTemplate template = XWPFTemplate.compile(masterTemplateFile, config).render(data);
		FileOutputStream out = new FileOutputStream(outputFile);
		template.write(out);
		out.flush();
		out.close();
	}   

}
