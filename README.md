# poi-tl-plus
Enhancement to POI-TL (https://github.com/Sayi/poi-tl).   
Support defining Table templates directly in Microsoft Word (Docx) file.

<pre>
POI-TL的 MiniTableRenderData 可以支持简单的表格,但是表格样式和内容的样式无法在 Word 中直接定制.  

POI-TL 还提供了 DynamicTableRenderPolicy 支持把需要动态渲染的部分单元格交给自定义模板渲染策略,
但是样例提供的程序代码只适用于此特定的样例,不是通用的代码.   

如果想在你自己的业务模块中, 也实现动态渲染的部分单元格交给自定义模板渲染策略, 避免不了写代码来操作word文件.  

鉴于此, 此 Repository 把动态渲染的部分单元格交给自定义模板渲染策略抽象出来,做成了通用的 Table 渲染策略,
期望适用于大部分自定义表格模板的场合.  

Maven 依赖:  
除了依赖 poi-tl (请参见 https://github.com/Sayi/poi-tl ) 之外,
只额外依赖 lombok (你也可以去掉,自己补全代码里面的 getter/setter 方法即可), 

Features:
1. 常见表格的定义完全可以实现在 Word 中定义: 完全所见即所得的在 Word 文档中定义表格的表头和样式,   
数据行的单元格和样式,设置数据行每列的绑定字段, 即可实现表格的自动生成.  

2. 表格中,除了可以绑定到来自数据库的源数据之外,还可增加行号列, 自动渲染出行号值.  

3. 完全兼容 POI-TL 已有的功能. 比如支持多文档合并, 每个文档内都可以使用动态表格渲染策略来渲染不同的多个表格.  


具体样例如下:   
1) Word模板以及表格的定义:  
* 模板主文档(projectTemplate.docx), 可以看到里面POI-TL的各自常用标签都可以使用,
而且表格可以自由定义包括样式 (包括底色,对齐,字体以及字体颜色等等),
注意: 
  在一个文档中, 可以添加多个表格,每个表格都可以绑定到不同的数据集上;
  表格的数据绑定: 例如 {{teamMembers}} 就把对应的表格绑定到下面 Java代码里的 teamMembers 集合上.
  表格各列的数据绑定: 在数据行的单元格里,使用 ==fieldName 即可绑定到表格数据的字段上.
  特别地, 使用 ==ROW_NUM 可以绑定到内置的自动生成的行号字段.
  
(贴图: https://github.com/wujianbin-sh/poi-tl-plus/blob/master/projectTemplateDocx.jpg )
<img src="https://github.com/wujianbin-sh/poi-tl-plus/blob/master/projectTemplateDocx.jpg"/> 

* 子文档模板 projectMilestoneTemplate.docx, 同上,可以使用 POI-TL库和本库提供的表格标签做数据绑定:
(todo: 贴图)

2) Java 代码准备数据模型:  
  // suppose data relationships are:       
  // 1:N project -> team members : project.getTeamMembers();     
  // 1:N project -> stake holders: project.getStakeholders();     
  // 1:N project -> milestones: project.getMilestones();     
  // 1:N milestone -> deliverables: milestone.getDeliverables();   

  // first get your data from Database(or somewhere else)   
  List<Project> projectData = projectDao.getById(projectId);   

  String outputFile = "outputFile.docx";   
  String projectTemplateFile = "Project.docx";   
  String milestoneTemplateFile = "projectMilestone.docx";   

  int startRow = 1;   

  // For each Table in the Docx template: define its TableRowRenderPloicy  
  TableRowRenderPloicy teamMemberTablePloicy = new TableRowRenderPloicy("teamMembers", startRow);   
  TableRowRenderPloicy stakeholderTablePloicy = new TableRowRenderPloicy("stakeholders", startRow);   
  TableRowRenderPloicy deliverableTablePloicy = new TableRowRenderPloicy("deliverables", startRow);     

  //Prepare data model for Docx template data binding:  
  @SuppressWarnings("serial")  
  Map<String, Object> data = new HashMap<String, Object>() {{  
    // for Project.docx: that's the master Docx template  
    put("project", projectData);   
    put(teamMemberTablePloicy.tableKey, projectData.getTeamMembers());   
    put(stakeholderTablePloicy.tableKey, projectData.getStakeholders());   


    // for projectMilestone.docx: that's the sub Docx template for each of project Milestones  
    DocxRenderData projectMilestonesDoc = new DocxTableRenderData(milestoneTemplateFile,
      projectData.getMilestones(), 
      (milestone)->{  
        return new HashMap<String, Object>() {{  
            put("milestone", milestone);   
            put(deliverableTablePloicy.tableKey, ((Milestone)milestone).getDeliverables());   
        }};   
    });   

    // add sub docx template data into master docx template data:  
    put("projectMilestonesDoc", projectMilestonesDoc);   
  }};   

  try {  
    // generate docx now: outputFile.docx will be generated  
    DocxTableRenderPolicy.renderDocx(projectTemplateFile, outputFile, data,    
      teamMemberTablePloicy, stakeholderTablePloicy, deliverableTablePloicy);     
  } catch (Exception e) {  
    e.printStackTrace();   
  }  


3) 渲染的结果(生成的文档):  
(todo: 贴图)  
