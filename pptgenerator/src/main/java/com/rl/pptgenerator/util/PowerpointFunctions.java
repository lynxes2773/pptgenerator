package com.rl.pptgenerator.util;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.sl.usermodel.TableCell.BorderEdge;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import com.aspose.slides.IAutoShape;
import com.aspose.slides.IPPImage;
import com.aspose.slides.IShape;
import com.aspose.slides.ISlide;
import com.aspose.slides.Presentation;

import java.awt.Color;
import java.awt.Point;
import java.awt.Rectangle;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.List;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

@Component
public class PowerpointFunctions {

	private static final Logger logger = LogManager.getLogger(PowerpointFunctions.class);
	
	private List sourceData;
	private XMLSlideShow ppt;
	private XSLFTable dashboardSlideTbl;
	private XDDFDataSource<String> projectStatuses;
	private XDDFNumericalDataSource<Double> counts;
	private boolean emptyCellerrorEncountered=false;
	private boolean fileOpenErrorEncountered=false;
	
	private Presentation asposePresentation; 
	
	// Table Formatting Properties
	private static final int HEADER_ROW_BACKGROUND_RED_VAL = 0;
	private static final int HEADER_ROW_BACKGROUND_GREEN_VAL = 45;
	private static final int HEADER_ROW_BACKGROUND_BLUE_VAL = 106;
	
	private static final int VALUE_ROW_EVEN_BACKGROUND_RED_VAL = 235;
	private static final int VALUE_ROW_EVEN_BACKGROUND_GREEN_VAL = 235;
	private static final int VALUE_ROW_EVEN_BACKGROUND_BLUE_VAL = 235;
	
	private static final int VALUE_ROW_ODD_BACKGROUND_RED_VAL = 255;
	private static final int VALUE_ROW_ODD_BACKGROUND_GREEN_VAL = 255;
	private static final int VALUE_ROW_ODD_BACKGROUND_BLUE_VAL = 255;
	
	private static final String GREEN_TEXT_VAL = "Green";
	private static final String RED_TEXT_VAL = "Red";
	private static final String YELLOW_TEXT_VAL = "Yellow";
	
	private static final String CHART_TYPE_PIE = "PIE";
	
	private static final String ACTIVE_ISSUES_TITLE = "Active Issues";
	private static final String ACTIVE_RISKS_TITLE = "Active Risks";
	
	private static final int SLIDE_FONT_STYLE_TITLE=0;
	private static final int SLIDE_FONT_STYLE_SUBHEADER=1;
	private static final int SLIDE_FONT_STYLE_COLORED_STATUS_GREEN=2;
	private static final int SLIDE_FONT_STYLE_COLORED_STATUS_RED=3;
	private static final int SLIDE_FONT_STYLE_COLORED_STATUS_YELLOW=4;
	private static final int SLIDE_FONT_STYLE_REGULAR_TEXT=5;
	
	// Table Values for Comparison
	private static final String BLANK_STRING_REPRESENTATION = "(blank)";

	@Value("${error.msg.empty.cell}")
	private String errorMsgEmptyCell;

	@Value("${chart.title.text}")
	private String chartTitleText;
	
	@Value("${chart.title.portfolio.overview}")
	private String chartTitlePortfolioOverview;
	
	@Value("${project.status.chart.title}")
	private String projectStatusChartTitle;
	
	@Value("${chart.series.title.portfolio}")
	private String chartSeriesTitlePortfolioStatuses;
	
	@Value("${dashboard.title}")
	private String dashboardTitle;
	
	@Value("${output.file.name}")
	private String outputFilename;
	
	@Value("${report.save.location}")
	private String reportSaveLocation;
	
	@Value("${report.column.0}")
	private String reportColumn0;

	@Value("${report.column.1}")
	private String reportColumn1;

	@Value("${report.column.2}")
	private String reportColumn2;

	@Value("${report.column.3}")
	private String reportColumn3;

	@Value("${report.column.4}")
	private String reportColumn4;
	
	@Value("${report.column.5}")
	private String reportColumn5;

	@Value("${report.column.6}")
	private String reportColumn6;

	@Value("${report.column.7}")
	private String reportColumn7;

	@Value("${report.column.8}")
	private String reportColumn8;

	@Value("${report.column.9}")
	private String reportColumn9;
	
	@Value("${report.column.10}")
	private String reportColumn10;

	@Value("${report.column.11}")
	private String reportColumn11;
	
	@Value("${report.column.12}")
	private String reportColumn12;

	@Value("${report.column.13}")
	private String reportColumn13;
	
	@Value("${details.subheader.risks}")
	private String detailsSubheaderRisks;

	@Value("${details.subheader.issues}")
	private String detailsSubheaderIssues;
	
	@Value("${details.subheader.dependencies}")
	private String detailsSubheaderDependencies;
	
	@Value("${details.subheader.key.milestones}")
	private String detailsSubheaderKeyMilestones;
	
	@Value("${details.subheader.baseline}")
	private String detailsSubheaderBaseline;
	
	@Value("${details.subheader.key.accomplishments}")
	private String detailsSubheaderKeyAccomplishments;
	
	@Value("${details.subheader.upcoming}")
	private String detailsSubheaderUpcomingActivities;
	
	@Value("${details.subheader.comments}")
	private String detailsSubheaderComments;
	
	private boolean imagesAvailable=false;
	private IPPImage topBarPicture=null, bottomBarPicture=null, gnmaLogoPicture=null;
	
	
	static void setAllCellBorders(XSLFTableCell cell, Color color) {
		  cell.setBorderColor(BorderEdge.top, color);
		  cell.setBorderColor(BorderEdge.right, color);
		  cell.setBorderColor(BorderEdge.bottom, color);
		  cell.setBorderColor(BorderEdge.left, color);
	}	

	
	public PowerpointFunctions()
	{
		//empty constructor
	}
	
	public boolean createPowerpoint(List reportData)
	{
		try
		{
			this.sourceData=reportData;
			this.createPowerpoint();
			this.addDashboardSlide(this.ppt);
			this.writePresentationToFile(this.ppt);
			//this.buildProjectStatusSlide();
			//this.buildProjectDetailsSlides();
			//this.addImagestoDashboardSlide();
		}
		catch(Exception e)
		{
			e.printStackTrace();
			if(!fileOpenErrorEncountered)
			{
				emptyCellerrorEncountered=true;
				JOptionPane.showMessageDialog(new JFrame(), errorMsgEmptyCell,"Error",JOptionPane.ERROR_MESSAGE);
			}
		}
		boolean errorEncountered = (emptyCellerrorEncountered | fileOpenErrorEncountered); 
		return errorEncountered;
	}
	
	private void createPowerpoint()
	{
		this.ppt = new XMLSlideShow();
	}
	
	private void addDashboardSlide(XMLSlideShow ppt)
	{
		this.ppt = ppt;
		java.awt.Dimension pgsize = ppt.getPageSize();
		int pgw = pgsize.width; //slide width in points
	    int pgh = pgsize.height; //slide height in points
	    
	    ppt.setPageSize(new java.awt.Dimension(2750,1536));
	      
		XSLFTable tbl = null;
		XSLFSlide dashboardSlide = this.addNewSlide();
		XSLFTextShape dashboardTitle = dashboardSlide.getPlaceholder(0);
		Rectangle2D.Double titleAnchor = new Rectangle2D.Double(100, 150, 2500, 150);
		dashboardTitle.setAnchor(titleAnchor);
		XSLFTextShape title = dashboardSlide.getPlaceholder(0);
		title.clearText();
		XSLFTextParagraph paragraph = title.addNewTextParagraph();
		TextParagraph.TextAlign alignment = paragraph.getTextAlign();
		paragraph.setTextAlign(paragraph.getTextAlign().LEFT);
		XSLFTextRun text = paragraph.addNewTextRun();
		text.setFontColor(PowerPointFontSylingFunctions.TITLE_COLOR);
		text.setFontSize((double)96);
		text.setText(this.dashboardTitle);
		
		String result ="";

		int numRows = sourceData.size();
		int numCols = 11;
		try
		{
			dashboardSlide = createTable(dashboardSlide, numRows, numCols);
			result = "Table populated.";
		}
		catch(IllegalArgumentException iae)
		{
			iae.printStackTrace();
			result = "Table not created";
		}
		catch(Exception e)
		{
			e.printStackTrace();
			result = "Table not created";
		}
		finally
		{
			logger.debug(result);
		}
	}
	
	/**
	 * Generates the table on the Dashboard slide
	 * @param slide
	 * @param numRows
	 * @param numCols
	 * @return
	 * @throws IllegalArgumentException
	 * @throws Exception
	 */
	private XSLFSlide createTable(XSLFSlide slide, int numRows, int numCols) throws IllegalArgumentException, Exception 
	{
		  if (numRows < 1 || numCols < 1) {
		   throw new IllegalArgumentException("numRows and numCols must be greater than 0");
		  }
		  XSLFTable tbl = slide.createTable();
		  tbl.setAnchor(new Rectangle(new Point(100, 325)));
		  
		  int numRowsWithHeader = numRows +1;
		  for(int r = 0; r<numRowsWithHeader; r++) 
		  {
			   XSLFTableRow row = tbl.addRow(); // this takes over all cells from present rows
			   for (int c = 0; c < numCols; c++) 
			   {
			    	XSLFTableCell cell = row.addCell();
			    	setAllCellBorders(cell, Color.BLACK);
			   }
		  }

		  XSLFTableCell cell = null;
		  XSLFTextParagraph para = null;
		  XSLFTextRun tRun = null;

		  /**
		   * populate Header row with column names 
		   */
		  for(int c=0;c<numCols; c++)
		  {
				cell = tbl.getCell(0, c);
				cell.setFillColor(new java.awt.Color(this.HEADER_ROW_BACKGROUND_RED_VAL,this.HEADER_ROW_BACKGROUND_GREEN_VAL,this.HEADER_ROW_BACKGROUND_BLUE_VAL));
				para = cell.addNewTextParagraph();
				tRun = para.addNewTextRun();
				tRun.setBold(true);
				tRun.setFontColor(new java.awt.Color(255,255,255));
				tRun.setFontSize(25.0);

				switch(c)
				{
					case 0:
					  	tbl.setColumnWidth(c, PowerPointTableColumnWidths.COLUMN_0_WIDTH);
						tRun.setText(reportColumn0); // Project Name
						break;
					case 1:
					  	tbl.setColumnWidth(c, PowerPointTableColumnWidths.COLUMN_4_WIDTH);
						tRun.setText(reportColumn4); // Go-Live Date
						break;
					case 2:
					  	tbl.setColumnWidth(c, PowerPointTableColumnWidths.COLUMN_13_WIDTH);
					  	tRun.setText(reportColumn13); // Lifecycle Stage
						break;
					case 3:
					  	tbl.setColumnWidth(c, PowerPointTableColumnWidths.COLUMN_2_WIDTH);
					  	tRun.setText(reportColumn2); // Ginnie Mae PM
						break;
					case 4:
					  	tbl.setColumnWidth(c, PowerPointTableColumnWidths.COLUMN_3_WIDTH);
					  	tRun.setText(reportColumn3); // IT PM
						break;
					case 5:
					  	tbl.setColumnWidth(c, PowerPointTableColumnWidths.COLUMN_7_WIDTH);
					  	tRun.setText(reportColumn7); // Schedule Status
						break;
					case 6:
					  	tbl.setColumnWidth(c, PowerPointTableColumnWidths.COLUMN_9_WIDTH);
					  	tRun.setText(reportColumn9); // Budget Status
						break;
					case 7:
					  	tbl.setColumnWidth(c, PowerPointTableColumnWidths.COLUMN_8_WIDTH);
					  	tRun.setText(reportColumn8); // Scope Status
						break;
					case 8:
					  	tbl.setColumnWidth(c, PowerPointTableColumnWidths.COLUMN_10_WIDTH);
					  	tRun.setText(reportColumn10); // Risk Status
						break;
					case 9:
					  	tbl.setColumnWidth(c, PowerPointTableColumnWidths.COLUMN_11_WIDTH);
					  	tRun.setText(reportColumn11); // Active Issues
						break;
					case 10:
					  	tbl.setColumnWidth(c, PowerPointTableColumnWidths.COLUMN_12_WIDTH);
					  	tRun.setText(reportColumn12); // Active Risks
						break;
				}		
 		  }
		  
		  /**
		   * Populate the table with the data
		   */
		  String cellValue=null;
		  HashMap mapObj = null;
		  List<XSLFTableRow> tableRowList = tbl.getRows();
		  XSLFTableRow tableRow = null;
		  for(int r=0; r<numRows; r++) 
		  {
			  tableRow = tableRowList.get(r+1);
	  		  mapObj = (HashMap)sourceData.get(r);
	  		  List<XSLFTableCell> tableRowCellsList = tableRow.getCells();
	  		  
			  for(int c=0; c<numCols; c++)
			  {
				  cell = tableRowCellsList.get(c);
				  para = cell.addNewTextParagraph();
				  tRun = para.addNewTextRun();
				  tRun.setBold(false);
				  tRun.setFontColor(new java.awt.Color(0,0,0));
				  tRun.setFontSize(22.0);
				  this.setRowBasedCellFillColor(cell, r);
				  
				  switch(c)
				  {
				  	case 0:
				  		cellValue = (mapObj.get(reportColumn0)).toString(); // Project Name
				  		if(cell!=null && cellValue!=null)
				  			tRun.setText(cellValue);
				  		break;
				  	case 1:	
				  		cellValue = (mapObj.get(reportColumn4)).toString(); // Go-Live Date
				  		if(cell!=null && cellValue!=null)
				  			if(cellValue.equals(this.BLANK_STRING_REPRESENTATION))
				  			{
				  				cellValue="";
				  			}
				  			tRun.setText(cellValue);
				  		break;
				  	case 2:
				  		cellValue = (mapObj.get(reportColumn13)).toString(); // Lifecycle Stage
				  		if(cell!=null && cellValue!=null)
				  			tRun.setText(cellValue);
				  		break;
				  	case 3:
				  		cellValue = (mapObj.get(reportColumn2)).toString(); // Ginnie Mae PM
				  		if(cell!=null && cellValue!=null)
				  			tRun.setText(cellValue);
				  		break;
				  	case 4:
				  		cellValue = (mapObj.get(reportColumn3)).toString(); // IT PM
				  		if(cell!=null && cellValue!=null)
				  			if(cellValue.equals(this.BLANK_STRING_REPRESENTATION))
				  			{
				  				cellValue="";
				  			}
				  			tRun.setText(cellValue);
				  		break;
				  	case 5:
				  		cellValue = (mapObj.get(reportColumn7)).toString(); // Schedule Status
				  		if(cell!=null && cellValue!=null)
				  			if(cellValue.equals(this.BLANK_STRING_REPRESENTATION))
				  			{
				  				cellValue="";
				  			}
				  			this.setStatusBasedCellFillColor(cell, cellValue);
				  			tRun.setText(cellValue);
				  		break;
				  	case 6:
				  		cellValue = (mapObj.get(reportColumn9)).toString(); // Budget Status
				  		if(cell!=null && cellValue!=null)
				  			if(cellValue.equals(this.BLANK_STRING_REPRESENTATION))
				  			{
				  				cellValue="";
				  			}
				  			this.setStatusBasedCellFillColor(cell, cellValue);
				  			tRun.setText(cellValue);
				  		break;
				  	case 7:
				  		cellValue = (mapObj.get(reportColumn8)).toString(); // Scope Status
				  		if(cell!=null && cellValue!=null)
				  			if(cellValue.equals(this.BLANK_STRING_REPRESENTATION))
				  			{
				  				cellValue="";
				  			}
				  			this.setStatusBasedCellFillColor(cell, cellValue);
				  			tRun.setText(cellValue);
				  		break;
				  	case 8:
				  		cellValue = (mapObj.get(reportColumn10)).toString(); // Risk Status 
				  		if(cell!=null && cellValue!=null)
				  			if(cellValue.equals(this.BLANK_STRING_REPRESENTATION))
				  			{
				  				cellValue="";
				  			}
				  			this.setStatusBasedCellFillColor(cell, cellValue);
				  			tRun.setText(cellValue);
				  		break;
				  	case 9:
				  		cellValue = (mapObj.get(reportColumn11)).toString(); // Active Issues
				  		if(cell!=null && cellValue!=null)
				  			if(cellValue.equals(this.BLANK_STRING_REPRESENTATION))
				  			{
				  				cellValue="0";
				  			}
				  			tRun.setText(cellValue);
				  		break;
				  	case 10:
				  		cellValue = (mapObj.get(reportColumn12)).toString(); // Active Risks
				  		if(cell!=null && cellValue!=null)
				  			if(cellValue.equals(this.BLANK_STRING_REPRESENTATION))
				  			{
				  				cellValue="0";
				  			}
				  			tRun.setText(cellValue);
				  		break;
				  }
			  }
		  }
		  
		  return slide;
	}
	
	private void setStatusBasedCellFillColor(XSLFTableCell cell, String cellValue)
	{
		if(cellValue.equals(this.GREEN_TEXT_VAL))
		{
			cell.setFillColor(PowerPointFontSylingFunctions.GREEN_COLOR);
		}
		if(cellValue.equals(this.RED_TEXT_VAL))
		{
			cell.setFillColor(PowerPointFontSylingFunctions.RED_COLOR);
		}
		if(cellValue.equals(this.YELLOW_TEXT_VAL))
		{
			cell.setFillColor(PowerPointFontSylingFunctions.YELLOW_COLOR);
		}
	}

	private void setRowBasedCellFillColor(XSLFTableCell cell, int rowNum)
	{
		if(rowNum%2==0)
		{
			cell.setFillColor(new java.awt.Color(this.VALUE_ROW_EVEN_BACKGROUND_RED_VAL, this.VALUE_ROW_EVEN_BACKGROUND_GREEN_VAL, this.VALUE_ROW_EVEN_BACKGROUND_BLUE_VAL));
		}
		else
		{
			cell.setFillColor(new java.awt.Color(this.VALUE_ROW_ODD_BACKGROUND_RED_VAL, this.VALUE_ROW_ODD_BACKGROUND_GREEN_VAL, this.VALUE_ROW_ODD_BACKGROUND_BLUE_VAL));
		}
	}
	
	private void writePresentationToFile(XMLSlideShow ppt)
	{
		this.ppt = ppt;
		try
		{
			File file = new File(reportSaveLocation+outputFilename);
		    FileOutputStream out = new FileOutputStream(file);
		    ppt.write(out);
		    out.close();
		}
		catch(Exception e)
		{
			fileOpenErrorEncountered=true;
			JOptionPane.showMessageDialog(new JFrame(),"Powerpoint file is already open. Please close it and retry.","Error",JOptionPane.ERROR_MESSAGE);
		}
	}

	
	
	/**
	 * Add new blank siide to be used further on
	 * For Apache POI implementation
	 * @return
	 */
	private XSLFSlide addNewSlide()
	{
		XSLFSlideMaster slideMaster = ppt.getSlideMasters().get(0);
		XSLFSlideLayout contentLayout = slideMaster.getLayout(SlideLayout.TITLE_ONLY);
		XSLFSlide slide = ppt.createSlide(contentLayout);
		return slide;
	}
	
	/**
	 * Removes the default placeholders in a slide.
	 * For Aspose implementation
	 * Use it to clean up Aspose generated new slide.
	 * @param slide
	 * @return
	 */
	private ISlide removeDefaultPlaceholdersFromSlide(ISlide slide)
	{
		ISlide cleanedUpSlide = slide;
		for (IShape shape : slide.getShapes())
        {
            if (shape instanceof IAutoShape) //Checks if shape supports text frame (IAutoShape). 
            {
                IAutoShape autoShape = (IAutoShape)shape;
                autoShape.removePlaceholder();
            }
        }
		return cleanedUpSlide;
	}
}
