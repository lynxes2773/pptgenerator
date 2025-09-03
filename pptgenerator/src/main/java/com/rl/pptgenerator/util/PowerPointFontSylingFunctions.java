package com.rl.pptgenerator.util;

import java.awt.Color;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import com.aspose.slides.FillType;
import com.aspose.slides.FontData;
import com.aspose.slides.IAutoShape;
import com.aspose.slides.IEffectFormat;
import com.aspose.slides.IOuterShadow;
import com.aspose.slides.IParagraph;
import com.aspose.slides.IPortion;
import com.aspose.slides.ITextFrame;
import com.aspose.slides.LineStyle;
import com.aspose.slides.TextAlignment;

public class PowerPointFontSylingFunctions {

	private static final Logger logger = LogManager.getLogger(PowerPointFontSylingFunctions.class);
	
	public static final java.awt.Color TITLE_COLOR = new java.awt.Color(2,43,105);
	public static final java.awt.Color SUBHEADER_COLOR = new java.awt.Color(125,125,125);
	public static final java.awt.Color VALUE_COLOR = new java.awt.Color(236,236,236);
	public static final java.awt.Color GREEN_COLOR = new java.awt.Color(185,213,178);
	public static final java.awt.Color RED_COLOR = new java.awt.Color(243,115,97);
	public static final java.awt.Color YELLOW_COLOR = new java.awt.Color(248,239,135);
	
	private static final String statusGreenColor = "Green";
	private static final String statusRedColor = "Red";
	private static final String statusYellowColor = "Yellow";
	
	private static final String blankValue = "(blank)";
	
	public static String cleanupBlankTextValue(String blankTextValue)
	{
		if(blankTextValue.trim().equals(blankValue))
		{
			blankTextValue="";
		}
		return blankTextValue;
	}
	
	public static String cleanuptActiveCounts(String countStr)
	{
		int cnt = 0;
		if(countStr==null || countStr.trim().equals("") || countStr.equals(blankValue))
		{
			countStr= cnt+"";
		}
		return countStr;
	}
	
	public static IAutoShape alignTextLeft(IAutoShape shape)
	{
		IParagraph paragraph = shape.getTextFrame().getParagraphs().get_Item(0);
		paragraph.getParagraphFormat().setAlignment(TextAlignment.Left);
		paragraph.getParagraphFormat().setIndent(10f);
		return shape;
	}

	public static IAutoShape alignTextCenter(IAutoShape shape)
	{
		IParagraph paragraph = shape.getTextFrame().getParagraphs().get_Item(0);
		paragraph.getParagraphFormat().setAlignment(TextAlignment.Center);
		return shape;
	}

	
	public static IAutoShape setTitle(IAutoShape shape, String text)
	{
		text = cleanupBlankTextValue(text);
		shape.addTextFrame(" ");
		shape.getFillFormat().setFillType(FillType.NoFill);
		shape.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getPortionFormat().getFillFormat().setFillType(FillType.Solid);
		shape.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getPortionFormat().getFillFormat().getSolidFillColor().setColor(TITLE_COLOR);
		shape.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getPortionFormat().setFontHeight(96);
		shape.getLineFormat().setWidth(0);
		shape.getLineFormat().getFillFormat().setFillType(FillType.Solid);
	    shape.getLineFormat().getFillFormat().getSolidFillColor().setColor(new java.awt.Color(255,255,255));		

		FontData fontData = new FontData("Arial");
		shape.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getPortionFormat().setComplexScriptFont(fontData);

		ITextFrame textFrame = shape.getTextFrame();
		textFrame.getParagraphs().get_Item(0).getParagraphFormat().setAlignment(TextAlignment.Left);
		textFrame.setText(text);

		return shape;
	}

	public static IAutoShape setSubheader(IAutoShape shape, String text)
	{
		text = cleanupBlankTextValue(text);
		shape.addTextFrame(" ");
		shape.getFillFormat().setFillType(FillType.Solid);
		shape.getFillFormat().getSolidFillColor().setColor(SUBHEADER_COLOR);
		shape.getLineFormat().setWidth(0);
		shape.getLineFormat().getFillFormat().setFillType(FillType.Solid);
	    shape.getLineFormat().getFillFormat().getSolidFillColor().setColor(SUBHEADER_COLOR);		
		
		shape.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getPortionFormat().getFillFormat().setFillType(FillType.Solid);
		shape.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getPortionFormat().getFillFormat().getSolidFillColor().setColor(new java.awt.Color(255,255,255));
		shape.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getPortionFormat().setFontHeight(28);
		
		FontData fontData = new FontData("Verdana");
		shape.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getPortionFormat().setComplexScriptFont(fontData);
		
		ITextFrame textFrame = shape.getTextFrame();
		textFrame.setText(text);

		return shape;
	}

	public static IAutoShape setValue(IAutoShape shape, String text)
	{
		text = cleanupBlankTextValue(text);
		shape.addTextFrame(" ");
		shape.getFillFormat().setFillType(FillType.Solid);
		shape.getFillFormat().getSolidFillColor().setColor(VALUE_COLOR);
		shape.getLineFormat().setWidth(0);
		shape.getLineFormat().getFillFormat().setFillType(FillType.Solid);
	    shape.getLineFormat().getFillFormat().getSolidFillColor().setColor(VALUE_COLOR);		
		
		shape.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getPortionFormat().getFillFormat().setFillType(FillType.Solid);
		shape.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getPortionFormat().getFillFormat().getSolidFillColor().setColor(new java.awt.Color(0,0,0));
		shape.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getPortionFormat().setFontHeight(24);
		
		FontData fontData = new FontData("Verdana");
		shape.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getPortionFormat().setComplexScriptFont(fontData);
		
		ITextFrame textFrame = shape.getTextFrame();
		textFrame.setText(text);

		return shape;
	}
	
	
	public static IAutoShape setStatusBasedValue(IAutoShape shape, String text)
	{
		text = cleanupBlankTextValue(text);
		java.awt.Color color = null;
		if(text.equals(statusGreenColor))
		{	
			color = GREEN_COLOR;
		}
		else if(text.equals(statusRedColor))
		{	
			color = RED_COLOR;
		}
		else if(text.equals(statusYellowColor))
		{	
			color = YELLOW_COLOR;
		}
		else
		{
			color = VALUE_COLOR;
		}
		shape.addTextFrame(" ");
		shape.getFillFormat().setFillType(FillType.Solid);
		shape.getFillFormat().getSolidFillColor().setColor(color);
		shape.getLineFormat().setWidth(0);
		shape.getLineFormat().getFillFormat().setFillType(FillType.Solid);
	    shape.getLineFormat().getFillFormat().getSolidFillColor().setColor(color);		
		
		shape.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getPortionFormat().getFillFormat().setFillType(FillType.Solid);
		shape.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getPortionFormat().getFillFormat().getSolidFillColor().setColor(new java.awt.Color(0,0,0));
		shape.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getPortionFormat().setFontHeight(24);
		
		FontData fontData = new FontData("Verdana");
		shape.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getPortionFormat().setComplexScriptFont(fontData);
		
		ITextFrame textFrame = shape.getTextFrame();
		textFrame.setText(text);

		return shape;
	}

}
