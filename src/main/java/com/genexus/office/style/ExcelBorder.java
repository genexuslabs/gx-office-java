package com.genexus.office.style;

public class ExcelBorder extends ExcelStyleDimension {
	private ExcelColor borderColor;
	private Boolean border;
	
	public ExcelColor getBorderColor() {
		return borderColor;
	}

	public Boolean getBorder() {
		return border;
	}

	public void setBorder(Boolean border) {
		this.border = border;
		setChanged();
	}

	public ExcelBorder() {		
		borderColor = new ExcelColor();
	}
	
	
}
