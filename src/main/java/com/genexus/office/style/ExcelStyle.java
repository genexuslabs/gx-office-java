package com.genexus.office.style;

public class ExcelStyle extends ExcelStyleDimension
{	
	private ExcelFill _cellFill;
	private ExcelFont _cellFont;
	private Boolean _locked;
	
	/*private ExcelBorder _cellLeftBorder;
	private ExcelBorder _cellRightBorder;
	private ExcelBorder _cellTopBorder;
	private ExcelBorder _cellBottomBorder;*/
	private ExcelAlignment _cellAlignment;
	
	public ExcelStyle() {		
		_cellFill = new ExcelFill();
		_cellFont = new ExcelFont();
		_cellAlignment = new ExcelAlignment();
		/*_cellLeftBorder = new ExcelBorder();
		_cellRightBorder = new ExcelBorder();
		_cellTopBorder = new ExcelBorder(); 
		_cellBottomBorder = new ExcelBorder();
		*/
	}
	
	public Boolean isLocked() {
		return _locked;
	}
	
	public void setLocked(boolean value) {
		_locked = value;
	}
	
	public ExcelAlignment getCellAlignment() {
		return _cellAlignment;
	}


	public ExcelFill getCellFill() {		
		return _cellFill;
	}
	
	public ExcelFont getCellFont() {		
		return _cellFont;
	}
	
	
	@Override
	public boolean isDirty() {
		return super.isDirty() || _cellFill.isDirty() || _cellFont.isDirty() || _cellAlignment.isDirty();
	}
	
}

