package com.appian.plugins.export_sql_to_excel;

import java.io.OutputStream;
import java.sql.Timestamp;
import java.util.Date;
import java.util.HashMap;
import java.util.Locale;

import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;

//import com.appiancorp.analytics.excel_export.ExportHelper;
import com.appiancorp.services.ASLIdentity;
import com.appiancorp.services.ServiceContext;
import com.appiancorp.services.WebServiceContextFactory;
import com.appiancorp.suiteapi.common.LocalObject;
import com.appiancorp.suiteapi.common.ObjectTypeMapping;
import com.appiancorp.suiteapi.common.ServiceLocator;
//import com.appiancorp.suiteapi.process.analytics2.Column;
//import com.appiancorp.suiteapi.process.analytics2.ProcessAnalyticsService;
//import com.appiancorp.suiteapi.process.analytics2.ProcessReport;
//import com.appiancorp.suiteapi.process.analytics2.ReportData;
//import com.appiancorp.suiteapi.process.analytics2.ReportResultPage;
//import com.appiancorp.suiteapi.type.TypedValue;

/**
 *
 * @author corbin.page Servlet to export SQL data to Excel.
 */

//public class exportSQLToExcel {
//
//	public static void main(String[] args) {
//		// TODO Auto-generated method stub
//
//	}
//
//}


public class exportSQLToExcel extends HttpServlet {

	private static final Logger LOG = Logger
			.getLogger(exportSQLToExcel.class);

	private static final String PARAM_REPORT_ID = "reportId";
	private static final String PARAM_FILENAME = "filename";
	private static final String PARAM_CONTEXT = "context";
	private static final String PARAM_START_INDEX = "startIndex";
	private static final String PARAM_BATCH_SIZE = "batchSize";

//	private ReportData reportData;
	private HSSFWorkbook wb;
	private Sheet sheet;
	private CellStyle styleHeader;
	private CellStyle styleDate;
	private CellStyle styleDateTime;

	@Override
  protected void doGet(HttpServletRequest q, HttpServletResponse r) {

		ServiceContext sc = WebServiceContextFactory.getServiceContext(q);
		Locale currentLocale = sc.getLocale();
//		ProcessAnalyticsService pas = ServiceLocator.getProcessAnalyticsService2(sc);

		// Parse Parameters
		Long myReportID = Long.parseLong(q.getParameter(PARAM_REPORT_ID));
		Long startIndex = new Long(0);
		Long batchSize = new Long(-1);
		try {
			startIndex = Long.parseLong(q.getParameter(PARAM_START_INDEX));
			batchSize = Long.parseLong(q.getParameter(PARAM_BATCH_SIZE));
		} catch (Exception e) {
		  LOG.debug("Error parsing startIndex or batchSize, using defaults");
		}
		String myFilename = q.getParameter(PARAM_FILENAME);
		if (StringUtils.isBlank(myFilename)) {
			myFilename = "Appian_Data_Export";
		}

		// Assume semicolon separated list of contexts.
		String myContextString = q.getParameter(PARAM_CONTEXT);

		try (OutputStream fileOut = r.getOutputStream()){

			r.setContentType("application/vnd.ms-excel");
			r.setHeader("Content-Disposition", "attachment; filename="+ myFilename + ".xls");

			// Create Workbook and set date formatting
			wb = new HSSFWorkbook();
			sheet = wb.createSheet("ExportedData");
			initWorkbookFormatting(currentLocale);

			// Prepare ReportData
//			ProcessReport report = pas.getProcessReport(myReportID);
//			reportData = report.getData();
//			setReportContext(reportData, myContextString, ((ASLIdentity) sc.getIdentity()).getIdentity());
//			reportData.setBatchSize(batchSize.intValue());
//			reportData.setStartIndex(startIndex.intValue());
//
//			ReportResultPage reportResultPage = pas.getReportPageWithTypedValues(reportData);
//
//			// Write to Workbook Cells
//			int numColumns = writeExcelWorkbook(reportResultPage);
//
//			// Done writing data. Resize columns nicely.
//			for (int i = 0; i < numColumns; i++) {
//				sheet.autoSizeColumn(i);
//			}
			// Write to Output Stream
			wb.write(fileOut);

		} catch (Exception e) {
			LOG.error("Unexpected error writing data to excel file", e);
		}

	}

	/*
	 * Sets date format and header column format.
	 */
	private void initWorkbookFormatting(Locale l_) {
		String dateformat = "dd/mm/yyyy";
		if (l_.equals(Locale.US)) {
			dateformat = "mm/dd/yyyy";
		}
		DataFormat df = wb.createDataFormat();
		styleDateTime = wb.createCellStyle();
		styleDateTime.setDataFormat(df.getFormat(dateformat + " hh:mm"));
		styleDate = wb.createCellStyle();
		styleDate.setDataFormat(df.getFormat(dateformat));

		styleHeader = wb.createCellStyle();
		Font font = wb.createFont();
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		styleHeader.setFont(font);
		styleHeader.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
		styleHeader.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
	}

	/**
	 * Deals with Process Model ID, Process ID, and User Context. If It's User
	 * Context-expecting report and no context is passed in, the default calling
	 * context will be used.
	 *
	 * @param rd_
	 * @param myContextString_
	 * @param scUsername_
	 */
//	private void setReportContext(ReportData rd_, String myContextString_,
//			String scUsername_) {
//
//		int ctxType = rd_.getContextType();
//
//		// Check for user task report, set custom user or Service Context user.
//		if (ctxType == ReportData.CONTEXT_TYPE_ATTRIBUTED_TO_USER
//				|| ctxType == ReportData.CONTEXT_TYPE_TASK_OWNER) {
//
//			String username = (myContextString_ == null) ? scUsername_
//					: myContextString_;
//			if (LOG.isDebugEnabled()) {
//				LOG.debug("Adding  User Context to Task Report: " + username);
//			}
//			LocalObject usercontext = new LocalObject();
//			usercontext.setType(ObjectTypeMapping.TYPE_USER);
//			usercontext.setStringId(username);
//			LocalObject[] ucarray = new LocalObject[1];
//			ucarray[0] = usercontext;
//			rd_.setContext(ucarray);
//
//		} else if (myContextString_ != null) { // Use integer Context Type
//			String[] contextIDs = myContextString_.split(";");
//			Long[] cidArray = new Long[contextIDs.length];
//			for (int i = 0; i < contextIDs.length; i++) {
//				cidArray[i] = Long.parseLong(contextIDs[i]);
//			}
//			int numIDs = cidArray.length;
//
//			if (ctxType == ReportData.CONTEXT_TYPE_PROCESS) {
//				if (LOG.isDebugEnabled()) {
//					LOG.debug("Adding Custom Context to Process Report: "
//							+ numIDs + " IDs");
//				}
//
//				LocalObject[] pcarray = new LocalObject[numIDs];
//				for (int i = 0; i < numIDs; i++) {
//					LocalObject intContext = new LocalObject();
//					intContext.setType(ObjectTypeMapping.TYPE_BPM_PROCESS);
//					intContext.setId(cidArray[i]);
//					pcarray[i] = intContext;
//				}
//				rd_.setContext(pcarray);
//			} else if (ctxType == ReportData.CONTEXT_TYPE_PROCESS_MODEL) {
//				if (LOG.isDebugEnabled()) {
//					LOG.debug("Adding Custom Context to Process Model Report: "
//							+ numIDs + " IDs");
//				}
//				LocalObject[] pmcarray = new LocalObject[numIDs];
//				for (int i = 0; i < numIDs; i++) {
//					LocalObject intContext = new LocalObject();
//					intContext.setType(ObjectTypeMapping.TYPE_BPM_PROCESS_MODEL);
//					intContext.setId(cidArray[i]);
//					pmcarray[i] = intContext;
//				}
//				rd_.setContext(pmcarray);
//			}
//		}
//
//	}

	/**
	 * Write Cell using crazy methods written by SS that i don't understand.
	 * Return total column count.
	 */
//	private int writeExcelWorkbook(ReportResultPage rrp_) {
//		Column[] columns = reportData.getColumns();
//
//		Cell reportStart = ExportHelper.getCell(sheet, "a1");
//		Cell current = reportStart;
//		int displayedColumnCount = 0;
//		// Include Header Row
//		for (Column c : columns) {
//			if (c.getShow()) {
//				displayedColumnCount++;
//				current.setCellValue(c.getName());
//				current.setCellStyle(styleHeader);
//				Cell next = current.getRow().getCell(
//						current.getColumnIndex() + 1);
//				if (next == null) {
//					next = current.getRow().createCell(
//							current.getColumnIndex() + 1);
//				}
//				current = next;
//			}
//		}
//		current = ExportHelper.getCell(sheet, reportStart.getRowIndex() + 1,
//				reportStart.getColumnIndex());
//
//		int rowi = current.getRowIndex();
//		int celli_s = reportStart.getColumnIndex();
//		for (int i = 0; i < rrp_.getResults().length; i++) {
//			HashMap rowData = (HashMap) rrp_.getResults()[i];
//			int celli = celli_s;
//			for (Column c : columns) {
//				if (c.getShow()) {
//					current = ExportHelper.getCell(sheet, rowi, celli);
//					TypedValue tv = (TypedValue) rowData.get(c.getStringId());
//					if (tv.getValue() instanceof String) {
//						current.setCellValue((String) tv.getValue());
//					} else if (tv.getValue() instanceof Timestamp) {
//						current.setCellValue((Timestamp) tv.getValue());
//						current.setCellStyle(styleDateTime);
//					} else if (tv.getValue() instanceof Date) {
//						current.setCellValue((Date) tv.getValue());
//						current.setCellStyle(styleDate);
//					} else if (tv.getValue() instanceof Number) {
//						current.setCellValue(Double.parseDouble(tv.getValue().toString()));
//					} else if (tv.getValue() instanceof String[]) {
//						String[] a = (String[]) tv.getValue();
//						if (a.length == 1) {
//							current.setCellValue(a[0]);
//						} else {
//							current.setCellValue(StringUtils.join(a, ", "));
//						}
//					}
//
//					else if (tv.getValue() instanceof Long[]) {
//						Long[] a = (Long[]) tv.getValue();
//						if (a.length == 1) {
//							current.setCellValue("" + a[0]);
//						} else {
//							current.setCellValue(StringUtils.join(a, ", "));
//						}
//					} else if (tv.getValue() instanceof LocalObject[]) {
//						LocalObject[] a = (LocalObject[]) tv.getValue();
//						if (a.length == 1) {
//							if (a[0].getId() == null) {
//								current.setCellValue(a[0].getStringId());
//							} else {
//								current.setCellValue(a[0].getId());
//							}
//						} else {
//							current.setCellValue(StringUtils.join(a, ", "));
//						}
//					} else if (tv.getValue() == null) {
//						current.setCellValue("");
//					} else {
//						LOG.error("cannot print this type: "+ tv.getValue().getClass().getName());
//					}
//					Cell next = current.getRow().getCell(
//							current.getColumnIndex() + 1);
//					if (next == null) {
//						next = current.getRow().createCell(current.getColumnIndex() + 1);
//					}
//					current = next;
//					celli++;
//					// s.getRow(0).get
//				}
//			}
//			rowi++;
//
//		} // End iteration
//		return displayedColumnCount;
//	}

}
