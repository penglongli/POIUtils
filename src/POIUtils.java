import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class POIUtils {
public static void main(String[] args) throws FileNotFoundException {
		String relativelyPath=System.getProperty("user.dir"); 
	    String fileName ="\\static\\test.xls";
	    File file = new File(relativelyPath+fileName);
	    InputStream inputStream = new FileInputStream(file);
	    /*1����������Ŀ
	    int count = getRecordsCountByInputStream(inputStream,1, true, 1);
	    System.out.print(count);*/
	    readXLSRecords(inputStream, true, 3);
	    
	    //��������Ŀ
	    //int count = getRecordsCountByInputStream(inputStream,1, true, 1);
	    //System.out.print(count);
	}
	
	/**
	 * ͨ��InputStream���������Excel����Ŀ
	 * @param inputStream ������
	 * @param type �ļ����ͣ�0Ϊxls��1Ϊxlsx
	 * @param isHeader �Ƿ��ͷ
	 * @param headerCount ��ͷ����
	 * @return
	 */
	@SuppressWarnings("unused")
	public static int getRecordsCountByInputStream(InputStream inputStream,int type, boolean isHeader, int headerCount){
		int count = 0;
		if(type != 0 && type != 1){
			return count;
		}
		try {
			Workbook wb = null;
			if(type == 0){
				//Excel 2003��ǰ�汾
				wb = new HSSFWorkbook(inputStream);
			}else{
				//Excel 2007�汾
				wb = new XSSFWorkbook(inputStream);
			}
			if(wb == null){
				return count;
			}
			Sheet sheet = wb.getSheetAt(0);
			int begin = sheet.getFirstRowNum();
			int end = sheet.getLastRowNum();
			System.out.println(end);
			if(isHeader){
				begin += headerCount;
			}
			for(int i = begin;i<=end;i++){
				if(sheet.getRow(i) == null){
					continue;
				}
				count++;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return count;
	}
	/**
	 * ��ȡ2003�汾����Excel��¼��ע������������˵�һ��Sheet
	 * @param inputStream ������
	 * @param isHeader �Ƿ������ͷ
	 * @param headerCount ��ͷ����
	 * @return ���ز�������ͷ����Ϣ�б�
	 */
	public static HashMap<String,List<Product>> readXLSRecords(InputStream inputStream, boolean isHeader, int headerCount){
		//�洢���ݸ�ʽ��String�洢�ŵ��ţ�List<HashMap>�洢��Ʒ��š���Ʒ�����б�
		HashMap<String,List<Product>> paramMapList = new HashMap<String,List<Product>>();
		
		try {
			HSSFWorkbook hswb = new HSSFWorkbook(inputStream);
			HSSFSheet hsSheet = hswb.getSheetAt(0);
			
			int begin = hsSheet.getFirstRowNum();
			//����б�ͷ��������ͷ
			if(isHeader){
				begin += headerCount;
			}
			
			HSSFRow row = null;
			int colNumber = 0;
			List<Product> paramMap = new ArrayList<Product>();
			//����Excel��
			for(int i = begin;i<hsSheet.getLastRowNum();i++){
				row = hsSheet.getRow(i);
				colNumber = row.getPhysicalNumberOfCells();
				
				if(row != null){
					//����Excel��
					Product product = new Product();
					for(int j = 1;j < colNumber; j++){
						HSSFCell cell = row.getCell(j);
						
						row.getCell(j).setCellType(Cell.CELL_TYPE_STRING);
						System.out.print(cell.getStringCellValue()+" ");
					}
				}
				System.out.println();
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		return paramMapList;
	}
	/**
	 * ��ȡ2007�汾��Excel��¼
	 * @param inputStream
	 * @param isHeader
	 * @param headerCount
	 * @return
	 */
	public static List<HashMap<Integer, Object>> readXLSXRecords(InputStream inputStream, boolean isHeader, int headerCount){
		List<HashMap<Integer, Object>> paramMapList = new ArrayList<HashMap<Integer,Object>>();
		
		try {
			XSSFWorkbook xswb = new XSSFWorkbook(inputStream);
			XSSFSheet xsSheet = xswb.getSheetAt(0);
			
			int begin = xsSheet.getFirstRowNum();
			//
			if(isHeader){
				begin += headerCount;
			}
			
			XSSFRow row = null;
			int colNumber = 0;
			HashMap<String, Object> paramMap = new HashMap<String, Object>();
			for(int i = begin;i < xsSheet.getLastRowNum();i++){
				row = xsSheet.getRow(i);
				
				if(row != null){
					colNumber = row.getPhysicalNumberOfCells();
					paramMap.clear();
					for(int j = 0;j < colNumber;j++){
						
						XSSFCell cell = row.getCell(j);
						row.getCell(j).setCellType(Cell.CELL_TYPE_STRING);
						System.out.print(cell.getStringCellValue()+"\t");
					}
					System.out.println();
				}
				
			}
			
		} catch (IOException e) {
			e.printStackTrace();
		}
		return paramMapList;
	}
}
