import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class PcMain {
	public static String url="https://book.douban.com";
	public static void main(String[] args) throws IOException {	
		//��ǰ��ҳ���ִӸߵ��͵����� �ٽ���ɸѡ
	    List list=getBookbyPc(1);
	    List list2=getBookbyPc(2);
	    List list3=getBookbyPc(3);
	    List list4=getBookbyPc(4);
	    list.addAll(list2);
	    list3.addAll(list);
	    list4.addAll(list3);
	    Collections.sort(list4, new Comparator<Books>() {      	  
            @Override  
            public int compare(Books o1, Books o2) {  
                // ����ѧ�������������������  
                if (Double.parseDouble(((Books)o1).getGrade()) > Double.parseDouble(((Books)o2).getGrade())) {  
                    return -1;  
                }  
                return 1;  
            }  
        });       
		exportExcel(list4);
     }
	/**
	 * ��ҳ������ÿҳ��limit�� ò���ں�̨д��Ϊ20��  startΪ��ʼ�±�  typeΪ������� SΪ����������  ÿ����20������
	 * @param page ҳ��
	 * @return
	 * @throws IOException
	 */
	public static List getBookbyPc(Integer page) throws IOException {
		//������ҳ��ҳ�����֪ ���¹��� �����ʼ�±�
		int start=(page*2-2)*10;	
		long startTime = System.currentTimeMillis();
		System.out.println("��ʼ�������ݡ���������URLΪ:https://book.douban.com/tag/%E7%BC%96%E7%A8%8B?start="+start+"&type=S");
		//��ȡ�༭�Ƽ�ҳ
        Document document=Jsoup.connect("https://book.douban.com/tag/%E7%BC%96%E7%A8%8B?start="+start+"&type=S")
                //ģ���������
                .userAgent("Mozilla/4.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0)")
                .get();
        ////System.out.println(document);
        Elements main=document.select(".subject-item");
       // //System.out.println(main);
        List<Books> list=new ArrayList<Books>();
        if(main.size()>0) {
        	for(Element obj : main) {
        		//�ж����������Ƿ񳬹�1000
        		Integer pcnumber = getNumber(obj.select(".info").select(".pl").text());
        		if(pcnumber>1000) {
		            	//����ʵ�������
		            	Books book=new Books();
		        		book.setGradeNumber(pcnumber);
		        		book.setBookName(obj.select(".info").select("h2").text());
		        		//System.out.print("����:"+obj.select(".info").select("h2").text());
		        		book.setGrade(obj.select(".info").select(".rating_nums").text());
		        		//System.out.print("����:"+obj.select(".info").select(".rating_nums").text());				        		
		        		//System.out.println("����:"+getNumber(obj.select(".info").select(".pl").text()));
		        		//�鿴��ҳ����--�ָ��ַ���
		        		String[] str=obj.select(".info").select(".pub").text().split("/");		
		        		if(str.length>0) {
		        			//��ʽ��ͬ ����
		        			if(str.length==5) {
		        				book.setAuthor(str[0]);
		            			////System.out.print("����:"+str[0]);
		            			book.setPublishHouse(str[2]);
		            			////System.out.print("������:"+str[2]);
		            			book.setPublishDate(str[3]);
		            			////System.out.print("��������:"+str[3]);
		            			book.setPrice(str[4]);
		            			////System.out.println("�۸�:"+str[4]);  
		        			}else if(str.length==4) {
		        				book.setAuthor(str[0]);
		        				////System.out.print("����:"+str[0]);
		        				book.setPublishHouse(str[1]);
		            			////System.out.print("������:"+str[1]);
		            			book.setPublishDate(str[2]);
		            			//System.out.print("��������:"+str[2]);
		            			book.setPrice(str[3]);
		            			//System.out.println("�۸�:"+str[3]); 
		        			}
		            		
		            	 }		
		        	list.add(book);
        		}


        	}     	
        }
        long endTime = System.currentTimeMillis();
        
        System.out.println("���к�ʱ"+(endTime-startTime)%1000+"��");
    	System.out.println("������ȡ�ɹ�����������");
		return list;
	}
	//��ȡ�ַ����е�����
	public static Integer getNumber(String str) {
		// str = "xxx��47297��33";
		String regex = "\\d*";
		Pattern p = Pattern.compile(regex);

		Matcher m = p.matcher(str);
        String number="";
		while (m.find()) {
		if (!"".equals(m.group()))
		
		   number= m.group();
		}
		return Integer.parseInt(number);
	}
	/**
	 * �������ݵ�excel�ķ���
	 * @throws FileNotFoundException 
	 */
    public static void exportExcel(List<Books> list) throws FileNotFoundException {
    	String[] biaotou=new String[] {"���","����","����","��������","����","������","��������","�۸�"};
    	//��ط��Լ��޸�
    	OutputStream out=new FileOutputStream(new File("D://book.xls"));
    	HSSFWorkbook workbook=new HSSFWorkbook();
    	HSSFSheet sheet= workbook.createSheet();
    	HSSFRow row=sheet.createRow(0);
    	//�����ͷ
    	for(int i=0;i<biaotou.length;i++) {
    		HSSFCell cell=row.createCell(i);
    		cell.setCellValue(biaotou[i]);		
    	}
    	//��������
    	if(list.size()>0) {
    		Integer indexId=1;
    		for(int j=0;j<list.size();j++) {
    			if(j<=39) {  				
    				HSSFRow currentRow=sheet.createRow(j+1);
    				Books currentBook=list.get(j);
    				currentRow.createCell(0).setCellValue(String.valueOf(indexId));
    				currentRow.createCell(1).setCellValue(currentBook.getBookName());
    				currentRow.createCell(2).setCellValue(currentBook.getGrade());
    				currentRow.createCell(3).setCellValue(currentBook.getGradeNumber());
    				currentRow.createCell(4).setCellValue(currentBook.getAuthor());
    				currentRow.createCell(5).setCellValue(currentBook.getPublishHouse());
    				currentRow.createCell(6).setCellValue(currentBook.getPublishDate());
    				currentRow.createCell(7).setCellValue(currentBook.getPrice());	
    				indexId++;
    			}
    		}
    	} 	
    	try {
			workbook.write(out);
			out.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

    	
    	
    	
    	
    }
}
