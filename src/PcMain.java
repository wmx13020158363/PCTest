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
/**
 *
 * @author Administrator
 *
 */
public class PcMain {
	public static String url="https://book.douban.com";
	public static void main(String[] args) throws IOException {	
		//拿前三页评分从高到低的数据 再进行筛选
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
                // 按照学生的年龄进行升序排列  
                if (Double.parseDouble(((Books)o1).getGrade()) > Double.parseDouble(((Books)o2).getGrade())) {  
                    return -1;  
                }  
                return 1;  
            }  
        });       
		exportExcel(list4);
     }
	/**
	 * 网页分析：每页的limit数 貌似在后台写死为20条  start为起始下标  type为排序类别 S为按评分排序  每次拿20条数据
	 * @param page 页数
	 * @return
	 * @throws IOException
	 */
	public static List getBookbyPc(Integer page) throws IOException {
		//分析网页分页排序得知 如下规律 算出起始下标
		int start=(page*2-2)*10;	
		long startTime = System.currentTimeMillis();
		System.out.println("开始爬虫数据。。。。。URL为:https://book.douban.com/tag/%E7%BC%96%E7%A8%8B?start="+start+"&type=S");
		//获取编辑推荐页
        Document document=Jsoup.connect("https://book.douban.com/tag/%E7%BC%96%E7%A8%8B?start="+start+"&type=S")
                //模拟火狐浏览器
                .userAgent("Mozilla/4.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0)")
                .get();
        ////System.out.println(document);
        Elements main=document.select(".subject-item");
       // //System.out.println(main);
        List<Books> list=new ArrayList<Books>();
        if(main.size()>0) {
        	for(Element obj : main) {
        		//判断评价人数是否超过1000
        		Integer pcnumber = getNumber(obj.select(".info").select(".pl").text());
        		if(pcnumber>1000) {
		            	//定义实体类对象
		            	Books book=new Books();
		        		book.setGradeNumber(pcnumber);
		        		book.setBookName(obj.select(".info").select("h2").text());
		        		//System.out.print("书名:"+obj.select(".info").select("h2").text());
		        		book.setGrade(obj.select(".info").select(".rating_nums").text());
		        		//System.out.print("评分:"+obj.select(".info").select(".rating_nums").text());				        		
		        		//System.out.println("人数:"+getNumber(obj.select(".info").select(".pl").text()));
		        		//查看网页内容--分割字符串
		        		String[] str=obj.select(".info").select(".pub").text().split("/");		
		        		if(str.length>0) {
		        			//格式不同 处理
		        			if(str.length==5) {
		        				book.setAuthor(str[0]);
		            			////System.out.print("作者:"+str[0]);
		            			book.setPublishHouse(str[2]);
		            			////System.out.print("出版社:"+str[2]);
		            			book.setPublishDate(str[3]);
		            			////System.out.print("出版日期:"+str[3]);
		            			book.setPrice(str[4]);
		            			////System.out.println("价格:"+str[4]);  
		        			}else if(str.length==4) {
		        				book.setAuthor(str[0]);
		        				////System.out.print("作者:"+str[0]);
		        				book.setPublishHouse(str[1]);
		            			////System.out.print("出版社:"+str[1]);
		            			book.setPublishDate(str[2]);
		            			//System.out.print("出版日期:"+str[2]);
		            			book.setPrice(str[3]);
		            			//System.out.println("价格:"+str[3]); 
		        			}
		            		
		            	 }		
		        	list.add(book);
        		}


        	}     	
        }
        long endTime = System.currentTimeMillis();
        
        System.out.println("运行耗时"+(endTime-startTime)%1000+"秒");
    	System.out.println("数据爬取成功。。。。。");
		return list;
	}
	//获取字符串中的数字
	public static Integer getNumber(String str) {
		// str = "xxx第47297章33";
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
	 * 到处数据到excel的方法
	 * @throws FileNotFoundException 
	 */
    public static void exportExcel(List<Books> list) throws FileNotFoundException {
    	String[] biaotou=new String[] {"序号","书名","评分","评价人数","作者","出版社","出版日期","价格"};
    	//这地方自己修改
    	OutputStream out=new FileOutputStream(new File("D://book.xls"));
    	HSSFWorkbook workbook=new HSSFWorkbook();
    	HSSFSheet sheet= workbook.createSheet();
    	HSSFRow row=sheet.createRow(0);
    	//定义表头
    	for(int i=0;i<biaotou.length;i++) {
    		HSSFCell cell=row.createCell(i);
    		cell.setCellValue(biaotou[i]);		
    	}
    	//数据填入
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
