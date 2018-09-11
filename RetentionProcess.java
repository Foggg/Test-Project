package retention;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Assert;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import main.JDBC;
import main.Order;
import main.OrderRunner;

@RunWith(OrderRunner.class)

//https://mantis.dyninno.net/view.php?id=125773

public class RetentionProcess extends JDBC { 
	
	Workbook wb = new HSSFWorkbook();
	
	Connection conn = null;
	Statement stmt = null;
	
	@Test  					
	@Order(order = 1)
	
	public void upsale7() {
		
		System.out.println("Upsales 7");
		
		try {
			Class.forName("org.postgresql.Driver");
			conn = DriverManager.getConnection(DB_URL, USER, PASS);
			System.out.println("Connected to " + DB_URL + " successfully...");

			stmt = conn.createStatement();

			Sheet sheet1= wb.createSheet("Upsale7");
			Sheet sheet2= wb.createSheet("Wrong7");
			
			ArrayList<String> clients = new ArrayList<String>();
			ArrayList<String> limits = new ArrayList<String>();
			ArrayList<String> visibleLimits = new ArrayList<String>();
			ArrayList<String> accept_marketings = new ArrayList<String>();
			ArrayList<String> no_communications = new ArrayList<String>();
			ArrayList<String> active_loans = new ArrayList<String>();
			ArrayList<String> id = new ArrayList<String>();
			ArrayList<String> ip = new ArrayList<String>();
			ArrayList<String> phone = new ArrayList<String>();
			ArrayList<String> region = new ArrayList<String>();
			ArrayList<String> address = new ArrayList<String>();
			ArrayList<String> passport = new ArrayList<String>();
			ArrayList<String> blackListFull = new ArrayList<>();
			ArrayList<String> r_blacklisted = new ArrayList<String>();
			ArrayList<String> black_list_ip = new ArrayList<String>();
			ArrayList<String> black_list_phone = new ArrayList<String>();
			ArrayList<String> black_list_pass = new ArrayList<String>();
			ArrayList<String> client_bl_ips = new ArrayList<String>();
			ArrayList<String> client_bl_phones = new ArrayList<String>();
			ArrayList<String> client_bl_passports = new ArrayList<String>();
			ArrayList<String> unhappy_clients = new ArrayList<String>();
			ArrayList<String> crm_comm_id = new ArrayList<String>();
			ArrayList<String> true_false1 = new ArrayList<String>();
			ArrayList<String> true_false3 = new ArrayList<String>();
			ArrayList<String> true_false5 = new ArrayList<String>();
			ArrayList<String> true_false6 = new ArrayList<String>();
			ArrayList<String> clients_tasks_upsale5 = new ArrayList<String>();
			ArrayList<String> bl_clients_tasks_upsale5 = new ArrayList<String>();
			ArrayList<String> clients_late = new ArrayList<String>();
			ArrayList<String> clients_task = new ArrayList<String>();
			ArrayList<String> wait_approval = new ArrayList<String>();
			ArrayList<String> page = new ArrayList<String>();
			ArrayList<String> latess = new ArrayList<String>();
			ArrayList<String> full_check = new ArrayList<String>();
			ArrayList<String> passport_clients = new ArrayList<String>();
			ArrayList<String> last_full_check_date = new ArrayList<String>();
			ArrayList<String> pass_issue = new ArrayList<String>();
			ArrayList<String> last_full_check_date_full = new ArrayList<String>();
			
			//Checking all clients with tasks
			String select = "select r_client_id, created from crm_task where crm_task_type_id = '39' and created::date=current_date::date - interval '1 day';";
			
			//Creating list of clients (Point 1)
			String select1_1 = "SELECT distinct r_client_id FROM r_loan WHERE PAYMENT_DATE::date = (current_date::date - interval '8 days') and status = 'Completed' and r_client_id not in (select r_client_id from r_loan where loan_date >= current_date::date - interval '8 days')";
			String select1_2 = "SELECT c1.r_loan_id, c1.r_client_id, c1.updated as complete_date, 'Canceled' as status,rc.mobile_phone,rc.client_no FROM (SELECT r_client_id, max(created) as created, max(updated) as updated, max(r_loan_id) as r_loan_id FROM r_loan rl WHERE rl.status = 'Canceled' GROUP BY r_client_id) c1 LEFT JOIN r_client rc on rc.r_client_id=c1.r_client_id left JOIN (SELECT r_client_id, max(created) as created FROM (SELECT rank() OVER (PARTITION BY r_client_id ORDER BY created) AS rank, * FROM r_loan WHERE status NOT IN ('Voided', 'Canceled') ORDER BY r_client_id) ranked GROUP BY r_client_id) other on other.r_client_id=c1.r_client_id and c1.created>other.created WHERE other.r_client_id is not NULL AND (-1)*extract(DAY FROM c1.updated-current_date) = 7";
			String select1_3 = "SELECT distinct r_client_id FROM r_loan WHERE status = 'Completed' and r_client_id IN (select sta.r_client_id from (SELECT r_client_id, max(r_application_id) as rid FROM r_application GROUP BY r_client_id) sta LEFT JOIN r_application ra on sta.rid=ra.r_application_id LEFT JOIN (SELECT r_client_id, max(rank) as rank FROM (SELECT rank() OVER (PARTITION BY r_client_id ORDER BY loan_date) AS rank, * FROM r_loan WHERE status NOT IN ('Voided', 'Canceled') ORDER BY r_client_id) as rakn GROUP BY r_client_id) rank on rank.r_client_id=sta.r_client_id LEFT JOIN r_client rc ON rc.r_client_id = sta.r_client_id WHERE ra.status='EXPIRED' AND  rank.rank>=1 AND (-1)*extract(DAY FROM ra.updated-current_date) = 7);";
			
			ResultSet rs = stmt.executeQuery(select1_1);
			while (rs.next()) {
				String client = rs.getString("r_client_id");
				clients.add(client);
			}
			
			ResultSet rs2 = stmt.executeQuery(select1_2);
			while (rs2.next()) {
				String client = rs2.getString("r_client_id");
				clients.add(client);
			}
			
			ResultSet rs3 = stmt.executeQuery(select1_3);
			while (rs3.next()) {
				String client = rs3.getString("r_client_id");
				clients.add(client);
			}

			ResultSet rs444 = stmt.executeQuery(select);
			while (rs444.next()) {
				String task = rs444.getString("r_client_id");
				clients_task.add(task);
			}
			
			//Checking clients with "No Communication"
			for (String c : clients) {
				String nc = "select * from r_client where IS_NO_COMMUNICATION = 't' and r_client_id ='" + c + "'";
				ResultSet rs44 = stmt.executeQuery(nc);
				while (rs44.next()) {
					String no_communication = rs44.getString("r_client_id");
					no_communications.add(no_communication);
				}
			}

			//Checking clients with Active loans
			for (String c : clients) {
				String active = "select * from r_loan where status = 'Active' and r_client_id ='" + c + "'and created::date != current_date";
				ResultSet rs5 = stmt.executeQuery(active);
				while (rs5.next()) {
					String active_loan = rs5.getString("r_client_id");
					active_loans.add(active_loan);
				}
			}
			
			//Checking clients with "Late 30+"
			for (String c : clients) {
				String lates = "SELECT rl.r_client_id, d.total_delay FROM ( SELECT max(r_loan_id) as r_loan_id, r_client_id FROM r_loan WHERE status in ('Completed') GROUP BY r_client_id ) as t LEFT JOIN r_loan rl using (r_loan_id) LEFT JOIN (SELECT (sum(late_days_actual)+sum(late_days_saved))::integer as total_delay, r_loan_id FROM 	r_loan_history GROUP BY r_loan_id) as d USING (r_loan_id) LEFT JOIN	r_client rc ON rc.r_client_id = t.r_client_id WHERE d.total_delay > '30' and rl.r_client_id = '" + c + "'";
				ResultSet lt = stmt.executeQuery(lates);
				while (lt.next()) {
					String late = lt.getString("r_client_id");
					clients_late.add(late);
				}
			}

			//Checking clients with full check 85 days +
			for (String c : clients) {
				String full_checked = "select * from r_client where full_check_date < current_date::date - interval '85 day' and r_client_id ='" + c + "'";
				ResultSet lt = stmt.executeQuery(full_checked);
				while (lt.next()) {
					String full_checking = lt.getString("r_client_id");
					full_check.add(full_checking);
				}
			}
			
			//Checking clients with "Wait ..." application 
			for (String c : clients) {
				String wait = "select * from r_application where status in ('WAIT_WEB_APPROVAL','WAIT_CHANGES_APPROVAL') and r_client_id = '" + c + "'";
				ResultSet wt = stmt.executeQuery(wait);
				while (wt.next()) {
					String wait_app = wt.getString("r_client_id");
					wait_approval.add(wait_app);
				}
			}
			
			//Checking "Late" clients 
			for (String c : clients) {
				String latee = "select * from r_loan where status = 'Late' and r_client_id ='" + c	+ "'and created::date != current_date";
				ResultSet wt = stmt.executeQuery(latee);
				while (wt.next()) {
					String late_client = wt.getString("r_client_id");
					latess.add(late_client);
				}
			}
			
			//Checking clients from Black Lists
			for (String c : clients) {
				String r_blacklist = "select * from r_blacklist where r_client_id ='" + c + "'";
				ResultSet rs_bl = stmt.executeQuery(r_blacklist);
				while (rs_bl.next()) {
					String blacklisted = rs_bl.getString("r_client_id");
					r_blacklisted.add(blacklisted);
					Collections.addAll(blackListFull, blacklisted);
				}
			}

			for (String c : clients) {
				String r_blacklist_ip = "select IP_ADDRESS from r_application where r_client_id ='" + c + "'";
				ResultSet rs_bl_ip = stmt.executeQuery(r_blacklist_ip);
				while (rs_bl_ip.next()) {
					String client_ip = rs_bl_ip.getString("IP_ADDRESS");
					ip.add(client_ip);
				}
			}
			
			for (String c : clients) {
				String bl = "select * from r_client where r_client_id ='" + c + "'";
				ResultSet rs6 = stmt.executeQuery(bl);
				while (rs6.next()) {
					String client_id = rs6.getString("R_CLIENT_ID");
					String client_phone = rs6.getString("MOBILE_PHONE");
					String client_region = rs6.getString("REG_COUNTRY_ID");
					String client_address = rs6.getString("DECL_ADDRESS_ID");
					String client_passport = rs6.getString("PASSPORT_NO");

					id.add(client_id);
					phone.add(client_phone);
					region.add(client_region);
					address.add(client_address);
					passport.add(client_passport.replaceAll("\\D+",""));
				}
			}
			
			for (String i : ip) {
				String select_ip = "select * from dms_blacklist_ip where ip ='" + i + "' and active_till is not NULL and active_till >= current_date";
				ResultSet rs7 = stmt.executeQuery(select_ip);
				while (rs7.next()) {
					String bl_ip = rs7.getString("ip");
					black_list_ip.add(bl_ip);
				}
			}

			for (String i2 : black_list_ip) {
				String select_ip2 = "select r_client_id from r_application where IP_ADDRESS ='" + i2 + "' order by created desc limit 1";
				ResultSet rs8 = stmt.executeQuery(select_ip2);
				while (rs8.next()) {
					String client_bl_ip = rs8.getString("r_client_id");
					client_bl_ips.add(client_bl_ip);
					Collections.addAll(blackListFull, client_bl_ip);
				}
			}

			for (String ph : phone) {
				String select_phone = "select * from dms_blacklist_phones where phone ='" + ph + "' and active_till is not NULL and active_till >= current_date";
				ResultSet rs9 = stmt.executeQuery(select_phone);
				while (rs9.next()) {
					String bl_phone = rs9.getString("phone");
					black_list_phone.add(bl_phone);
				}
			}

			for (String ph2 : black_list_phone) {
				String select_phone2 = "select r_client_id from r_client where mobile_phone ='" + ph2 + "'";
				ResultSet rs10 = stmt.executeQuery(select_phone2);
				while (rs10.next()) {
					String client_bl_phone = rs10.getString("r_client_id");
					client_bl_phones.add(client_bl_phone);
					Collections.addAll(blackListFull, client_bl_phone);
				}
			}

			for (String pass : passport) {
				String select_pass = "select * from dms_blacklist_passports where passport_no ='" + pass + "' and active_till is not NULL and active_till >= current_date";
				ResultSet rs11 = stmt.executeQuery(select_pass);
				while (rs11.next()) {
					String bl_pass = rs11.getString("passport_no");
					black_list_pass.add(bl_pass);
				}
			}

			for (String pass2 : black_list_pass) {
				String select_pass2 = "select r_client_id from r_client where passport_no ='" + pass2 + "'";
				ResultSet rs12 = stmt.executeQuery(select_pass2);
				while (rs12.next()) {
					String client_bl_pass = rs12.getString("r_client_id");
					client_bl_passports.add(client_bl_pass);
					Collections.addAll(blackListFull, client_bl_pass);
				}
			}

			for (String up5 : blackListFull) {
				String blacklistTaskUpsale5 = "select * from crm_task where crm_task_type_id = '39' and r_client_id ='"
						+ up5 + "' and created::date = current_date::date - interval '1 day'";
				ResultSet blc = stmt.executeQuery(blacklistTaskUpsale5);
				while (blc.next()) {
					String task_for_bl_clients = blc.getString("r_client_id");
					bl_clients_tasks_upsale5.add(task_for_bl_clients);
				}
			}
			
			for (String up5 : clients) {
				String tasks = "select * from crm_task where crm_task_type_id = '39' and r_client_id ='"
						+ up5 + "' and created::date = current_date::date - interval '1 day'";
				ResultSet tk = stmt.executeQuery(tasks);
				while (tk.next()) {
					String task_for_clients = tk.getString("r_client_id");
					clients_tasks_upsale5.add(task_for_clients);
				}
			}

			//Checking "Unhappy clients"
			for (String c : clients) {
				String unhappy = "select * from crm_task where crm_task_status_id = '515' and r_client_id ='" + c
						+ "'  and (CRM_TASK_STATUS_ID = '513' or crm_task_status_id = '514' or crm_task_status_id = '515')";
				ResultSet rs13 = stmt.executeQuery(unhappy);
				while (rs13.next()) {
					String unhappy_client = rs13.getString("r_client_id");
					unhappy_clients.add(unhappy_client);
				}
			}

			//Bad Passport
			for (String c : clients) {
				String twenty = "SELECT r_client_id, PASSPORT_ISSUED_DATE, DOB from r_client where (EXTRACT(DAY FROM ((r_client.dob + INTERVAL '20 years') - r_client.dob)) <= current_date - r_client.dob AND EXTRACT(DAY FROM ((r_client.dob + INTERVAL '20 years') - r_client.dob)) > passport_issued_date - r_client.dob) and r_client_id = '" + c	+ "'";
				ResultSet rs20 = stmt.executeQuery(twenty);
				while (rs20.next()) {
					String client_passport20 = rs20.getString("r_client_id");
					passport_clients.add(client_passport20);
				}
			}
			
			for (String c : clients) {
				String fourtyfive = "SELECT r_client_id, PASSPORT_ISSUED_DATE, DOB from r_client where (EXTRACT(DAY FROM ((r_client.dob + INTERVAL '45 years') - r_client.dob)) <= current_date - r_client.dob AND EXTRACT(DAY FROM ((r_client.dob + INTERVAL '45 years') - r_client.dob)) > passport_issued_date - r_client.dob) and r_client_id = '" + c	+ "'";
				ResultSet rs45 = stmt.executeQuery(fourtyfive);
				while (rs45.next()) {
					String client_passport45 = rs45.getString("r_client_id");
					passport_clients.add(client_passport45);
				}
			}

			//Remove clients from the list (Point 2)
			for (Object obj : clients) {
				clients_task.remove(obj);
			}	
			for (Object obj : no_communications) {
				clients.remove(obj);
			}
			for (Object obj : active_loans) {
				clients.remove(obj);
			}
			for (Object obj : clients_late) {
				clients.remove(obj);
			}
			for (Object obj : unhappy_clients) {
				clients.remove(obj);
			}
			for (Object obj : blackListFull) {
				clients.remove(obj);
			}
			for (Object obj : wait_approval) {
				clients.remove(obj);
			}
			for (Object obj : latess) {
				clients.remove(obj);
			}
			for (Object obj : full_check) {
				clients.remove(obj);
			}
			for (Object obj : passport_clients) {
				clients.remove(obj);
			}
			
			System.out.println("Total clients with task, but not in list "+clients_task);
			
			for (String t : clients_task){
				String last_full_check = "select (current_date-FULL_CHECK_DATE::date) from r_client where r_client_id = '" + t + "'";
				ResultSet rs999 = stmt.executeQuery(last_full_check);
				while (rs999.next()) {
					String lfc = rs999.getString("?COLUMN?");
					last_full_check_date.add(lfc);
				}
			}
			
			for (String c : clients){
				String last_full_check = "select (current_date-FULL_CHECK_DATE::date) from r_client where r_client_id = '" + c + "'";
				ResultSet rs999 = stmt.executeQuery(last_full_check);
				while (rs999.next()) {
					String lfc = rs999.getString("?COLUMN?");
					last_full_check_date_full.add(lfc);
				}
			}
			
			for (String t : clients_task){
				String pass_issue_date = "SELECT r_client_id, PASSPORT_ISSUED_DATE, DOB from r_client where (EXTRACT(YEAR FROM PASSPORT_ISSUED_DATE) - EXTRACT(YEAR FROM DOB) < '20')  and EXTRACT(YEAR FROM DOB) > '20' and EXTRACT(YEAR FROM current_date) - EXTRACT(YEAR FROM PASSPORT_ISSUED_DATE)> '20' and EXTRACT(YEAR FROM DOB) = EXTRACT(YEAR FROM PASSPORT_ISSUED_DATE) and r_client_id ='"+ t +"'";
				ResultSet rs998 = stmt.executeQuery(pass_issue_date);
				while (rs998.next()) {
					String psis = rs998.getString("PASSPORT_ISSUED_DATE");
					pass_issue.add(psis);
				}
				if (pass_issue.size() > 0) {
					System.out.println(t + " has bad pass");
					true_false1.add("NOK");
				} else {
					System.out.println("OK for " + t);
					true_false1.add("OK");
				}
			}
			
			System.out.println("Total clients After Cut "+clients.size());
			
			//Checking Limit for Clients (Point 3)
			for (String c : clients) {
				driver.get("http://dms.fin.dyninno.net/api/client/" + c + "/nextlimit");
				String limit = driver.findElement(By.xpath("/html/body/pre")).getText();
				limit = limit.substring(0, Math.min(limit.length(), 39));
				limit = limit.substring(34);
				limits.add(limit);
			}

			//Checking that limit is visible on client's page and first page is "Мои Акции"(Point 4, Point 5) 
			driver.get(dyninno);
			driver.manage().window().maximize();
			driver.findElement(By.name("login")).sendKeys("k.smirnovs");
			driver.findElement(By.name("pwd")).sendKeys("zaq1@WSX");
			driver.findElement(By.xpath("//*[@id='ajaxModal']/div/div/div/form/div[2]/span/button")).click();

			for (String c : clients) {
				driver.findElement(By.linkText("Clients")).click();
				driver.findElement(By.id("id")).sendKeys(c);
				driver.findElement(By.id("btn-search")).click();
				sleep(1);
				driver.findElement(By.xpath("//*[@id='tbl']/tbody/tr/td[1]/a")).click();
				driver.findElement(By.xpath("//*[@id='content']/div/div[1]/section[1]/header/ul/li[2]/a")).click();
				driver.findElement(By.id("hijack")).click();
				sleep(1);
				ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());
				driver.switchTo().window(tabs.get(1));
				sleep(1);
				 WebElement element = driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div/div[2]/div[1]/h1"));
				 String test = element.getText();
				 if (test != "Мои акции") {
				 System.out.println("Page Мои Акции doesn't open for "+c);
				 page.add(test);
				 }
				 else{
					 System.out.println(c+ "client is on Мои Акции page");
					 page.add(test); 
				 }
				driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div/div[1]/nav/a[3]")).click();
				sleep(1);
				WebElement element2 = driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div/div[2]/div[2]"));
				String test2 = element2.getText();
				System.out.println(c + ": " + test2);
				test2 = (test2.substring(test2.length() - 49));
				String noLimit = "Вы получите информацию о специальном предложении.";
				if (test2.contentEquals(noLimit)) {
					String fail = "0";
					visibleLimits.add(fail);
					driver.findElement(By.className("cabinet-logout-link")).click();
					driver.close();
					driver.switchTo().window(tabs.get(0));
				} else {
					driver.get("http://dms.fin.dyninno.net/api/client/" + c + "/nextlimit");
					String limit = driver.findElement(By.xpath("/html/body/pre")).getText();
					limit = limit.substring(0, Math.min(limit.length(), 39));
					limit = limit.substring(34);
					visibleLimits.add(limit);
					driver.close();
					driver.switchTo().window(tabs.get(0));
				}
			}

			//Checking clients with Accept Marketing = false
			for (String c : clients) {
				String am = "select * from r_client where ACCEPT_MARKETING = 'f' and r_client_id ='" + c + "'";
				ResultSet rs14 = stmt.executeQuery(am);
				while (rs14.next()) {
					String accept_marketing = rs14.getString("r_client_id");
					accept_marketings.add(accept_marketing);
				}
			}
			
			System.out.println(accept_marketings);
			
			//Remove clients with Accept Marketing = false (Point 6)
//			for (Object obj : accept_marketings) {
//				clients.remove(obj);
//			}
			
			//if marketing accepted
			for (String c : clients) {
				if (accept_marketings.contains(c)) {
					System.out.println(c + " Marketing Accepted = false");
					true_false6.add("N");
				} else {
					System.out.println(c + " Marketing Accepted = true");
					true_false6.add("Y");
				}
			}
			
			// if sent (Point 7)
			for (String c : clients) {
				crm_comm_id.clear();
				String crm = "select created from crm_communication where r_client_id ='" + c
						+ "' and (MESSAGE_IDENTIFIER = 'RETENTION - EMAIL 2' or MESSAGE_IDENTIFIER = 'RETENTION - 2 SMS') and created::date = current_date::date - interval '1 day'";
				ResultSet rs15 = stmt.executeQuery(crm);
				while (rs15.next()) {
					String comm_id = rs15.getString("created");
					crm_comm_id.add(comm_id);
				}
				if (crm_comm_id.size() != 2) {
					System.out.println("No record for " + c + " created on current date");
					true_false3.add("Not Sent");
				} else {
					System.out.println("SMS and Email were sent for " + c);
					true_false3.add("Sent");
				}
			}
			
			//if task (Point 8)
			for (String c : clients) {
				if (clients_tasks_upsale5.contains(c)) {
					System.out.println(c + " has Task");
					true_false5.add("Task created");
				} else {
					System.out.println(c + " hasn't Task");
					true_false5.add("Task doesn't created");
				}
			}

			//Creating xls Report
			for (int RowNum = 0; RowNum < clients.size(); RowNum++) {
				Row row2 = sheet1.createRow(RowNum);
				Cell cell1 = row2.createCell(0);
				Cell cell2 = row2.createCell(1);
				Cell cell3 = row2.createCell(2);
				Cell cell4 = row2.createCell(3);
				Cell cell5 = row2.createCell(4);
				Cell cell6 = row2.createCell(5);
				Cell cell7 = row2.createCell(6);
				cell1.setCellValue("Client Id: "+clients.get(RowNum));
				cell2.setCellValue("Limit: "+limits.get(RowNum));
				cell3.setCellValue(visibleLimits.get(RowNum));
				cell4.setCellValue(true_false3.get(RowNum));
				cell5.setCellValue(true_false5.get(RowNum));
				cell6.setCellValue("Last full check: "+last_full_check_date_full.get(RowNum)+" days ago");
				cell7.setCellValue("Marketing Accepted: "+true_false6.get(RowNum));
			}
			
			for (int RowNum = 0; RowNum < clients_task.size(); RowNum++) {
				Row row1 = sheet2.createRow(RowNum);
				Cell cell1 = row1.createCell(0);
				Cell cell2 = row1.createCell(1);
				Cell cell3 = row1.createCell(2);
				cell1.setCellValue("Client Id: "+clients_task.get(RowNum));
				cell2.setCellValue("Last full check: "+last_full_check_date.get(RowNum)+" days ago");
				cell3.setCellValue("Pass issue: "+true_false1.get(RowNum));
			}
			
			//Amount of clients with tasks, but no in list 
			Row row3 = sheet1.createRow(75);
			Cell cell7 = row3.createCell(10);
			cell7.setCellValue(clients_task.size());
			
			//Save file
			FileOutputStream fileOut = new FileOutputStream("/home/ksmirnovs/retention/retention_"+latestDB+".xls");
			wb.write(fileOut);
			fileOut.close();
			driver.close();

		} catch (SQLException se) {
			se.printStackTrace();
			Assert.fail();
		} catch (Exception e) {

			e.printStackTrace();
		} finally {

			try {
				if (stmt != null)
					conn.close();
			} catch (SQLException se) {
			}
			try {
				if (conn != null)
					conn.close();
			} catch (SQLException se) {
				se.printStackTrace();
			}
		}
	}
	
	@Test
	@Order(order = 2)
	
		public void upsale31() {

		System.out.println("Upsales 31");
		
			try {
				Class.forName("org.postgresql.Driver");
				conn = DriverManager.getConnection(DB_URL, USER, PASS);
				System.out.println("Connected to " + DB_URL + " successfully...");

				stmt = conn.createStatement();

				File file = new File("/home/ksmirnovs/retention/retention_"+latestDB+".xls");
				Workbook wb = WorkbookFactory.create(file);
				
				Sheet sheet3= wb.createSheet("Upsale31");
				Sheet sheet4= wb.createSheet("Wrong31");
				
				ArrayList<String> clients = new ArrayList<String>();
				ArrayList<String> limits = new ArrayList<String>();
				ArrayList<String> visibleLimits = new ArrayList<String>();
				ArrayList<String> accept_marketings = new ArrayList<String>();
				ArrayList<String> no_communications = new ArrayList<String>();
				ArrayList<String> active_loans = new ArrayList<String>();
				ArrayList<String> id = new ArrayList<String>();
				ArrayList<String> ip = new ArrayList<String>();
				ArrayList<String> phone = new ArrayList<String>();
				ArrayList<String> region = new ArrayList<String>();
				ArrayList<String> address = new ArrayList<String>();
				ArrayList<String> passport = new ArrayList<String>();
				ArrayList<String> blackListFull = new ArrayList<>();
				ArrayList<String> r_blacklisted = new ArrayList<String>();
				ArrayList<String> black_list_ip = new ArrayList<String>();
				ArrayList<String> black_list_phone = new ArrayList<String>();
				ArrayList<String> black_list_pass = new ArrayList<String>();
				ArrayList<String> client_bl_ips = new ArrayList<String>();
				ArrayList<String> client_bl_phones = new ArrayList<String>();
				ArrayList<String> client_bl_passports = new ArrayList<String>();
				ArrayList<String> unhappy_clients = new ArrayList<String>();
				ArrayList<String> crm_comm_id = new ArrayList<String>();
				ArrayList<String> true_false1 = new ArrayList<String>();
				ArrayList<String> true_false3 = new ArrayList<String>();
				ArrayList<String> true_false5 = new ArrayList<String>();
				ArrayList<String> true_false6 = new ArrayList<String>();
				ArrayList<String> clients_tasks_upsale5 = new ArrayList<String>();
				ArrayList<String> bl_clients_tasks_upsale5 = new ArrayList<String>();
				ArrayList<String> clients_late = new ArrayList<String>();
				ArrayList<String> clients_task = new ArrayList<String>();
				ArrayList<String> wait_approval = new ArrayList<String>();
				ArrayList<String> page = new ArrayList<String>();
				ArrayList<String> latess = new ArrayList<String>();
				ArrayList<String> full_check = new ArrayList<String>();
				ArrayList<String> passport_clients = new ArrayList<String>();
				ArrayList<String> last_full_check_date = new ArrayList<String>();
				ArrayList<String> pass_issue = new ArrayList<String>();
				ArrayList<String> last_full_check_date_full = new ArrayList<String>();

				//Checking all clients with tasks
				String select = "select r_client_id, created from crm_task where crm_task_type_id = '34' and created::date=current_date::date - interval '1 day';";
				
				//Creating list of clients (Point 1)
				String select1_1 = "SELECT distinct r_client_id FROM r_loan WHERE PAYMENT_DATE::date = (current_date::date - interval '32 days') and status = 'Completed' and r_client_id not in (select r_client_id from r_loan where loan_date >= current_date::date - interval '32 days')";
				String select1_2 = "SELECT c1.r_loan_id, c1.r_client_id, c1.updated as complete_date, 'Canceled' as status,rc.mobile_phone,rc.client_no FROM (SELECT r_client_id, max(created) as created, max(updated) as updated, max(r_loan_id) as r_loan_id FROM r_loan rl WHERE rl.status = 'Canceled' GROUP BY r_client_id) c1 LEFT JOIN r_client rc on rc.r_client_id=c1.r_client_id left JOIN (SELECT r_client_id, max(created) as created FROM (SELECT rank() OVER (PARTITION BY r_client_id ORDER BY created) AS rank, * FROM r_loan WHERE status NOT IN ('Voided', 'Canceled') ORDER BY r_client_id) ranked GROUP BY r_client_id) other on other.r_client_id=c1.r_client_id and c1.created>other.created WHERE other.r_client_id is not NULL AND (-1)*extract(DAY FROM c1.updated-current_date) = 31;";
				String select1_3 = "SELECT distinct r_client_id FROM r_loan WHERE status = 'Completed' and r_client_id IN (select sta.r_client_id from (SELECT r_client_id, max(r_application_id) as rid FROM r_application GROUP BY r_client_id) sta LEFT JOIN r_application ra on sta.rid=ra.r_application_id LEFT JOIN (SELECT r_client_id, max(rank) as rank FROM (SELECT rank() OVER (PARTITION BY r_client_id ORDER BY loan_date) AS rank, * FROM r_loan WHERE status NOT IN ('Voided', 'Canceled') ORDER BY r_client_id) as rakn GROUP BY r_client_id) rank on rank.r_client_id=sta.r_client_id LEFT JOIN r_client rc ON rc.r_client_id = sta.r_client_id WHERE ra.status='EXPIRED' AND  rank.rank>=1 AND (-1)*extract(DAY FROM ra.updated-current_date) = 31);";
				
				ResultSet rs = stmt.executeQuery(select1_1);
				while (rs.next()) {
					String client = rs.getString("r_client_id");
					clients.add(client);
				}

				ResultSet rs2 = stmt.executeQuery(select1_2);
				while (rs2.next()) {
					String client = rs2.getString("r_client_id");
					clients.add(client);
				}

				ResultSet rs3 = stmt.executeQuery(select1_3);
				while (rs3.next()) {
					String client = rs3.getString("r_client_id");
					clients.add(client);
				}

				ResultSet rs444 = stmt.executeQuery(select);
				while (rs444.next()) {
					String task = rs444.getString("r_client_id");
					clients_task.add(task);
				}
				
				//Checking clients with "No Communication"
				for (String c : clients) {
					String nc = "select * from r_client where IS_NO_COMMUNICATION = 't' and r_client_id ='" + c + "'";
					ResultSet rs44 = stmt.executeQuery(nc);
					while (rs44.next()) {
						String no_communication = rs44.getString("r_client_id");
						no_communications.add(no_communication);
					}
				}

				//Checking clients with Active loans
				for (String c : clients) {
					String active = "select * from r_loan where status = 'Active' and r_client_id ='" + c + "'and created::date != current_date";
					ResultSet rs5 = stmt.executeQuery(active);
					while (rs5.next()) {
						String active_loan = rs5.getString("r_client_id");
						active_loans.add(active_loan);
					}
				}
				
				//Checking clients with "Late 30+"
				for (String c : clients) {
					String lates = "SELECT rl.r_client_id, d.total_delay FROM ( SELECT max(r_loan_id) as r_loan_id, r_client_id FROM r_loan WHERE status in ('Completed') GROUP BY r_client_id ) as t LEFT JOIN r_loan rl using (r_loan_id) LEFT JOIN (SELECT (sum(late_days_actual)+sum(late_days_saved))::integer as total_delay, r_loan_id FROM 	r_loan_history GROUP BY r_loan_id) as d USING (r_loan_id) LEFT JOIN	r_client rc ON rc.r_client_id = t.r_client_id WHERE d.total_delay > '30' and rl.r_client_id = '" + c + "'";
					ResultSet lt = stmt.executeQuery(lates);
					while (lt.next()) {
						String late = lt.getString("r_client_id");
						clients_late.add(late);
					}
				}
				
				//Checking clients with full check 85 days +
				for (String c : clients) {
					String full_checked = "select * from r_client where full_check_date < current_date::date - interval '85 day' and r_client_id ='" + c + "'";
					ResultSet lt = stmt.executeQuery(full_checked);
					while (lt.next()) {
						String full_checking = lt.getString("r_client_id");
						full_check.add(full_checking);
					}
				}

				//Checking "Late" clients 
				for (String c : clients) {
					String latee = "select * from r_loan where status = 'Late' and r_client_id ='" + c	+ "'and created::date != current_date";
					ResultSet wt = stmt.executeQuery(latee);
					while (wt.next()) {
						String late_client = wt.getString("r_client_id");
						latess.add(late_client);
					}
				}

				//Checking clients with "Wait ..." application 
				for (String c : clients) {
					String wait = "select * from r_application where status in ('WAIT_WEB_APPROVAL','WAIT_CHANGES_APPROVAL') and r_client_id = '" + c + "'";
					ResultSet wt = stmt.executeQuery(wait);
					while (wt.next()) {
						String wait_app = wt.getString("r_client_id");
						wait_approval.add(wait_app);
					}
				}
				
				//Checking clients from Black Lists
				for (String c : clients) {
					String r_blacklist = "select * from r_blacklist where r_client_id ='" + c + "'";
					ResultSet rs_bl = stmt.executeQuery(r_blacklist);
					while (rs_bl.next()) {
						String blacklisted = rs_bl.getString("r_client_id");
						r_blacklisted.add(blacklisted);
						Collections.addAll(blackListFull, blacklisted);
					}
				}

				for (String c : clients) {
					String r_blacklist_ip = "select IP_ADDRESS from r_application where r_client_id ='" + c + "'";
					ResultSet rs_bl_ip = stmt.executeQuery(r_blacklist_ip);
					while (rs_bl_ip.next()) {
						String client_ip = rs_bl_ip.getString("IP_ADDRESS");
						ip.add(client_ip);
					}
				}
				
				for (String c : clients) {
					String bl = "select * from r_client where r_client_id ='" + c + "'";
					ResultSet rs6 = stmt.executeQuery(bl);
					while (rs6.next()) {
						String client_id = rs6.getString("R_CLIENT_ID");
						String client_phone = rs6.getString("MOBILE_PHONE");
						String client_region = rs6.getString("REG_COUNTRY_ID");
						String client_address = rs6.getString("DECL_ADDRESS_ID");
						String client_passport = rs6.getString("PASSPORT_NO");

						id.add(client_id);
						phone.add(client_phone);
						region.add(client_region);
						address.add(client_address);
						passport.add(client_passport.replaceAll("\\D+",""));

					}
				}

				for (String i : ip) {
					String select_ip = "select * from dms_blacklist_ip where ip ='" + i + "' and active_till is not NULL and active_till >= current_date";
					ResultSet rs7 = stmt.executeQuery(select_ip);
					while (rs7.next()) {
						String bl_ip = rs7.getString("ip");
						black_list_ip.add(bl_ip);
					}
				}

				for (String i2 : black_list_ip) {
					String select_ip2 = "select r_client_id from r_application where IP_ADDRESS ='" + i2 + "' order by created desc limit 1";
					ResultSet rs8 = stmt.executeQuery(select_ip2);
					while (rs8.next()) {
						String client_bl_ip = rs8.getString("r_client_id");
						client_bl_ips.add(client_bl_ip);
						Collections.addAll(blackListFull, client_bl_ip);
					}
				}

				for (String ph : phone) {
					String select_phone = "select * from dms_blacklist_phones where phone ='" + ph + "' and active_till is not NULL and active_till >= current_date";
					ResultSet rs9 = stmt.executeQuery(select_phone);
					while (rs9.next()) {
						String bl_phone = rs9.getString("phone");
						black_list_phone.add(bl_phone);
					}
				}

				for (String ph2 : black_list_phone) {
					String select_phone2 = "select r_client_id from r_client where mobile_phone ='" + ph2 + "'";
					ResultSet rs10 = stmt.executeQuery(select_phone2);
					while (rs10.next()) {
						String client_bl_phone = rs10.getString("r_client_id");
						client_bl_phones.add(client_bl_phone);
						Collections.addAll(blackListFull, client_bl_phone);
					}
				}

				for (String pass : passport) {
					String select_pass = "select * from dms_blacklist_passports where passport_no ='" + pass + "' and active_till is not NULL and active_till >= current_date";
					ResultSet rs11 = stmt.executeQuery(select_pass);
					while (rs11.next()) {
						String bl_pass = rs11.getString("passport_no");
						black_list_pass.add(bl_pass);
					}
				}

				for (String pass2 : black_list_pass) {
					String select_pass2 = "select r_client_id from r_client where passport_no ='" + pass2 + "'";
					ResultSet rs12 = stmt.executeQuery(select_pass2);
					while (rs12.next()) {
						String client_bl_pass = rs12.getString("r_client_id");
						client_bl_passports.add(client_bl_pass);
						Collections.addAll(blackListFull, client_bl_pass);
					}
				}

				for (String up5 : blackListFull) {
					String blacklistTaskUpsale5 = "select * from crm_task where crm_task_type_id = '39' and r_client_id ='"	+ up5 + "' and created::date = current_date";
					ResultSet blc = stmt.executeQuery(blacklistTaskUpsale5);
					while (blc.next()) {
						String task_for_bl_clients = blc.getString("r_client_id");
						bl_clients_tasks_upsale5.add(task_for_bl_clients);
					}
				}
				
				for (String up5 : clients) {
					String tasks = "select * from crm_task where crm_task_type_id = '34' and r_client_id ='"
							+ up5 + "' and created::date = current_date::date - interval '1 day'";
					ResultSet tk = stmt.executeQuery(tasks);
					while (tk.next()) {
						String task_for_clients = tk.getString("r_client_id");
						clients_tasks_upsale5.add(task_for_clients);
					}
				}

				//Checking "Unhappy clients"
				for (String c : clients) {
					String unhappy = "select * from crm_task where crm_task_status_id = '515' and r_client_id ='" + c + "'  and (CRM_TASK_STATUS_ID = '513' or crm_task_status_id = '514' or crm_task_status_id = '515')";
					ResultSet rs13 = stmt.executeQuery(unhappy);
					while (rs13.next()) {
						String unhappy_client = rs13.getString("r_client_id");
						unhappy_clients.add(unhappy_client);
					}
				}
				
				//Bad Passport
				for (String c : clients) {
					String twenty = "SELECT r_client_id, PASSPORT_ISSUED_DATE, DOB from r_client where (EXTRACT(DAY FROM ((r_client.dob + INTERVAL '20 years') - r_client.dob)) <= current_date - r_client.dob AND EXTRACT(DAY FROM ((r_client.dob + INTERVAL '20 years') - r_client.dob)) > passport_issued_date - r_client.dob) and r_client_id = '" + c	+ "'";
					ResultSet rs20 = stmt.executeQuery(twenty);
					while (rs20.next()) {
						String client_passport20 = rs20.getString("r_client_id");
						passport_clients.add(client_passport20);
					}
				}
				
				for (String c : clients) {
					String fourtyfive = "SELECT r_client_id, PASSPORT_ISSUED_DATE, DOB from r_client where (EXTRACT(DAY FROM ((r_client.dob + INTERVAL '45 years') - r_client.dob)) <= current_date - r_client.dob AND EXTRACT(DAY FROM ((r_client.dob + INTERVAL '45 years') - r_client.dob)) > passport_issued_date - r_client.dob) and r_client_id = '" + c	+ "'";
					ResultSet rs45 = stmt.executeQuery(fourtyfive);
					while (rs45.next()) {
						String client_passport45 = rs45.getString("r_client_id");
						passport_clients.add(client_passport45);
					}
				}

				//Remove clients from the list (Point 2)
				for (Object obj : clients) {
					clients_task.remove(obj);
				}
				for (Object obj : no_communications) {
					clients.remove(obj);
				}
				for (Object obj : active_loans) {
					clients.remove(obj);
				}
				for (Object obj : clients_late) {
					clients.remove(obj);
				}
				for (Object obj : unhappy_clients) {
					clients.remove(obj);
				}
				for (Object obj : blackListFull) {
					clients.remove(obj);
				}
				for (Object obj : wait_approval) {
					clients.remove(obj);
				}
				for (Object obj : latess) {
					clients.remove(obj);
				}
				for (Object obj : full_check) {
					clients.remove(obj);
				}
				for (Object obj : passport_clients) {
					clients.remove(obj);
				}
				
				System.out.println("Total clients with task, but not in list "+clients_task);
				
				for (String t : clients_task){
					String last_full_check = "select (current_date-FULL_CHECK_DATE::date) from r_client where r_client_id = '" + t + "'";
					ResultSet rs999 = stmt.executeQuery(last_full_check);
					while (rs999.next()) {
						String lfc = rs999.getString("?COLUMN?");
						last_full_check_date.add(lfc);
					}
				}
				
				for (String c : clients){
					String last_full_check = "select (current_date-FULL_CHECK_DATE::date) from r_client where r_client_id = '" + c + "'";
					ResultSet rs999 = stmt.executeQuery(last_full_check);
					while (rs999.next()) {
						String lfc = rs999.getString("?COLUMN?");
						last_full_check_date_full.add(lfc);
					}
				}

				for (String t : clients_task){
					String pass_issue_date = "SELECT r_client_id, PASSPORT_ISSUED_DATE, DOB from r_client where (EXTRACT(YEAR FROM PASSPORT_ISSUED_DATE) - EXTRACT(YEAR FROM DOB) < '20')  and EXTRACT(YEAR FROM DOB) > '20' and EXTRACT(YEAR FROM current_date) - EXTRACT(YEAR FROM PASSPORT_ISSUED_DATE)> '20' and EXTRACT(YEAR FROM DOB) = EXTRACT(YEAR FROM PASSPORT_ISSUED_DATE) and r_client_id ='"+ t +"'";
					ResultSet rs998 = stmt.executeQuery(pass_issue_date);
					while (rs998.next()) {
						String psis = rs998.getString("PASSPORT_ISSUED_DATE");
						pass_issue.add(psis);
					}
					if (pass_issue.size() > 0) {
						System.out.println(t + " has bad pass");
						true_false1.add("NOK");
					} else {
						System.out.println("OK for " + t);
						true_false1.add("OK");
					}
				}
				
				System.out.println("Total clients After Cut "+clients.size());
				
				//Checking Limit for Clients (Point 3)
				for (String c : clients) {
					driver.get("http://dms.fin.dyninno.net/api/client/" + c + "/nextlimit");
					String limit = driver.findElement(By.xpath("/html/body/pre")).getText();
					limit = limit.substring(0, Math.min(limit.length(), 39));
					limit = limit.substring(34);
					limits.add(limit);
				}

				//Checking that limit is visible on client's page and first page is "Мои Акции"(Point 4, Point 5) 
				driver.get(dyninno);
				driver.manage().window().maximize();
				driver.findElement(By.name("login")).sendKeys("k.smirnovs");
				driver.findElement(By.name("pwd")).sendKeys("zaq1@WSX");
				driver.findElement(By.xpath("//*[@id='ajaxModal']/div/div/div/form/div[2]/span/button")).click();

				for (String c : clients) {
					driver.findElement(By.linkText("Clients")).click();
					driver.findElement(By.id("id")).sendKeys(c);
					driver.findElement(By.id("btn-search")).click();
					sleep(1);
					driver.findElement(By.xpath("//*[@id='tbl']/tbody/tr/td[1]/a")).click();
					driver.findElement(By.xpath("//*[@id='content']/div/div[1]/section[1]/header/ul/li[2]/a")).click();
					driver.findElement(By.id("hijack")).click();
					sleep(1);
					ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());
					driver.switchTo().window(tabs.get(1));
					sleep(1);
					 WebElement element = driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div/div[2]/div[1]/h1"));
					 String test = element.getText();
					 if (test != "Мои акции") {
					 System.out.println("Page Мои Акции doesn't open for "+c);
					 page.add(test);
					 }
					 else{
						 System.out.println(c+ "client is on Мои Акции page");
						 page.add(test); 
					 }
					driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div/div[1]/nav/a[3]")).click();
					sleep(1);
					WebElement element2 = driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div/div[2]/div[2]"));
					String test2 = element2.getText();
					System.out.println(c + ": " + test2);
					test2 = (test2.substring(test2.length() - 49));
					String noLimit = "Вы получите информацию о специальном предложении.";
					if (test2.contentEquals(noLimit)) {
						String fail = "0";
						visibleLimits.add(fail);
						driver.findElement(By.className("cabinet-logout-link")).click();
						driver.close();
						driver.switchTo().window(tabs.get(0));
					} else {
						driver.get("http://dms.fin.dyninno.net/api/client/" + c + "/nextlimit");
						String limit = driver.findElement(By.xpath("/html/body/pre")).getText();
						limit = limit.substring(0, Math.min(limit.length(), 39));
						limit = limit.substring(34);
						visibleLimits.add(limit);
						driver.close();
						driver.switchTo().window(tabs.get(0));
					}
				}

				//Checking clients with Accept Marketing = false
				for (String c : clients) {
					String am = "select * from r_client where ACCEPT_MARKETING = 'f' and r_client_id ='" + c + "'";
					ResultSet rs14 = stmt.executeQuery(am);
					while (rs14.next()) {
						String accept_marketing = rs14.getString("r_client_id");
						accept_marketings.add(accept_marketing);
					}
				}

				//Remove clients with Accept Marketing = false (Point 7)
//				for (Object obj : accept_marketings) {
//					clients.remove(obj);
//				}

				//if marketing accepted
				for (String c : clients) {
					if (accept_marketings.contains(c)) {
						System.out.println(c + " Marketing Accepted = false");
						true_false6.add("N");
					} else {
						System.out.println(c + " Marketing Accepted = true");
						true_false6.add("Y");
					}
				}
				
				// if sent (Point 8)
				for (String c : clients) {
					crm_comm_id.clear();
					String crm = "select created from crm_communication where r_client_id ='" + c + "' and ((MESSAGE_IDENTIFIER = 'RETENTION 3 (SMS) W/O BONUS' or MESSAGE_IDENTIFIER = 'RETENTION - EMAIL 3-2 (W/O bonus)') or (MESSAGE_IDENTIFIER = 'RETENTION - 3 (SMS) W BONUS' or MESSAGE_IDENTIFIER = 'RETENTION - EMAIL 3-1 (W bonus)')) and created::date = current_date::date - interval '1 day';";
					ResultSet rs15 = stmt.executeQuery(crm);
					while (rs15.next()) {
						String comm_id = rs15.getString("created");
						crm_comm_id.add(comm_id);
					}
					if (crm_comm_id.size() != 2) {
						System.out.println("No record for " + c + " created on current date");
						true_false3.add("Not Sent");
					} else {
						System.out.println("SMS and Email were sent for " + c);
						true_false3.add("Sent");
					}
				}
				
				//if task (Point 9)
				for (String c : clients) {
					if (clients_tasks_upsale5.contains(c)) {
						System.out.println(c + " has Task");
						true_false5.add("Task created");
					} else {
						System.out.println(c + " hasn't Task");
						true_false5.add("Task doesn't created");
					}
				}
				
				//Creating xls Report
				for (int RowNum = 0; RowNum < clients.size(); RowNum++) {
					Row row2 = sheet3.createRow(RowNum);
					Cell cell1 = row2.createCell(0);
					Cell cell2 = row2.createCell(1);
					Cell cell3 = row2.createCell(2);
					Cell cell4 = row2.createCell(3);
					Cell cell5 = row2.createCell(4);
					Cell cell6 = row2.createCell(5);
					Cell cell7 = row2.createCell(6);
					cell1.setCellValue("Client Id: "+clients.get(RowNum));
					cell2.setCellValue("Limit: "+limits.get(RowNum));
					cell3.setCellValue(visibleLimits.get(RowNum));
					cell4.setCellValue(true_false3.get(RowNum));
					cell5.setCellValue(true_false5.get(RowNum));
					cell6.setCellValue("Last full check: "+last_full_check_date_full.get(RowNum)+" days ago");
					cell7.setCellValue("Marketing Accepted: "+true_false6.get(RowNum));
				}
				
				for (int RowNum = 0; RowNum < clients_task.size(); RowNum++) {
					Row row1 = sheet4.createRow(RowNum);
					Cell cell1 = row1.createCell(0);
					Cell cell2 = row1.createCell(1);
					Cell cell3 = row1.createCell(2);
					cell1.setCellValue("Client Id: "+clients_task.get(RowNum));
					cell2.setCellValue("Last full check: "+last_full_check_date.get(RowNum)+" days ago");
					cell3.setCellValue("Pass issue: "+true_false1.get(RowNum));
				}
				
				//Amount of clients with tasks, but no in list 
				Row row3 = sheet3.createRow(75);
				Cell cell7 = row3.createCell(10);
				cell7.setCellValue(clients_task.size());
				
				//Save file
				FileOutputStream fileOut = new FileOutputStream(new File("/home/ksmirnovs/retention/retention_"+latestDB+".xls"));
				wb.write(fileOut);
				fileOut.close();
				driver.close();

			} catch (SQLException se) {
				se.printStackTrace();
				Assert.fail();
			} catch (Exception e) {

				e.printStackTrace();
			} finally {

				try {
					if (stmt != null)
						conn.close();
				} catch (SQLException se) {
				}
				try {
					if (conn != null)
						conn.close();
				} catch (SQLException se) {
					se.printStackTrace();
				}
			}
	}
	
	@Test
	@Order(order = 3)

	//Sending report to Oleg and Vadim
	
		public void send() {

			final String username = "retentionReport@gmail.com";
			final String password = "zaq123$%^YHN";

			Calendar calendar = Calendar.getInstance();
			SimpleDateFormat format = new SimpleDateFormat("yyyyMMdd");
			String todaysDate = format.format(calendar.getTime());
			
			Properties props = new Properties();
			
			props.put("mail.smtp.ssl.enable", true);
			props.put("mail.smtp.auth", true);
			props.put("mail.smtp.host", "smtp.gmail.com");
			props.put("mail.smtp.port", "465");

			Session session = Session.getInstance(props, new javax.mail.Authenticator() {
				protected PasswordAuthentication getPasswordAuthentication() {
					return new PasswordAuthentication(username, password);
				}
			});

			try {

				Message message = new MimeMessage(session);
				message.setFrom(new InternetAddress("RetentionReport"));
				message.setRecipients(Message.RecipientType.TO, InternetAddress.parse("kirill.smirnov995@gmail.com, k.smirnovs@eco-fin.eu, o.magid@eco-fin.eu, v.vodeaniuc@eco-fin.eu"));
				message.setSubject("Retention "+todaysDate);
				message.setText("PFA");

				MimeBodyPart messageBodyPart = new MimeBodyPart();

				Multipart multipart = new MimeMultipart();

				messageBodyPart = new MimeBodyPart();
				String file = "/home/ksmirnovs/retention/retention_"+todaysDate+".xls";
				String fileName = "Retention.xls";
				DataSource source = new FileDataSource(file);
				messageBodyPart.setDataHandler(new DataHandler(source));
				messageBodyPart.setFileName(fileName);
				multipart.addBodyPart(messageBodyPart);

				message.setContent(multipart);

				System.out.println("Sending");

				Transport.send(message);
				System.out.println("Done");

			} catch (MessagingException e) {
				e.printStackTrace();
			}
		}
	}