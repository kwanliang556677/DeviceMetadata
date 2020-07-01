package com.example.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.stream.Collectors;

import org.apache.commons.exec.CommandLine;
import org.apache.commons.exec.DefaultExecutor;
import org.apache.commons.exec.ExecuteException;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;
import org.springframework.stereotype.Controller;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ResponseBody;

@Controller
public class GenerateMetadata {
	private static Integer CONTROLLER_THRESHOLD = 1;

	@ResponseBody
	@GetMapping("/")
	public void generateMetadata() {
		List<Device> deviceList = getDeviceList();
		updateDeviceId(deviceList);
		generatePrivatePublicKey(deviceList);
		generateJSON(deviceList);
		generateXLSX(deviceList);
	}
	
	private static void generateXLSX(List<Device> deviceList) {
		Comparator<Device> comparePredicate = Comparator.comparing(Device::getControllerName).thenComparing(Device::getDeviceId);
		List<Device> sortedDeviceList = deviceList.stream().sorted(comparePredicate).collect(Collectors.toList());
		SXSSFWorkbook wb = new SXSSFWorkbook(100);
        Sheet sh = wb.createSheet();
        for (int i = 0; i < sortedDeviceList.size(); i++) {
        	Row row = sh.createRow(i);
        	Device device = sortedDeviceList.get(i);
  			 row.createCell(0).setCellValue(device.getControllerName());
  			 row.createCell(1).setCellValue(device.getDeviceId());
  			 row.createCell(2).setCellValue(device.getPointName());
		}
		try {
			File file = new File("C:\\Users\\PTPL\\Documents\\GoogleProject\\", "gcpDevices.xlsx");
			FileOutputStream out = new FileOutputStream(file);
		    wb.write(out);
		    out.close();
		    wb.close();
		}catch(Exception e) {
		}
	}
	
	private static void generateJSON(List<Device> deviceList) {
		try {
			Map<String, List<Device>> deviceMap = deviceList.stream().collect(Collectors.groupingBy(Device::getDeviceId));
			deviceMap.entrySet().forEach(x->{
				String deviceId = x.getKey();
				List<Device> devices =  x.getValue();
				String guid = devices.get(0).getGuid();
				JSONObject json = new JSONObject();
				JSONObject json2 = new JSONObject();
				JSONObject pointsetJSON = new JSONObject();
				json.put("pointset", pointsetJSON);
				pointsetJSON.put("points", json2);
				json.put("version", 1);
				json.put("timestamp", Instant.now());
				json.put("hash", "58d9d336");
				deviceList.forEach(y->{
					JSONObject json3 = new JSONObject();
					json2.put(y.getPointName(), json3);
					json3.put("units", y.getEngUnit());
				}); 
				
				JSONObject json4 = new JSONObject();
				JSONObject json5 = new JSONObject();
				JSONObject json6 = new JSONObject();
				JSONObject json7 = new JSONObject();
				JSONObject json8 = new JSONObject();
				JSONObject json9 = new JSONObject();
				json.put("system", json4);
				json.put("cloud", json8);
				json8.put("auth_type", "RS256");
				json4.put("location", json5);
				json5.put("site_name", "SG-MBC2-B80");
				json5.put("section", "MBC2-B80");
				json5.put("position", json9);
				json9.put("x", 0);
				json9.put("y", 0);
				json4.put("physical_tag", json6);
				json6.put("asset", json7);
				json7.put("guid", guid);
				json7.put("name", "SG-MBC2-B80"+'_'+deviceId);
				try {
					FileWriter file = new FileWriter("C:\\Users\\PTPL\\Documents\\GoogleProject\\Device_Repo_Directory\\"+deviceId+"\\metadata.json");
					file.write(json.toString());
					file.close();
				}catch(Exception e) {
				}
			});
		}catch(Exception e) {
		}
	}

	private static void generatePrivatePublicKey(List<Device> deviceList) {
		Map<String, List<Device>> deviceMap = deviceList.stream().collect(Collectors.groupingBy(Device::getDeviceId));
		deviceMap.entrySet().forEach(x->{
			String deviceDirectory = "C:\\Users\\PTPL\\Documents\\GoogleProject\\Device_Repo_Directory\\"+x.getKey();
			File file = new File(deviceDirectory);
			if (!file.exists()) {
				file.mkdir();
			}
			try {
				CommandLine cmdLine = CommandLine.parse("cmd /c start /B C:\\Users\\PTPL\\Desktop\\GoogleProj\\OpenSSL-Win64\\bin\\keys.bat");
				DefaultExecutor executor = new DefaultExecutor();
				int exitValue = executor.execute(cmdLine);
			} catch (IOException e1) {
			}			       
			
			File from = new File("C:\\Users\\PTPL\\Downloads\\demo\\rsa_cert.pem");
			File from2 = new File("C:\\Users\\PTPL\\Downloads\\demo\\rsa_private.pem");
			File from3 = new File("C:\\Users\\PTPL\\Downloads\\demo\\rsa_private.pkcs8");
			File to = new File(deviceDirectory);
			try {
				FileUtils.copyFileToDirectory(from, to);
				FileUtils.copyFileToDirectory(from2, to);
				FileUtils.copyFileToDirectory(from3, to);
			} catch (IOException e) {
				e.printStackTrace();
			}
		});
	}

	private static void updateDeviceId(List<Device> deviceList) {
		Map<String, Map<String, List<Device>>> deviceMap = deviceList.stream()
				.collect(Collectors.groupingBy(Device::getDeviceId, Collectors.groupingBy(Device::getControllerName)));
		for (Entry<String, Map<String, List<Device>>> map : deviceMap.entrySet()) {
			Map<String, List<Device>> controllerMap = map.getValue();
			int size = controllerMap.entrySet().size();
			int i = 1;
			if (size > CONTROLLER_THRESHOLD) {
				for (Entry<String, List<Device>> map2 : controllerMap.entrySet()) {
					List<Device> controllerDeviceList = map2.getValue();
					System.out.println(map2.getKey());
					for (Device device : controllerDeviceList) {
						device.setDeviceId(device.getDeviceId() + "-" + i);
						System.out.println(device.getPointName());
					}
					i++;
				}
			}
		}
	}

	private static List<Device> getDeviceList() {
		List<Device> deviceList = new ArrayList<>();
		try {
			File file = ResourceUtils.getFile("classpath:metadata.xlsx");
			FileInputStream fis = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sh = wb.getSheet("Sum L11");
			DataFormatter formatter = new DataFormatter();
			for (int i = 1; i < sh.getLastRowNum() + 1; i++) {
				Device device = new Device();
				device.setDeviceId(formatter.formatCellValue(sh.getRow(i).getCell(0)).replace('_', '-'));
				device.setPointName(formatter.formatCellValue(sh.getRow(i).getCell(1)));
				device.setControllerName(formatter.formatCellValue(sh.getRow(i).getCell(2)));
				device.setGuid(sh.getRow(i).getCell(3).getStringCellValue());
				device.setEngUnit(formatter.formatCellValue(sh.getRow(i).getCell(4)));
				deviceList.add(device);
			}
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		return deviceList;
	}
}
