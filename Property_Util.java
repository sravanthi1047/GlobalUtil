package GenericHelper_Utilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.Properties;

public class Property_Util {
	private Properties prop = new Properties();
	private File prop_file= null;
	
	public Property_Util(String filePath) {
		try {
			prop_file=new File(filePath);
			prop.load(new FileInputStream(prop_file));
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public Property_Util(FileInputStream inStrFile) {
		System.out.println("Not yet implemented");
	}
	
	public Property_Util(File propFile) {
		prop_file=propFile;
		System.out.println("Not yet implemented");
	}

	public String GetPropertyValue(String key) {
		System.out.println(prop.keySet());
		return prop.getProperty(key);
	}
	
	public boolean SetPropertyValue(String key, String value) {
		try {
			FileOutputStream outF = new FileOutputStream(prop_file);
			prop.setProperty(key, value);
			LocalDateTime nowTime = java.time.LocalDateTime.now();
			prop.store(outF, "Modified at "+ nowTime.toString() );
			return true;
		} catch(Exception e) {
			e.printStackTrace();
			return false;
		}
	}

}
