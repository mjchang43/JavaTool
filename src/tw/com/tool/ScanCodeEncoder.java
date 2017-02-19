package tw.com.tool;

import java.io.File;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.Hashtable;

import com.google.zxing.BarcodeFormat;
import com.google.zxing.EncodeHintType;
import com.google.zxing.MultiFormatWriter;
import com.google.zxing.WriterException;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.qrcode.decoder.ErrorCorrectionLevel;

public class ScanCodeEncoder {

public static boolean generateBarcodePic(String keyword,String picPath){
		
		try {
			
			MultiFormatWriter writer = new MultiFormatWriter();   
			Hashtable<EncodeHintType,Object> hints = new Hashtable<EncodeHintType,Object>();
			hints.put( EncodeHintType.CHARACTER_SET, "UTF-8" );
			hints.put( EncodeHintType.MARGIN, 1 );
			hints.put( EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.Q );
			
			File file = new File(picPath);
			if(file.mkdirs()){
				
				MatrixToImageWriter.writeToPath( writer.encode( keyword, BarcodeFormat.CODE_128, 396, 75, hints ), "JPG", Paths.get(file.getPath()) );
				return true;
			}
		  } catch (IOException e) {
			  
			  System.out.println("Generate Barcode Picture failed - " + e.getMessage());
		  } catch (WriterException e) {
			  
			  System.out.println("Generate Barcode Picture failed - " + e.getMessage());
		  } catch (Exception e) {
			  
			  System.out.println("Generate Barcode Picture failed - " + e.getMessage());
		  }
		return false;
	}
	
	public static boolean generateQrcodePic(String keyword,String picPath){
		
		try {

			MultiFormatWriter writer = new MultiFormatWriter();   
			Hashtable<EncodeHintType,Object> hints = new Hashtable<EncodeHintType,Object>();
			hints.put( EncodeHintType.CHARACTER_SET, "UTF-8" );
			hints.put( EncodeHintType.MARGIN, 1 );
			hints.put( EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.Q );
			
			File file = new File(picPath);
			if(file.mkdirs()){
				
				MatrixToImageWriter.writeToPath( writer.encode( keyword, BarcodeFormat.QR_CODE, 140, 140, hints ), "JPG",  Paths.get(file.getPath())  );
				return true;
			}
		  } catch (IOException e) {
			  
			  System.out.println("Generate Qrcode Picture failed - " + e.getMessage());
		  } catch (WriterException e) {
			  
			  System.out.println("Generate Qrcode Picture failed - " + e.getMessage());
		  } catch (Exception e) {
			  
			  System.out.println("Generate Qrcode Picture failed - " + e.getMessage());
		  }
		return false;
	}
}
