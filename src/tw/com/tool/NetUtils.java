package tw.com.tool;

import java.io.BufferedReader;
import java.io.DataOutputStream;
import java.io.InputStreamReader;
import java.io.UnsupportedEncodingException;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLEncoder;
import java.security.cert.CertificateException;
import java.security.cert.X509Certificate;
import java.util.Map;

import javax.net.ssl.HttpsURLConnection;
import javax.net.ssl.SSLContext;
import javax.net.ssl.TrustManager;
import javax.net.ssl.X509TrustManager;

import org.slf4j.Logger;

public class NetUtils {

	protected static void doTrustToCertificates(Logger log) {
		
		TrustManager[] trustAllCerts = new TrustManager[]{new TrustAnyTrustManager()};
	
	    // Install the all-trusting trust manager
	    try 
	    {
	        SSLContext sc = SSLContext.getInstance("SSL");
	        sc.init(null, trustAllCerts, new java.security.SecureRandom());
	        HttpsURLConnection.setDefaultSSLSocketFactory(sc.getSocketFactory());
	    } 
	    catch (Exception e) 
	    {
	    	log.debug(e.getMessage());
	    }
	}
	
	protected static class TrustAnyTrustManager implements X509TrustManager {  
		  
	    public void checkClientTrusted(X509Certificate[] chain, String authType)  
	            throws CertificateException {  
	    }  
	
	    public void checkServerTrusted(X509Certificate[] chain, String authType)  
	            throws CertificateException {  
	    }  
	
	    public X509Certificate[] getAcceptedIssuers() {  
	        return new X509Certificate[] {};  
	    }  
	}
	
	protected static String connection(Logger log, String url, String encoder, String method, Map<String,String> headers, String requestString){
		
		URL goUrl = null;
		HttpURLConnection connection = null;
		DataOutputStream writer = null;
		BufferedReader reader = null;
		int responseCode = 0;
		
		try{
			
			StringBuffer response = new StringBuffer();
			goUrl = new URL(url);
			connection = (HttpURLConnection)goUrl.openConnection();
			
			if(headers != null && !headers.isEmpty()){
				
				for(String key : headers.keySet())
					connection.setRequestProperty(key, headers.get(key));
			}
			
			connection.setRequestMethod(method.toUpperCase());
			connection.setDoInput(true);
			connection.setDoOutput(true);
			connection.setConnectTimeout(120000);
			connection.setReadTimeout(120000);
			connection.setUseCaches(false);
			connection.setDefaultUseCaches(false);
			
			if(!requestString.isEmpty()){
				
				writer = new DataOutputStream(connection.getOutputStream());
			    writer.write(requestString.getBytes(encoder.toUpperCase()));
			    writer.flush();
			    writer.close();
			}
			
			responseCode = connection.getResponseCode();
			if (responseCode == HttpURLConnection.HTTP_OK) {
				
				reader = new BufferedReader(new InputStreamReader(connection.getInputStream(),encoder.toUpperCase()));
				String lines;

				while ((lines = reader.readLine()) != null) {

					response.append(lines);
	            };
			}			
            
            if(response.length() > 0){

            	return response.toString();
            }
            	
		}catch(MalformedURLException mue){
			
			log.debug(mue.getMessage());
		}catch(Exception e){
			
			log.debug(e.getMessage());
		}finally{
			
			goUrl = null;
			connection = null;
			reader = null;
		}
		
		return null;
	}
	
	public static String getData(Logger log, Boolean ssl, String url, String encoder, Map<String,String> headers, Map<String,String> parameters){
		
		if(ssl)
			doTrustToCertificates(log);
		
		if(parameters != null && !parameters.isEmpty()){
			
			StringBuffer paramStr = new StringBuffer();
			
			try {
				
				for(String parameter : parameters.keySet())
					paramStr.append("&" + parameter + "=" + URLEncoder.encode(parameters.get(parameter), encoder.toUpperCase()));
			} catch (UnsupportedEncodingException e) {
					
				log.debug(e.getMessage());
			}
			
			if(!url.contains("?"))
				url += ("?" + paramStr.substring(1));
			else
				url += paramStr;
		}
		
		return connection(log, url, encoder, "GET", headers, "");
	}
	
	public static String postXml(Logger log, Boolean ssl, String url, String encoder, Map<String,String> headers, String requestXml){
		
		if(ssl)
			doTrustToCertificates(log);
		
		headers.put("Content-Type", "text/xml;charset=utf-8");
		return connection(log, url, encoder, "POST", headers, requestXml);
	}

	public static String postJson(Logger log, Boolean ssl, String url, String encoder, Map<String,String> headers, String requestJson){
		
		if(ssl)
			doTrustToCertificates(log);
		
		headers.put("Content-Type", "application/json; charset=utf-8");
		return connection(log, url, encoder, "POST", headers, requestJson);
	}
	
	public static String postHttpContent(Logger log, Boolean ssl, String url, String encoder, Map<String,String> headers, Map<String,String> parameters){
		
		if(ssl)
			doTrustToCertificates(log);
		
		StringBuffer requestStr = new StringBuffer();
		if(parameters != null && !parameters.isEmpty()){
			
			try {
				
				for(String parameter : parameters.keySet())
					requestStr.append("&" + parameter + "=" + URLEncoder.encode(parameters.get(parameter), encoder.toUpperCase()));
			} catch (UnsupportedEncodingException e) {
					
				log.debug(e.getMessage());
			}
		}
		
		headers.put("Content-Type", "application/x-www-form-urlencoded");
		return connection(log, url, encoder, "POST", headers, requestStr.substring(1));
	}
}
