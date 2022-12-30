package com.report;


import java.net.URI;
import java.net.URISyntaxException;

import com.thed.zephyr.cloud.rest.ZFJCloudRestClient;
import com.thed.zephyr.cloud.rest.client.JwtGenerator;


public class JWT_Token_Generator {

	public static String key(String url) throws URISyntaxException {
		
		
			String zephyrBaseUrl = "https://prod-api.zephyr4jiracloud.com/connect";
			String accessKey = "amlyYTpkZWM5MzA4OC02YmRhLTQ0ZDAtOWM0YS1hMjI5M2Q1MTc4OTQgNjI5MDZiZTQxZTRmNGIwMDY4MWRmNjdjIFVTRVJfREVGQVVMVF9OQU1F";
			String secretkey = "V5EJ5IE8nbA3KFouedCz-YTyXpzVhoe5PzwbWNlhmGM";
			String accountId = "62906be41e4f4b00681df67c";

			 ZFJCloudRestClient client = ZFJCloudRestClient.restBuilder(zephyrBaseUrl, accessKey, secretkey, accountId).build();
			 JwtGenerator jwtGenerator = client.getJwtGenerator();
	
			
			 URI uri = new URI(url);
			 int expirationInSec = 3600;
			   String jwt = jwtGenerator.generateJWT("GET", uri, expirationInSec);
			
			 // Print the URL and JWT token to be used for making the REST call
		//	 System.out.println("FINAL API : "+uri.toString());
		//	 System.out.println(jwt); 
			return jwt;
		

	}

	


}
