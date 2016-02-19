package za.co.cmb.sharepoint;

import org.apache.http.*;
import org.apache.http.auth.AuthScope;
import org.apache.http.auth.NTCredentials;
import org.apache.http.client.CredentialsProvider;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpHead;
import org.apache.http.client.methods.HttpPut;
import org.apache.http.client.protocol.HttpClientContext;
import org.apache.http.entity.ByteArrayEntity;
import org.apache.http.impl.client.BasicCredentialsProvider;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.http.impl.client.DefaultHttpRequestRetryHandler;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.protocol.HttpContext;
import org.apache.http.util.EntityUtils;
import org.apache.log4j.Logger;
import org.codehaus.jackson.JsonNode;
import org.codehaus.jackson.map.ObjectMapper;

import za.co.cmb.sharepoint.dto.SharepointSearchResult;
import za.co.cmb.sharepoint.dto.SharepointUser;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.util.*;

public class SharepointManagerSupport implements SharepointManager {

    private Logger LOG = Logger.getLogger(getClass());

    private DefaultHttpClient httpClient;
    private String serverUrl;
    private int port;
    private String domain;
    private String scheme;
    private String urlPrefix;

    public SharepointManagerSupport(String serverUrl, int port, String domain) {
        this.serverUrl = serverUrl;
        this.port = port;
        this.domain = domain;
        scheme = port == 443 ? "https" : "http";
        urlPrefix = port == 443 ? "https://" : "http://";
    }

    private void initHttpConnection() {
        httpClient = new DefaultHttpClient();
        httpClient.addRequestInterceptor(new HttpRequestInterceptor() {
            public void process(HttpRequest request, HttpContext context) throws HttpException, IOException {
                request.setHeader("Accept", "application/json;odata=verbose");
            }
        });
    }

    @Override
    public List<SharepointUser> findAllUsers(String username, String password) throws IOException {
        List<SharepointUser> sharepointUsers = new ArrayList<SharepointUser>();
        String endpoint = urlPrefix + serverUrl + URL_LISTDATA + LIST_USERS;
        JsonNode results = getJsonNode(username, password, endpoint, "results");
        Iterator<JsonNode> iterator = results.getElements();
        while (iterator.hasNext()) {
            JsonNode user = iterator.next();
            SharepointUser sharpointUser = new SharepointUser();
            sharepointUsers.add(sharpointUser);
            sharpointUser.setId(user.get("Id").asText());
            sharpointUser.setName(user.get("Name").getTextValue());
            String pictureUrl = user.get("Picture").getTextValue();
            if (pictureUrl != null) {
                // WTF? why does sharepoint send the picture URL twice ... only Microsoft
                sharpointUser.setPictureUrl(pictureUrl.substring(0, pictureUrl.indexOf(", ")));
            }
        }
        return sharepointUsers;
    }
    
    @Override
	public void UploadFile(String username, String password, String folderRelativeURL, String fileName, byte[] fileData) throws Exception {
    	
    	// reference: http://www.rgagnon.com/javadetails/java-put-document-sharepoint-library.html
    	
    	CloseableHttpClient httpclient = HttpClients.custom().setRetryHandler(new DefaultHttpRequestRetryHandler(0,false)).build();
    	
	     CredentialsProvider credsProvider = new BasicCredentialsProvider();
	     credsProvider.setCredentials(AuthScope.ANY, new NTCredentials(username, password, getComputerName(), this.domain));
	
	     // You may get 401 if you go through a load-balancer.
	     // To fix this, go directly to one the sharepoint web server or
	     // change the config. See this article :
	     // http://blog.crsw.com/2008/10/14/unauthorized-401-1-exception-calling-web-services-in-sharepoint/
	     HttpHost target = new HttpHost(serverUrl, 80, scheme);
	     HttpClientContext context = HttpClientContext.create();
	     context.setCredentialsProvider(credsProvider);
	
	     // The authentication is NTLM.
	     // To trigger it, we send a minimal http request
	     HttpHead request1 = new HttpHead("/");
	     CloseableHttpResponse response1 = null;
	     try {
	       response1 = httpclient.execute(target, request1, context);
	       EntityUtils.consume(response1.getEntity());
	       System.out.println("1 : " + response1.getStatusLine().getStatusCode());
	     }
	     finally {
	       if (response1 != null ) response1.close();
	     }
	
	     // The real request, reuse authentication
	     String file = String.format("/%1s/%2s", folderRelativeURL, fileName);
	     HttpPut request2 = new HttpPut(file);  // target
	     request2.setEntity(new ByteArrayEntity(fileData));
	     //request2.setEntity(new FileEntity(new File("c:/temp/jira.log")));// source
	     CloseableHttpResponse response2 = null;
	     
	     try {
	       response2 = httpclient.execute(target, request2, context);
	       EntityUtils.consume(response2.getEntity());
	       int rc = response2.getStatusLine().getStatusCode();
	       String reason = response2.getStatusLine().getReasonPhrase();
	       // The possible outcomes :
	       //    201 Created
	       //        The request has been fulfilled and resulted in a new resource being created
	       //    200 OK
	       //        Standard response for successful HTTP requests.
	       //    others
	       //        we have a problem
	       if (rc == HttpStatus.SC_CREATED) {
	    	   LOG.info(file + " is copied (new file created)");
	       }
	       else if (rc == HttpStatus.SC_OK) {
	    	   LOG.info(file + " is copied (original overwritten)");
	       }
	       else {
	         throw new Exception("Problem while copying " + file + "  reason " + reason + "  httpcode : " + rc);
	       }
	     }
	     finally {
	       if (response2 != null) response2.close();
	     }
	         
    	
    	/* Java Sample
		    CredentialsProvider credsProvider = new BasicCredentialsProvider();
		    credsProvider.setCredentials(
		            new AuthScope(AuthScope.ANY),
		            new NTCredentials("username", "password", "https://hostname", "domain"));
		    CloseableHttpClient httpclient = HttpClients.custom()
		            .setDefaultCredentialsProvider(credsProvider)
		            .build();
		    try {
		        HttpGet httpget = new HttpGet("http://hostname/_api/web/lists");
		
		        System.out.println("Executing request " + httpget.getRequestLine());
		        CloseableHttpResponse response = httpclient.execute(httpget);
		        try {
		            System.out.println("----------------------------------------");
		            System.out.println(response.getStatusLine());
		            EntityUtils.consume(response.getEntity());
		        } finally {
		            response.close();
		        }
		    } finally {
		        httpclient.close();
		    }

    	 */
    	
    	/* C# sample
    	 * http://sharepoint.stackexchange.com/questions/111674/upload-attachment-to-sharepoint-list-item-using-rest-api-from-java-client
			var fileName = Path.GetFileName(uploadFilePath);
			var requestUrl = string.Format("{0}/_api/web/Lists/GetByTitle('{1}')/items({2})/AttachmentFiles/add(FileName='{3}')", webUrl, listTitle, itemId,fileName);
			var request = (HttpWebRequest)WebRequest.Create(requestUrl);
			request.Credentials = credentials;  //SharePointOnlineCredentials object 
			request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
			
			request.Method = "POST";
			request.Headers.Add("X-RequestDigest", requestDigest);
			
			var fileContent = System.IO.File.ReadAllBytes(uploadFilePath);
			request.ContentLength = fileContent.Length;
			using (var requestStream = request.GetRequestStream())
			{
			    requestStream.Write(fileContent, 0, fileContent.Length);
			}
			
			var response = (HttpWebResponse)request.GetResponse();
    	 */
    	
    	/* javascript code
    	 * https://msdn.microsoft.com/en-us/library/office/dn769086.aspx
    	 * https://msdn.microsoft.com/en-us/library/office/dn268594.aspx
        // Get the file name from the file input control on the page.
        var parts = fileInput[0].value.split('\\');
        var fileName = parts[parts.length - 1];

        // Construct the endpoint.
        var fileCollectionEndpoint = String.format(
            "{0}/_api/sp.appcontextsite(@target)/web/getfolderbyserverrelativeurl('{1}')/files" +
            "/add(overwrite=true, url='{2}')?@target='{3}'",
            appWebUrl, serverRelativeUrlToFolder, fileName, hostWebUrl);

        // Send the request and return the response.
        // This call returns the SharePoint file.
        return jQuery.ajax({
            url: fileCollectionEndpoint,
            type: "POST",
            data: arrayBuffer,
            processData: false,
            headers: {
                "accept": "application/json;odata=verbose",
                "X-RequestDigest": jQuery("#__REQUESTDIGEST").val(),
                "content-length": arrayBuffer.byteLength
            }
        });
        */
    }

    @Override
    public List<SharepointSearchResult> search(String username, String password, String searchWord)
            throws IOException {
        List<SharepointSearchResult> results = new ArrayList<SharepointSearchResult>();
        String endpoint = urlPrefix + serverUrl + URL_SEARCH + "'" + searchWord + "'";
        JsonNode rows = getJsonNode(username, password, endpoint, "Rows");

        Map<String, String> values = new HashMap<String, String>();
        if (rows != null) {
            for (JsonNode cell : rows.findValues("Cells")) {
                for (final JsonNode cellResult : cell.get("results")) {
                    values.put(cellResult.get("Key").getTextValue(), cellResult.get("Value").getTextValue());
                }
                SharepointSearchResult sharepointSearchResult = new SharepointSearchResult();
                sharepointSearchResult.setPath(values.get("Path"));
                sharepointSearchResult.setParentFolder(values.get("ParentLink"));
                sharepointSearchResult.setAuthor(values.get("Author"));
                sharepointSearchResult.setHitHighlightedSummary(values.get("HitHighlightedSummary"));
                sharepointSearchResult.setLastModified(values.get("LastModifiedTime"));
                sharepointSearchResult.setRank(Double.parseDouble(values.get("Rank")));
                sharepointSearchResult.setSiteName(values.get("SiteName"));
                sharepointSearchResult.setTitle(values.get("Title"));
                results.add(sharepointSearchResult);
            }
        }
        return results;
    }

    @Override
    public boolean test(String username, String password) {
        String endpoint = urlPrefix + serverUrl + URL_TEST;
        int responseCode;
        try {
            initHttpConnection();
            addCredentials(username, password);
            HttpGet httpget = new HttpGet(endpoint);

            LOG.debug("Executing request: " + httpget.getRequestLine());
            HttpResponse response = httpClient.execute(httpget);
            responseCode = response.getStatusLine().getStatusCode();
        } catch (Exception e) {
            return false;
            //code
        } finally {
            httpClient.getConnectionManager().shutdown();
        }
        return responseCode == 200;
    }

    private void addCredentials(String username, String password) {
        httpClient.getCredentialsProvider().setCredentials(
                new AuthScope(serverUrl, port),
                new NTCredentials(username, password, "WORKSTATION", domain));
    }

    private JsonNode getJsonNode(String username, String  password, String endpointURL) throws IOException {
        return getJsonNode(username, password, endpointURL, null);
    }

    private JsonNode getJsonNode(String username, String  password, String endpointURL, String elementName)
            throws IOException {
        BufferedReader reader = null;
        JsonNode node = null;
        try {
            initHttpConnection();
            addCredentials(username, password);
            HttpGet httpget = new HttpGet(endpointURL);

            LOG.debug("Executing request: " + httpget.getRequestLine());
            HttpResponse response = httpClient.execute(httpget);
            HttpEntity entity = response.getEntity();

            LOG.debug(response.getStatusLine());
            if (entity != null) {
                LOG.debug("Response content length: " + entity.getContentLength());
            } else {
                LOG.error("No response received");
                return null;
            }

            StringBuilder stringBuilder = new StringBuilder();
            reader = new BufferedReader(new InputStreamReader(entity.getContent()));
            String line;
            while ((line = reader.readLine()) != null) {
                stringBuilder.append(line);
            }
//            System.out.println(stringBuilder.toString());
            EntityUtils.consume(entity);

            ObjectMapper mapper = new ObjectMapper();
            JsonNode rootNode = mapper.readValue(stringBuilder.toString(), JsonNode.class);
            if (elementName == null) {
                node = rootNode;
            } else {
                node =  rootNode.findValue(elementName);
            }
        } catch (Exception e) {
            //code
        } finally {
            if (reader != null) {
                reader.close();
            }
            httpClient.getConnectionManager().shutdown();
        }
        return node;
    }

    private String getComputerName()
    {
        Map<String, String> env = System.getenv();
        if (env.containsKey("COMPUTERNAME"))
            return env.get("COMPUTERNAME");
        else if (env.containsKey("HOSTNAME"))
            return env.get("HOSTNAME");
        else
        {
        	try
        	{
        	    InetAddress addr;
        	    addr = InetAddress.getLocalHost();
        	    return addr.getHostName();
        	}
        	catch (UnknownHostException ex)
        	{
        	    return "Unknown Host";
        	}
        }
    }
    
}
