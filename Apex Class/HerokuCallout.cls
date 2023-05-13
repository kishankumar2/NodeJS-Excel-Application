/**
 * @description       : 
 * @author            : Kishan Kumar
 * @group             : 
 * @last modified on  : 05-13-2023
 * @last modified by  : Kishan Kumar
**/
public class HerokuCallout {
    
    public static void makeCalloutNodeJS(){
        Http http = new Http();
        HttpRequest request = new HttpRequest();
        request.setEndpoint('callout:Heroku_NODEJS');
        request.setMethod('GET');
        HttpResponse response = http.send(request);
        // If the request is successful, parse the JSON response.
        if(response.getStatusCode() == 200) {
            // Deserialize the JSON string into collections of primitive data types.
            
            System.debug('Received the following :'+response.getBody());
            createExcel(response.getBody());
            
        }
    }
    
    public static void makeCalloutJava(){
        
        System.debug('Invoking Heroku');
        ContentVersion cv = [SELECT Id, ContentDocumentId, VersionData FROM ContentVersion where ContentDocumentId = '0692x00000HPgp9AAD'];
        Blob data=cv.VersionData;
        String st=EncodingUtil.Base64Encode(data);
        input inp = new input(st,'Heroku Sign','05-13-2023');
        
        Http http = new Http();
        HttpRequest request = new HttpRequest();
        request.setEndpoint('callout:Heroku_Java');
        request.setMethod('POST');
        request.setHeader('Content-Type', 'application/json');
        request.setBody(Json.serialize(inp));
        HttpResponse response = http.send(request);
        System.debug('response'+response);
        // If the request is successful, parse the JSON response.
        if(response.getStatusCode() == 200) {
            // Deserialize the JSON string into collections of primitive data types.
            
            System.debug('Received the following :'+response.getBody());
            createPDF(response.getBody());
            
        }
    }
    
    public static void createPDF(String data) {
        
        ContentVersion conver = new ContentVersion();
        conver.Title = 'HerokuPDF';
        conver.PathOnClient = 'HerokuPDF' + '.PDF';
        conver.VersionData = EncodingUtil.base64Decode(data);//blob.valueOf(response.content);
        insert conver;
        
        Id contentDocumentId = [SELECT ContentDocumentId FROM ContentVersion WHERE Id = :conver.Id].ContentDocumentId;
        
        // Link the content document to sObject record in this case it is account 
        ContentDocumentLink condocLink = new ContentDocumentLink();
        condocLink.ContentDocumentId = contentDocumentId;
        //Link the File to CRM Record  
        condocLink.LinkedEntityId = '0012x00000hIpGrAAK';
        condocLink.ShareType = 'V';
        insert condocLink;
    }
    
    
    
    public static void createExcel(String data) {
        
        ContentVersion conver = new ContentVersion();
        conver.Title = 'HerokuExcel';
        conver.PathOnClient = 'HerokuExcel' + '.xlsx';
        conver.VersionData = EncodingUtil.base64Decode(data);//blob.valueOf(response.content);
        insert conver;
        
        Id contentDocumentId = [SELECT ContentDocumentId FROM ContentVersion WHERE Id = :conver.Id].ContentDocumentId;
        
        // Link the content document to sObject record in this case it is account 
        ContentDocumentLink condocLink = new ContentDocumentLink();
        condocLink.ContentDocumentId = contentDocumentId;
        //Link the File to CRM Record  
        condocLink.LinkedEntityId = '0012x00000hIpGrAAK';
        condocLink.ShareType = 'V';
        insert condocLink;
    }
    
    
    
    Public class input{
        Public String pdfData;
        Public String sign;
        Public String signDate;
        
        public input(String pdfData,String sign,String signDate){
            this.pdfData=pdfData;
            this.sign=sign;
            this.signDate=signDate;
        }
    }
    
}