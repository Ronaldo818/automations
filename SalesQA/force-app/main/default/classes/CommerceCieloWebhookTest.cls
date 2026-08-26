@isTest
public with sharing class CommerceCieloWebhookTest {
    @TestSetup
    static void makeData(){
        TestUtil util = new TestUtil();
        util.genStoreData();
    }

    @IsTest
    static void test(){

        Test.setMock(HttpCalloutMock.class, new CieloIntegrationMock());

        String JSONMsg = '{"PaymentId":"Test","ChangeType":1}';
                
        RestRequest req = new RestRequest();
        RestResponse res = new RestResponse();
        
        req.requestURI = '/services/apexrest/cielo/webhook';//Request URL
        req.httpMethod = 'POST';//HTTP Request Type
        req.requestBody = Blob.valueOf(JSONMsg);
        
        RestContext.request = req;
        RestContext.response= res;
        
        Test.startTest();
            CommerceCieloWebhook.handleWebhook();
        Test.stopTest();
        
    }
}