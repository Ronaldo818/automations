@RestResource(urlMapping='/cielo/webhook')
global with sharing class CommerceCieloWebhook {

    @HttpPost
    global static void handleWebhook() {
        RestRequest req = RestContext.request;
        RestResponse res = RestContext.response;
        IntegrationLog__c ilog = new IntegrationLog__c();
        try {

            String body = req.requestBody.toString();
            ilog.ApexClass__c = 'CommerceCieloWebhook';
            iLog.ApexMethod__c = 'handleWebhook';
            iLog.EndPoint__c = req.resourcePath;
            iLog.IntegrationMethod__c = req.httpMethod;
            iLog.NomeLog__c = 'Log de Integração Cielo - Webhook';
            iLog.Request__c = body.left(32768);

            CieloWebhookPayload payload = 
                (CieloWebhookPayload) JSON.deserialize(body, CieloWebhookPayload.class);

            res.statusCode = 200;
            ilog.IntegrationStatusCode__c = '200';

            if(payload.ChangeType == 1){
                CommerceCieloIntegration cci = new CommerceCieloIntegration('CieloGetPayment', payload.PaymentId);
                cci.process();
                iLog.Response__c = '{"message": "'+ cci.message +'"}';
            }
            else{
                iLog.Response__c = '{"message": "Nenhuma operacao realizada"}';
            }
            insert iLog;

        } catch (Exception e) {
            ilog.IntegrationStatusCode__c = '500';
            iLog.Response__c = e.getMessage() + ' | ' + e.getStackTraceString();
            res.statusCode = 500;
            insert iLog;
        }
        res.responseBody = Blob.valueOf(ilog.Response__c);
    }

    // Classe do payload
    global class CieloWebhookPayload {
        public String RecurrentPaymentId;
        public String PaymentId;
        public Integer ChangeType;
    }
}