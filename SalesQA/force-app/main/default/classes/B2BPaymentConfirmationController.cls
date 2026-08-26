public with sharing class B2BPaymentConfirmationController {

    public class PaymentData{
        String type;
        Datetime expiresAt;
        String copyAndPaste;
        String qrCodeUrl;
    }

    @AuraEnabled
    public static string getPaymentData(String orderNumber){
        try {
            PaymentData pd = new PaymentData();
            Order o = [SELECT MetodoPagamento__c, ChavePix__c, URLPix__c, PaymentExpiration__c FROM Order WHERE OrderReferenceNumber = :orderNumber  LIMIT 1];
            pd.type = o.MetodoPagamento__c;
            pd.expiresAt = o.PaymentExpiration__c;
            pd.copyAndPaste = o.ChavePix__c;
            pd.qrCodeUrl = o.URLPix__c;

            return Json.serialize(pd);
        } catch (Exception e) {
            throw new AuraHandledException(e.getMessage());
        }
    }
}