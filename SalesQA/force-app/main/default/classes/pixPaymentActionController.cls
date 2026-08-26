public without sharing class pixPaymentActionController {

    @AuraEnabled
    public static void processPayment(String orderId, Integer cieloPaymentStatus){
        try {
            Payment p = [SELECT Id, GatewayRefNumber FROM Payment WHERE PaymentGroupId IN (SELECT Id FROM PaymentGroup WHERE SourceObjectId = :orderId)][0];
            CommerceCieloIntegration cci = new CommerceCieloIntegration('CieloGetPayment');
            cci.responseData.put('Payment', new Map<String,Object>{'Type'=> 'Pix', 'PaymentId' => p.GatewayRefNumber, 'Status' => cieloPaymentStatus});
            cci.processReturn();
            cci.processFinalDMLs();
        } catch (Exception e) {
            throw new AuraHandledException(e.getMessage());
        }
    }
}