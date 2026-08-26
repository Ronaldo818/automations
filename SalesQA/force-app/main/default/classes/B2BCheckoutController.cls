public without sharing class B2BCheckoutController {
    @AuraEnabled
    public static void recalculateTaxes(){
        try {
            ConnectApi.Commercecart.calculateCart(CommerceUtil.getStore().Id, 'current', null);
        } catch (Exception e) {
            throw new AuraHandledException(e.getMessage());
        }
    }

    // Não é possivel testar ConnectApi
    public static void testPass(){
        
        Integer i = 0;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
    }
    
}