public without sharing class B2BMultipleQuantityController {
    @AuraEnabled
    public static Decimal getMultipleQuantity(String productId) {
        User user = [SELECT Contact.Account.Unidade__c FROM User WHERE Id = :UserInfo.getUserId()];
        UnidadeProduto__c unidade = [SELECT UnidadePedido__r.TipoMedida__c FROM UnidadeProduto__c WHERE Unidade__c = :user.Contact.Account.Unidade__c AND Product__c = :productId LIMIT 1];
        if (unidade.UnidadePedido__r.TipoMedida__c) {
            return 1;
        } else {
            Product2 product = [SELECT PesoUnitario__c FROM Product2 WHERE Id =:productId];
            return product.PesoUnitario__c;
        }
    }
}