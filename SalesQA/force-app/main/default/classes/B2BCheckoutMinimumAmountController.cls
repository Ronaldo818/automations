public without sharing class B2BCheckoutMinimumAmountController {

    // @AuraEnabled(cacheable=true)
    // public static Decimal getMinimumAmount(){
    //     try {
    //         return Decimal.valueOf([SELECT Value__c FROM StandardValue__mdt WHERE DeveloperName = 'B2BMinimumOrderAmount' LIMIT 1][0].Value__c);
    //     } catch (Exception e) {
    //         throw new AuraHandledException(e.getMessage());
    //     }
    // }

    @AuraEnabled
    public static Map<String,Decimal> getLimits(String accountId){

        try {
            Map<String,Decimal> mapResult = new Map<String,Decimal>();
            Decimal minimumAmount = Decimal.valueOf([SELECT Value__c FROM StandardValue__mdt WHERE DeveloperName = 'B2BMinimumOrderAmount' LIMIT 1][0].Value__c);
            Decimal maxPackPerMonth = CommerceCheckoutValidationController.getMaxMonthPacks();
            Decimal totalCustomerMonthPacks = CommerceCheckoutValidationController.getTotalCustomerMonthPacks(accountId) ?? 0;

            mapResult.put('minimumAmount', minimumAmount);
            mapResult.put('maxPackPerMonth', maxPackPerMonth);
            mapResult.put('totalCustomerMonthPacks', totalCustomerMonthPacks);
            return mapResult;
        } catch (Exception e) {
            throw new AuraHandledException(e.getMessage());
        }
    }

    public class CartItemValidation{
        @AuraEnabled
        public String cartItemId;
        @AuraEnabled
        public String message;

        CartItemValidation(String cartItemId, String message){
            this.cartItemId = cartItemId;
            this.message = message;
        }
    }

    @AuraEnabled
    public static List<CartItemValidation> checkCartItems(List<String> cartItemsIds){
        try {
            List<CartItemValidation> cartItemsValidations = new List<CartItemValidation>();
            for(CartItem ci : [SELECT Id, Name, MasterOrderItemId__r.AvailableQuantity__c, MasterOrderItemId__r.MasterOrderId__r.Status__c, MasterOrderItemId__r.MasterOrderId__r.Name, Quantity FROM CartItem WHERE Id IN :cartItemsIds]){
                if(ci.MasterOrderItemId__r.MasterOrderId__r.Status__c != 'Active'){
                    cartItemsValidations.add(new CartItemValidation(ci.Id, 'Negociação solicitada indiponível: ' + ci.MasterOrderItemId__r.MasterOrderId__r.Name + ', remova o item do carrinho e adicione novamente'));
                }
                else if(ci.MasterOrderItemId__r.AvailableQuantity__c < ci.Quantity){
                    cartItemsValidations.add(new CartItemValidation(ci.Id, 'Quantidade solicitada excede a disponível. Solicitada: '+ ci.Quantity +', Disponível: ' + ci.MasterOrderItemId__r.AvailableQuantity__c));
                }
            }
            return cartItemsValidations;
        } catch (Exception e) {
            throw new AuraHandledException(e.getMessage());
        }
    }


}