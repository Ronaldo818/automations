@IsTest
public with sharing class B2BCheckoutMinimumAmountControllerTest {

    @TestSetup
    static void makeData(){
        TestUtil util = new TestUtil('b2b');
        util.genAllData();
    }

    @IsTest
    public static void cartLimitesTest(){
        String accountId = [SELECT Id FROM Account LIMIT 1]?.Id;
	    B2BCheckoutMinimumAmountController.getLimits(accountId);
    }

    @IsTest
    public static void itemsLimitsTest(){
        List<String> ciIds = new List<String>();
        for(CartItem ci : [SELECT Id FROM CartItem WHERE Cart.Status != 'Closed' LIMIT 1]){
            MasterOrderItem__c moi = [SELECT Id, AvailableQuantity__c, ConsumedQuantity__c FROM MasterOrderItem__c LIMIT 1][0];
            ci.MasterOrderItemId__c = moi.Id;
            update ci;
            moi.ConsumedQuantity__c++;
            update moi;
            ciIds.add(ci.Id);
        }
	    B2BCheckoutMinimumAmountController.checkCartItems(ciIds);
    }

}