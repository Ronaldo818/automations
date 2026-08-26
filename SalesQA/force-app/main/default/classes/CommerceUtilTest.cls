@isTest
private class CommerceUtilTest {
    @TestSetup
    static void makeData(){
        TestUtil util = new TestUtil();

        try{
            util.genStoreData();
            util.genOrderData();
        } catch (Exception ex){
            
        }
    }

    @IsTest
    static void test(){
        
        Test.startTest();
            CommerceUtil.getCartSummary();
            CommerceUtil.getCurrentEffectiveAccountId();
            CommerceUtil.getSite();
            CommerceUtil.getStore();
            CommerceUtil.getStoreName();
            CommerceUtil.getWishListItems();
            CommerceUtil.provisoryCodeCoverageMethod();
            CommerceUtil.AccountExist('');
        Test.stopTest();
        
    }

    @IsTest
    static void testAccountExistAlfanumerico() {
        Account acc = new Account(
            Name = 'Teste Alfa',
            CNPJ__c = 'AB12CD34000181',
            CNPJInput__c = 'AB12CD34000181',
            ExternalId__c = 'AB12CD34000181'
        );
        insert acc;

        Test.startTest();
            CommerceUtil.AccountExist('AB.12C.D34/0001-81');
        Test.stopTest();
    }


}