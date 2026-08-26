@IsTest
public with sharing class B2BAccountSelectorTest {

    @TestSetup
    static void makeData(){
        TestUtil util = new TestUtil();
        util.genAccountContactData();

        // EstruturaVenda__c ev = new EstruturaVenda__c (

        // );
        // insert ev;

    }

    @IsTest
    public static void mainTest(){
        B2BAccountSelector.getCurrentUserAccount(UserInfo.getUserId());
        B2BAccountSelector.getRelatedAccounts(UserInfo.getUserId());
        B2BAccountSelector.provisoryCodeCoverageMethod();
    }

}