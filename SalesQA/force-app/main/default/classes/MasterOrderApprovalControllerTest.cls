@isTest(testFor='ApexClass:MasterOrderApprovalController')
public with sharing class MasterOrderApprovalControllerTest {

    @TestSetup
    static void makeData(){
        TestUtil util = new TestUtil();
        util.genMasterOrderData();
    }

    @IsTest
    static void getInitDataTest() {
        MasterOrderApprovalController.getinitData(getMasterOrderId());
    }

    @IsTest
    static void doApprovalTest() {
        MasterOrderApprovalController.doApproval(getMasterOrderId(), 'Active');
    }

    private static String getMasterOrderId(){
        return ([SELECT Id FROM MasterOrder__c LIMIT 1]?.Id);
    }


}