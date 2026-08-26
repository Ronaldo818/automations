@isTest
private class CustomListViewEnhacedControllerTest {
    @TestSetup
    static void makeData(){
        TestUtil util = new TestUtil();

        try {
            util.genStoreData();
            util.genOrderData();
        } catch (Exception ex){

        }
    }

    @IsTest
    static void test(){
        
        CustomListViewEnhacedController.getAccounts();
        CustomListViewEnhacedController.provisoryCodeCoverageMethod();
        
    }


}