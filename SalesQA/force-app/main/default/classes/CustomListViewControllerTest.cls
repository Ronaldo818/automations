@IsTest
public  class CustomListViewControllerTest {
    @TestSetup
    static void makeData(){
        TestUtil util = new TestUtil();
        util.genMasterOrderData();
    }

    @IsTest
    static void testNoFilter(){
        Test.startTest();
            CustomListViewController.getObjectRecords('', 'Name', 'Account', 'Name', 'DESC');
        Test.stopTest();
    }

    @IsTest
    static void testFilter(){
        Test.startTest();
            CustomListViewController.getObjectRecords('Name LIKE \''+'%'+'Test'+'%'+'\'', 'Name', 'Account', 'Name', 'DESC');
        Test.stopTest();        
    }

    @IsTest
    static void testNoOrder(){
        Test.startTest();
            CustomListViewController.getObjectRecords('Name LIKE \''+'%'+'Test'+'%'+'\'', 'Name', 'Account', '', '');
        Test.stopTest();        
    }
}