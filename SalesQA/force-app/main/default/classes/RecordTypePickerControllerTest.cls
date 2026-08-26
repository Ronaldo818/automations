@IsTest
public with sharing class RecordTypePickerControllerTest {
    
    @isTest
    public static void testGetRecordTypes() {
        Test.startTest();
        List<RecordType> result = RecordTypePickerController.getRecordTypes('Case');
        Test.stopTest();
        System.assert(result.size()>0);
    }
}