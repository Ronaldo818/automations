@IsTest
public with sharing class ScheduleCreateCommerceRecordsTest {
    
    @IsTest
    static void startScheduleSuccess01() {
        Test.startTest();
        ScheduleCreateCommerceRecords.startSchedule();
        Test.stopTest();
    }
}