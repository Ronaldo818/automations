global with sharing class RecordTypePickerController {

    @AuraEnabled(cacheable=true)
    public static List<RecordType> getRecordTypes(String sObjectApiName) {
    
        Set<Id> availableRTForSelectedObject = RecordTypePickerController.availableRecordTypeIdsFromObjectApiNames(
            new List<String>{ sObjectApiName }
        );

        return [
            SELECT Id, Name, Description, DeveloperName 
            FROM RecordType 
            WHERE Id IN :availableRTForSelectedObject
        ];
    }


    private static Set<Id> availableRecordTypeIdsFromObjectApiNames(List<String> objectApiNames){
        Set<Id> resultRecordTypeIds = new Set<Id>();
        
        for (Schema.DescribeSObjectResult describeInfo : Schema.describeSObjects(objectApiNames)){
            for (Schema.RecordTypeInfo rtInfo : describeInfo.getRecordTypeInfos()){
                if (rtInfo.isActive() && rtInfo.isAvailable()) {
                    resultRecordTypeIds.add(rtInfo.getRecordTypeId());
                    }
                }
            }
        return resultRecordTypeIds;
    }
}