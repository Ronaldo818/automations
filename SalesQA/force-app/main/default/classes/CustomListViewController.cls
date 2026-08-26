public with sharing class CustomListViewController {
    @AuraEnabled
    public static List<sObject> getObjectRecords(String filter, String fields, String objectName, String sortField, String sortDirect) {
        String baseQuery = '';
        List<String> filterData = filter.split(',');
        // for custom filters:
        List<String> accs;

        if(String.isNotBlank(filter)){
            filter = ' WHERE ' + filter;
        }
        else{
            filter = '';
        }

        if(sortField != '' && sortField != null){
            sortField = sortField + ' ' + (String.isBlank(sortDirect)? 'ASC' : sortDirect);
            baseQuery = 'SELECT Id,' + fields + '  FROM ' + objectName + filter + ' ORDER BY ' + sortField + ' LIMIT 1000';
        }
        else {
            baseQuery = 'SELECT Id,' + fields + '  FROM ' + objectName + filter + ' LIMIT 1000';
        }


        System.debug(baseQuery);
        return Database.query(baseQuery);
    }
}