public without sharing class CustomMasterOrderListViewController {
    @AuraEnabled
    public static List<sObject> getObjectRecords(String filter, String fields, String objectName, String sortField, String sortDirect) {
        String baseQuery = '';
        List<String> filterData = filter.split(',');
        // for custom filters:
        List<String> accs;

        if(filterData[0] == 'custom'){
            switch on filterData[1] {
                when 'fvMasterOrder' {
                    if(filterData.size() != 3){
                        throw new AuraHandledException('No effectiveAccountId provided, should be: custom,fvMasterOrder,{effectiveAccountId}' );
                    }
                    accs = new List<String>{filterData[2]};
                    Account acc = [SELECT Id, ParentId FROM Account WHERE Id = :filterData[2] LIMIT 1];
                    if(acc.ParentId != null){
                        accs.add(acc.ParentId);
                    }
                    filter = ' WHERE AccountId__c IN :accs';
                }
                when else {
                    throw new AuraHandledException('No custom filter named ' + filterData[1] + ', use: custom,{filterName},data' );
                }
            }
        }
        else if(String.isNotBlank(filter)){
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