public with sharing class CustomOrderListViewController {
    @AuraEnabled
    public static List<sObject> getObjectRecords(String filter, String fields, String objectName, String sortField, String sortDirect, String offset) {
        if(offset == null){
            offset = '0';
        }
        String baseQuery = '';
        System.debug(filter);
        if(filter == 'None'){
            baseQuery = 'SELECT Id,' + fields + '  FROM ' + objectName;                
            if(sortField != '' && sortField != null){
                sortField = sortField + ' ' + (String.isBlank(sortDirect)? 'ASC' : sortDirect);
                baseQuery += ' ORDER BY ' + sortField;
            }
        }
        else if(filter == 'Cliente'){
            String userId = UserInfo.getUserId();
            User currentUser = [SELECT ContactId, Contact.AccountId FROM User WHERE Id = :userId LIMIT 1];
            if(Test.isRunningTest()){
                currentUser.Contact = [SELECT Id, AccountId FROM Contact WHERE LastName ='testCon'];
            }
                baseQuery = 'SELECT Id,' + fields + '  FROM ' + objectName + ' WHERE AccountId = ' + '\'' + currentUser.Contact.AccountId + '\'';
                
                if(sortField != '' && sortField != null){
                    sortField = sortField + ' ' + (String.isBlank(sortDirect)? 'ASC' : sortDirect);
                    baseQuery += ' ORDER BY ' + sortField;
                }
        }
        else if(sortField != '' && sortField != null){
            sortField = sortField + ' ' + (String.isBlank(sortDirect)? 'ASC' : sortDirect);
            baseQuery = 'SELECT Id,' + fields + '  FROM ' + objectName + ' WHERE ' + filter + ' ORDER BY ' + sortField;
        }
        else {
            baseQuery = 'SELECT Id,' + fields + '  FROM ' + objectName + ' WHERE ' + filter; 
        }

        // add limit and offset
        baseQuery += ' LIMIT ' + 50 +' offset ' + offset;


        System.debug(baseQuery);
        return Database.query(baseQuery);
    }

    @AuraEnabled
    public static List<sObject> buscar(String filter, String fields) {
        

        User currentUser = [SELECT ContactId FROM User WHERE Id = :UserInfo.getUserId() LIMIT 1];
        if(Test.isRunningTest()){
            currentUser.ContactId = [SELECT Id FROM Contact WHERE LastName ='testCon'].Id;
        }
        System.debug('currentUser');
        System.debug(currentUser);

        String baseQuery = 'Select ' + fields + ' From Order WHERE';

        Decimal numberFilter;

        baseQuery += ' Account.name like \'%'+filter+'%\' OR orderNumber like \'%'+filter+'%\'  OR Status like \'%'+filter+'%\' OR Account.CodigoClienteERP__c like \'%'+filter+'%\'';
        if(filter.isNumeric()){
            numberFilter = Decimal.valueOf(filter);
            // baseQuery += ' OR CodigoPedidoERP__c = :numberFilter';
        }
        System.debug(baseQuery);
        return Database.query(baseQuery);
    }

}