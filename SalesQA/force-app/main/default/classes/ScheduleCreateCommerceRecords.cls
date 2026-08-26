public without sharing class ScheduleCreateCommerceRecords implements Schedulable {

    public void execute(SchedulableContext SC) {
        // Se este schedule não estiver em execução, iniciar o processo.
        if (validateAsyncApexJob()) {
            List<Account> lstAccount = [
                SELECT Id, Name, Unidade__c, CodigoClienteERP__c, ClienteEcommerce__c
	            FROM Account
	            WHERE Id IN (
	            	SELECT BuyerId
                 	FROM BuyerAccount
	            	WHERE PendingEcommerceRecords__c = true
	            )
                AND IsBuyer = true
            ];

            if(lstAccount.size() > 0) {
                B2BAgrosysPricebookBatch pricebookBatch = new B2BAgrosysPricebookBatch(lstAccount);
                if(!Test.isRunningTest()){
                    Database.executeBatch(pricebookBatch, 1);
                }
            }
        }
    }

    public static void startSchedule() {
        ScheduleCreateCommerceRecords routine = new ScheduleCreateCommerceRecords();
        Datetime systemNow = System.now().addMinutes(1);
        String sch = '0 ' + String.valueof(systemNow.minute())+ ' ' + String.valueof(systemNow.hour()) + ' ' + String.valueof(systemNow.day()) + ' ' + String.valueof(systemNow.month()) + ' ? ' + String.valueof(systemNow.year());

        System.schedule('New B2BAgrosysPricebookBatch - ' + systemNow, sch, routine);
    }

    public static Boolean validateAsyncApexJob() {
        return validateScheduleJob() && validateBatchJob();
    }

    private static Boolean validateScheduleJob() {
        String JOBTYPE = 'ScheduledApex';
        String APEXCLASSNAME = Test.isRunningTest() ? 'isTest' : 'ScheduleCreateCommerceRecords';
        
        Integer count = [
            SELECT COUNT()
            FROM AsyncApexJob
            WHERE JobType = :JOBTYPE
            AND ApexClass.Name = :APEXCLASSNAME
            AND CronTrigger.NextFireTime >= :System.now()
        ];

        if(count == 0) {
            return true;
        }
        else {
            return false;
        }
    }

    private static Boolean validateBatchJob() {
        String JOBTYPE = 'BatchApex';
        String APEXCLASSNAME = Test.isRunningTest() ? 'isTest' : 'B2BAgrosysPricebookBatch';
        List<String> JOBSTATUS = new List<String> {'Processing', 'Preparing'};
    
        Integer count = [
            SELECT COUNT()
            FROM AsyncApexJob
            WHERE JobType = :JOBTYPE
            AND ApexClass.Name = :APEXCLASSNAME
            AND Status IN :JOBSTATUS
        ];
    
        if(count == 0) {
            return true;
        }
        else {
            return false;
        }
    }

}