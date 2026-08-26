Global without sharing class B2BScheduleAgrosys implements Schedulable {
    Integer clienteUnidade = 1; 
        
    public void execute(SchedulableContext SC) {
        if(validateBatchJob()){
            if(!Test.isRunningTest()){
                Database.executeBatch(new B2BAgrosysClienteBatch(clienteUnidade), 1);
            }
        }
    }

    private static Boolean validateBatchJob() {
        String JOBTYPE = 'BatchApex';
        String APEXCLASSNAME = 'B2BAgrosysClienteBatch';
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