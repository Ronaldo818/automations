public without sharing class B2BCancelUnpaidOrdersSchedulable implements Schedulable{
    public void execute(SchedulableContext SC) {
        Database.executeBatch(new B2BCancelUnpaidOrdersBatch(), 1);
    }
}