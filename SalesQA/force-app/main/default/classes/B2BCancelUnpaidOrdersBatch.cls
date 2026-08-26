public without sharing class B2BCancelUnpaidOrdersBatch implements Database.Batchable<SObject>, Database.Stateful, Database.AllowsCallouts {


    public B2BCancelUnpaidOrdersBatch(){
    }


    public Database.QueryLocator start(Database.BatchableContext BC) {
        Datetime now = Datetime.now();
        String query = 'SELECT Id, MetodoPagamento__c FROM Order WHERE Type = \'B2B Commerce\' AND MetodoPagamento__c = \'pix\' AND PaymentExpiration__c <= :now AND Status = \'Activated\' AND StatusPagamento__c != \'paid\'';
       	return Database.getQueryLocator(query);
    }

    public void execute(Database.BatchableContext BC, List<Order> scope) {
        B2BAgrosysPedido.SendOrderParameter param = new B2BAgrosysPedido.SendOrderParameter();
        param.operacao = 4;
        param.orderId = scope[0].Id;
        param.formaPagamento = scope[0].MetodoPagamento__c;
        B2BAgrosysPedido.setPedido(System.JSON.serialize(param));
    }
    
    public void finish(Database.BatchableContext BC) {

    }
}