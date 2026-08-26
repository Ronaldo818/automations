public without sharing class B2BAgrosysPricebookBatch implements Database.Batchable<SObject>, Database.Stateful, Database.AllowsCallouts {

    class CustomException extends Exception {}
    List<Account> lstAccount;
    
    public B2BAgrosysPricebookBatch(List<Account> lstAccount) {
        this.lstAccount = lstAccount;
    }

    public Iterable<Account> start(Database.BatchableContext BC) {
           return this.lstAccount;
    }

    public void execute(Database.BatchableContext BC, List<Account> scope) {
        for(Account acc : scope) {
            try {
                if(Test.isRunningTest() && acc.Name.equalsIgnoreCase('iserror')) {
                    throw new AuraHandledException('isError');
                }

                B2bGerenciadorPricebook.sincronizarAccount(acc);
            }
            catch(Exception ex) {
                ErrorLog__e errorLog = new ErrorLog__e();
                errorLog.AutomationType__c = 'Apex';
                errorLog.AutomationName__c = 'B2bGerenciadorPricebook';
                errorLog.Status__c = 'Erro';
                errorLog.Name__c = 'ERROR B2bGerenciadorPricebook';

                if(acc != null) {
                    errorLog.Name__c += ' - ' + acc.CodigoClienteERP__c;
                    errorLog.ErrorMessage__c += 'Account: ' + acc.Name + ' - ' + acc.Id + '\n\n';
                }

                errorLog.ErrorMessage__c += 'Método: sincronizar' + '\n\nType: ' + ex.getTypeName() + '\n\nCause: ' + ex.getCause() + '\n\nLine Number: ' + ex.getLineNumber() + '\n\nStack Trace: ' + ex.getStackTraceString() + '\n\nMessage: ' + ex.getMessage();

                String eventUUID;
                Database.SaveResult sr = EventBus.publish(errorLog);
                if (sr.isSuccess()) {
                    eventUUID = EventBus.getOperationId(sr);
                } else {
                    for(Database.Error err : sr.getErrors()) {
                        System.debug('Error returned: ' + err.getStatusCode() + ' - ' + err.getMessage());
                    }
                }
                
                if(!Test.isRunningTest()) {
                    throw new CustomException('Ocorreu um erro, por favor tente novamente. Se o erro persistir informe ao representante de vendas ou administrador do sistema junto com o UUID do erro: ' + eventUUID);
                }
            }
        }
    }
 
    public void finish(Database.BatchableContext BC) {}
}