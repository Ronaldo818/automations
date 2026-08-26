@IsTest
public class B2BAgrosysPreCadastroTest {
	
    @TestSetup
    static void makeData() {
        Lead newLead = new Lead(
            LastName = 'isTest',
        	Company = 'isTest',
            Phone = '99999999999',
            PostalCode = 'isTest',
            Street = 'isTest',
            City = 'isTest',
            State = 'isTest',
        	NomeFantasia__c = 'isTest',
            CNPJInput__c = '26162472000134',
            RamoAtividade__c = '4',
            NumeroEndereco__c = '00',
            BairroEndereco__c = 'isTest',
            ComplementoEndereco__c = 'isTest'
        );
        insert newLead;
    }
    
    @IsTest
    static void setCustomerSuccess01() {
        Lead lead = [
            SELECT Company, Phone, PostalCode, Street, City, State, NomeFantasia__c, CNPJInput__c, RamoAtividade__c, NumeroEndereco__c, BairroEndereco__c, ComplementoEndereco__c
            FROM Lead
            LIMIT 1
        ];
        
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo" : 1,"msg_descricao":"Integração realizada com sucesso","msg_retorno":[]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
    	B2BAgrosysPreCadastro preCadastro = new B2BAgrosysPreCadastro();
        
        test.startTest();
        preCadastro.setCustomer(lead);
        test.stopTest();
    }
    
     @IsTest
    static void setCustomerError01() {
        Lead lead = [
            SELECT Company, Phone, PostalCode, Street, City, State, NomeFantasia__c, CNPJInput__c, RamoAtividade__c, NumeroEndereco__c, BairroEndereco__c, ComplementoEndereco__c
            FROM Lead
            LIMIT 1
        ];
        
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo" : 9999,"msg_descricao":"Integração realizada com sucesso","msg_retorno":[]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
    	B2BAgrosysPreCadastro preCadastro = new B2BAgrosysPreCadastro();
        
        test.startTest();
        preCadastro.setCustomer(lead);
        test.stopTest();
    }

}