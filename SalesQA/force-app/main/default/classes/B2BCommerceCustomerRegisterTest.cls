@IsTest
public class B2BCommerceCustomerRegisterTest {
    private static final String CNPJ_ALFA_LIMPO = 'RX5TCT6K000177';
    private static final String CNPJ_ALFA_MASCARADO = 'RX.5TC.T6K/0001-77';

	
    @TestSetup
    static void makeData() {
    	 Lead newLead = new Lead(
            LastName 				= 'isTest',
        	Company 				= 'isTest',
            Phone 					= '99999999999',
            PostalCode 				= 'isTest',
            Street 					= 'isTest',
            City 					= 'isTest',
            State 					= 'isTest',
        	NomeFantasia__c 		= 'isTest',
            CNPJInput__c 			= '32313227000183',
            RamoAtividade__c 		= '4',
            NumeroEndereco__c 		= '00',
            BairroEndereco__c 		= 'isTest',
            ComplementoEndereco__c	= 'isTest'
        );
        insert newLead;
        
        Account newAccount = new Account(
        	Name 					= 'isTest',
        	Unidade__c 				= 1,
        	Empresa__c 				= 1,
        	CodigoClienteERP__c		= '1',
            CNPJInput__c 				= '32313227000183',
        	Ocorrencia__c 			= 1,
        	Lista__c 				= 27
        );
        insert newAccount;
    }
    
    @IsTest
    static void isValidCNPJSuccess01() {
        test.startTest();
		B2BCommerceCustomerRegisterController.isValidCNPJ('32313227000183');
        test.stopTest();
    }
    
    @IsTest
    static void isValidCNPJError01() {
        test.startTest();
		B2BCommerceCustomerRegisterController.isValidCNPJ('323132270001');
        test.stopTest();
    }
    
    @IsTest
    static void verifyCustomerSuccess01() {
        test.startTest();
		B2BCommerceCustomerRegisterController.verifyCustomer('26162472000134');
        test.stopTest();
    }
    
    @IsTest
    static void isValidCNPJAlfanumerico01() {
        test.startTest();
        String resultJson = B2BCommerceCustomerRegisterController.isValidCNPJ(CNPJ_ALFA_MASCARADO);
        test.stopTest();

        Map<String, Object> result = (Map<String, Object>) JSON.deserializeUntyped(resultJson);
        System.assertEquals(CNPJ_ALFA_LIMPO, (String) result.get('cnpj'));
        System.assertEquals(true, result.get('valid'));
    }

    @IsTest
    static void verifyCustomerLeadSuccess01() {
        final String CNPJ = '26162472000134';
        Lead lead = [
            SELECT Id, CNPJInput__c
            FROM Lead
            LIMIT 1
        ];
        lead.CNPJInput__c = CNPJ;
        update lead;
        
        test.startTest();
		B2BCommerceCustomerRegisterController.verifyCustomer(CNPJ);
        test.stopTest();
    }
    
    @IsTest
    static void verifyCustomerAccountSuccess01() {
        final String CNPJ = '26162472000134';
        Account acc = [
            SELECT Id, CNPJInput__c
            FROM Account
            LIMIT 1
        ];
        acc.CNPJInput__c = CNPJ;
        update acc;
        
        test.startTest();
		B2BCommerceCustomerRegisterController.verifyCustomer(CNPJ);
        test.stopTest();
    }

    @IsTest
    static void verifyCustomerLeadAlfanumerico01() {
        Lead lead = [
            SELECT Id, CNPJ__c, CNPJInput__c
            FROM Lead
            LIMIT 1
        ];
        lead.CNPJ__c = CNPJ_ALFA_LIMPO;
        lead.CNPJInput__c = CNPJ_ALFA_LIMPO;
        update lead;

        test.startTest();
        String resultJson = B2BCommerceCustomerRegisterController.verifyCustomer(CNPJ_ALFA_MASCARADO);
        test.stopTest();

        Map<String, Object> result = (Map<String, Object>) JSON.deserializeUntyped(resultJson);
        System.assertEquals(false, result.get('valid'));
        System.assertEquals(String.valueOf(lead.Id), (String) result.get('leadId'));
    }

    @IsTest
    static void verifyCustomerAccountAlfanumerico01() {
        Account acc = [
            SELECT Id, CNPJ__c, CNPJInput__c
            FROM Account
            LIMIT 1
        ];
        acc.CNPJ__c = CNPJ_ALFA_LIMPO;
        acc.CNPJInput__c = CNPJ_ALFA_LIMPO;
        update acc;

        test.startTest();
        String resultJson = B2BCommerceCustomerRegisterController.verifyCustomer(CNPJ_ALFA_MASCARADO);
        test.stopTest();

        Map<String, Object> result = (Map<String, Object>) JSON.deserializeUntyped(resultJson);
        System.assertEquals(false, result.get('valid'));
        System.assertEquals(String.valueOf(acc.Id), (String) result.get('accountId'));
    }
    
    @IsTest
    static void createLeadSuccess01() {
        B2BCommerceCustomerRegisterController.LeadData data = new B2BCommerceCustomerRegisterController.LeadData();
        data.nomeResponsavel 		= 'isTest';
        data.sobrenomeResponsavel	= 'isTest';
        data.telefoneResponsavel 	= '99999999999';
        data.razaoSocial 			= 'isTest';
        data.cnpj 					= '32313227000183';
        data.nomeFantasia 			= 'isTest';
        data.email 					= 'istest.istest@istest.com';
        data.telefone				= '99999999999';
        data.ramoAtividade 			= '4';
        data.categoria 				= '1';
        data.subcategoria 			= '1-1';
        data.rua 					= 'isTest';
        data.bairro 				= 'isTest';
        data.cidade 				= 'isTest';
        data.estado 				= 'isTest';
        data.numero 				= 'isTest';
        data.cep 					= 'isTest';
        data.complemento 			= 'isTest';
        
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo" : 1,"msg_descricao":"Integração realizada com sucesso","msg_retorno":[]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        test.startTest();
		String isTest = B2BCommerceCustomerRegisterController.createLead(System.JSON.serialize(data));
        test.stopTest();
    }

    @IsTest
    static void createLeadAlfanumerico01() {
        B2BCommerceCustomerRegisterController.LeadData data = new B2BCommerceCustomerRegisterController.LeadData();
        data.nomeResponsavel        = 'isTest';
        data.sobrenomeResponsavel   = 'isTest';
        data.telefoneResponsavel    = '99999999999';
        data.razaoSocial            = 'isTest';
        data.cnpj                   = CNPJ_ALFA_MASCARADO;
        data.nomeFantasia           = 'isTest';
        data.email                  = 'istest.istest@istest.com';
        data.telefone               = '99999999999';
        data.ramoAtividade          = '4';
        data.categoria              = '1';
        data.subcategoria           = '1-1';
        data.rua                    = 'isTest';
        data.bairro                 = 'isTest';
        data.cidade                 = 'isTest';
        data.estado                 = 'isTest';
        data.numero                 = 'isTest';
        data.cep                    = 'isTest';
        data.complemento            = 'isTest';

        HttpCalloutMockTools mock = new HttpCalloutMockTools();
        mock.setResponseBody('{"tp_retorno":[{"msg_codigo" : 1,"msg_descricao":"Integração realizada com sucesso","msg_retorno":[]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);

        test.startTest();
        String leadId = B2BCommerceCustomerRegisterController.createLead(System.JSON.serialize(data));
        test.stopTest();

        Lead createdLead = [
            SELECT Id, CNPJInput__c
            FROM Lead
            WHERE Id = :leadId
            LIMIT 1
        ];
        System.assertEquals(CNPJ_ALFA_LIMPO, createdLead.CNPJInput__c);
    }
    
}