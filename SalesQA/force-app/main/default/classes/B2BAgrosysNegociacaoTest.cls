@IsTest
public class B2BAgrosysNegociacaoTest {
	
    @TestSetup
    static void makeData() {
        String uniqueUserName = 'istestuser' + DateTime.now().getTime() + '@testavivar.com';
        Profile customerProfile = [
            SELECT Id
            FROM Profile
            WHERE Name = 'Customer Community Plus User'
        ];
        
        Account newAccount = new Account(
			Name				= 'isTest ',
			CodigoClienteERP__c	= '999',
			Empresa__c 			= 1,
			unidade__c 			= 1
        );
        insert newAccount;
        
        Contact newContact = new Contact(
        	LastName = 'isTest',
            AccountId = newAccount.Id
        );
        insert newContact;
        
        User newUser = new User(
            Alias 				= 'istest',
            Email				= 'standarduser@testorg.com',
        	EmailEncodingKey	= 'UTF-8',
            LastName			= 'Testing',
            LanguageLocaleKey	= 'en_US',
        	LocaleSidKey		= 'en_US',
            ProfileId 			= customerProfile.Id,
        	TimeZoneSidKey		= 'America/Los_Angeles',
            ContactId			= newContact.Id,
         	UserName			= uniqueUserName,
            IsActive 			= true
        );
        insert newUser;
    }
    
    @IsTest
    static void getNegociacaoSucces01() {
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracao realizada com sucesso","msg_retorno":[{"tp_clineg":[{"temnegociacao":false}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
    	User u = [
        	SELECT Id
            FROM User
            WHERE IsActive = true
            AND Alias = 'istest'
            LIMIT 1
        ];
        
        System.runAs(u) {
            Test.startTest();
            Boolean hasNegociacao = true;
            hasNegociacao = B2BAgrosysNegociacao.getCustomerCode();
            
            System.assertEquals(false, hasNegociacao);
            Test.stopTest();
    	}       
    }
    
    @IsTest
    static void getNegociacaoSucces02() {
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
    	User u = [
        	SELECT Id
            FROM User
            WHERE IsActive = true
            AND Alias = 'istest'
            LIMIT 1
        ];
        
        System.runAs(u) {
            Test.startTest();
            Boolean hasNegociacao = true;
            hasNegociacao = B2BAgrosysNegociacao.getCustomerCode();
            
            System.assertEquals(null, hasNegociacao);
            Test.stopTest();
    	}       
    }
    
}