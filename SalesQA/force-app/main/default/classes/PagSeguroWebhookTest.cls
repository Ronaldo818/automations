@IsTest
public class PagSeguroWebhookTest {
	
    @TestSetup
    static void makeData() {
        Pricebook2 newPriceBook = new Pricebook2(Name = 'isTest');
		insert newPriceBook;
        
        CondicaoPagamento__c newCondPagto = new CondicaoPagamento__c(
        	AVista__c = true,
        	CodigoCondicaoPagamento__c = '1',
        	Descricao__c = 'isTest',
        	Name = 'isTest',
        	NumeroParcelas__c = 1);
        insert newCondPagto;
        
        Account newAccount = new Account(
        	Name = 'isTest',
        	Unidade__c = 1,
        	Empresa__c = 1,
        	CodigoClienteERP__c = '1',
        	CondicaoPagamento__c = newCondPagto.Id,
        	Ocorrencia__c = 1,
        	Lista__c = 27);
        insert newAccount;
        
        Order newOrder = new Order(
        	EffectiveDate = Date.today(),
        	Status = 'Draft',
        	ValorCashback__c = 2,
        	AccountId = newAccount.Id,
        	Pricebook2Id = newPriceBook.Id);
        insert newOrder;
        
        Profile profile = [SELECT Id FROM Profile WHERE Name='System Administrator' LIMIT 1];
        
        User newUser = new User(
            Alias='istest',
            username='istest@istest.com',
            EmailEncodingKey='UTF-8',
            LanguageLocaleKey='pt_BR',
			LocaleSidKey='pt_BR',
            TimeZoneSidKey='America/Sao_Paulo',
            ProfileId = profile.Id,
            lastname='isTest',
			email='teste@teste.com',
            phone='+5511999999999',
            IsActive = true
        );
        insert newUser;
    }
    
    @IsTest
    static void handlePostSuccess01() {
        User u = [
        	SELECT Id, lastname
            FROM User
            WHERE IsActive = true
            AND Alias = 'istest'
        ];
        
        u.lastname = 'Success';
        update u;
        
        System.runAs(u) {
			PagSeguroWebhook.handlePost();    
        }
    }
    
    @IsTest
    static void handlePostError01() {
        User u = [
        	SELECT Id, lastname
            FROM User
            WHERE IsActive = true
            AND Alias = 'istest'
        ];
        
        u.lastname = 'Error';
        update u;
        
        System.runAs(u) {
			PagSeguroWebhook.handlePost();    
        }
    }
    
}