@IsTest
public with sharing class B2BAgrosysPricebookBatchTest {
    
    @TestSetup
    static void makeData() {
        WebStore newWebStore = new WebStore(
            Name = 'Avivar LWR'
        );
        insert newWebStore;

        Account newAccount = new Account(
            Name				= 'isTest',
            CodigoClienteERP__c	= '123',
            Unidade__c			= 1,
            Empresa__c			= 1,
			Ocorrencia__c		= 1,
        	Lista__c			= 1
        );
        insert newAccount;

        Product2 newProduct = new Product2 (
            Name 				= 'isTest',
            ProductCode 		= '10101',
            Family 				= '035',
            IsActive 			= true,
            Description 		= 'isTest',
            CodigoProdutoERP__c	= '10101',
            PesoUnitario__c 	= 1,
            UnidadeMedida__c 	= 'Kg',
            StockKeepingUnit 	= ' 10101'
        );
        insert newProduct;
        
        delete [SELECT Id FROM BuyerAccount];
        BuyerAccount newBuyerAccount = new BuyerAccount(
        	Name = newAccount.Name,
            BuyerId = newAccount.Id,
            IsActive = true
        );
		insert newBuyerAccount;
    }
    
    @IsTest
    static void b2bAgrosysPricebookBatchSuccess01() {
        Account account = [
            SELECT Id, Name, Empresa__c, Unidade__c, Ocorrencia__c, CodigoClienteERP__c
            FROM Account
            LIMIT 1
        ];

        Product2 product = [
            SELECT Id, CodigoProdutoERP__c
            FROM Product2
            LIMIT 1
        ];
        
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo" : 1,"msg_descricao":"Integracao realizada com sucesso","msg_retorno":[{"tp_tabpreco": [ { "emp_empresa": ' + account.Empresa__c + ', "uni_unidade": ' + account.Unidade__c + ', "cli_codigo": ' + account.CodigoClienteERP__c + ', "pre_tiporeg": 1, "pit_tabpre": ' + account.Ocorrencia__c + ', "pro_codigo": ' + product.CodigoProdutoERP__c + ', "pre_dtatual": "2001-01-01", "pre_tabela": 1, "pre_vlrst": 1, "pre_vfcpst": 1 }]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        Test.startTest();
        B2BAgrosysPricebookBatch newBatch = new B2BAgrosysPricebookBatch(new List<Account> {account});
        Database.executeBatch(newBatch, 1);
        Test.stopTest();
    }

    @IsTest
    static void b2bAgrosysPricebookBatchError01() {
        Account account = [
            SELECT Id, Name, Empresa__c, Unidade__c, Ocorrencia__c, CodigoClienteERP__c
            FROM Account
            LIMIT 1
        ];

        account.Name = 'isError';
        update account;

        Product2 product = [
            SELECT Id, CodigoProdutoERP__c
            FROM Product2
            LIMIT 1
        ];
        
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo" : 1,"msg_descricao":"Integracao realizada com sucesso","msg_retorno":[{"tp_tabpreco": [ { "emp_empresa": ' + account.Empresa__c + ', "uni_unidade": ' + account.Unidade__c + ', "cli_codigo": ' + account.CodigoClienteERP__c + ', "pre_tiporeg": 1, "pit_tabpre": ' + account.Ocorrencia__c + ', "pro_codigo": ' + product.CodigoProdutoERP__c + ', "pre_dtatual": "2001-01-01", "pre_tabela": 1, "pre_vlrst": 1, "pre_vfcpst": 1 }]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        Test.startTest();
        B2BAgrosysPricebookBatch newBatch = new B2BAgrosysPricebookBatch(new List<Account> {account});
        Database.executeBatch(newBatch, 1);
        Test.stopTest();
    }
    
}