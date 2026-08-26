@IsTest(SeeAllData=False)
public class B2bGerenciadorPricebookTest {
	
    @TestSetup
    static void makeData() {
        UserRole userrole = [Select Id, DeveloperName From UserRole Where DeveloperName = 'Admin' Limit 1];

        User adminUser = [Select Id, UserRoleId From User Where Profile.Name='System Administrator' AND IsActive = true Limit 1];

        adminUser.UserRoleId = userRole.Id;
        update adminUser;
        System.runAs(adminUser){

        String profileId = [
            SELECT Id
            FROM Profile
            WHERE Name = 'Customer Community Plus User'
            LIMIT 1
        ].Id;
   
        Account newAccount = new Account(
            Name				= 'isTest',
            CodigoClienteERP__c	= '123',
            Unidade__c			= 1,
            Empresa__c			= 1,
			Ocorrencia__c		= 1,
        	Lista__c			= 1
        );
        insert newAccount;
                 
        Contact newContact = new Contact(
            LastName 	= 'Teste',
            AccountId 	= newAccount.id,
            Email 		= 'teste.teste@teste3.com',
        	MobilePhone	= '0000000000',
        	Phone 		= '0000000000');
        insert newContact;
        
        ContactPointAddress newCPA = new ContactPointAddress(
            Name 		= 'isTest',
            Street 		= 'isTest', 
            City 		= 'isTest', 
            State 		= 'isTest', 
            PostalCode	= '101010', 
            Country 	= 'BR', 
            UsageType 	= 'Shipping', 
            IsDefault 	= true,
            AddressType = 'Shipping',
            ParentId 	= newAccount.Id
        );
        insert newCPA;

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

        UnidadeProduto__c newUnidadeProduto = new UnidadeProduto__c();
        newUnidadeProduto.Key__c = 'isTest';
        newUnidadeProduto.Unidade__c = 1;
        insert newUnidadeProduto;
        
        WebStore newWebStore = new WebStore(
            Name = 'Avivar LWR'
        );
        insert newWebStore;
        
        delete [SELECT Id FROM BuyerAccount];
        BuyerAccount newBuyerAccount = new BuyerAccount(
        	Name = newAccount.Name,
            BuyerId = newAccount.Id,
            IsActive = true
        );
		insert newBuyerAccount;
		
		String key = newAccount.CodigoClienteERP__c;
		Pricebook2 newPricebook = new Pricebook2();
		newPricebook.Empresa__c             = newAccount.Empresa__c;
		newPricebook.Unidade__c             = newAccount.Unidade__c;
		newPricebook.Ocorrencia__c          = newAccount.Ocorrencia__c;
		newPricebook.CodigoClienteERP__c    = Integer.valueOf(newAccount.CodigoClienteERP__c);
		newPricebook.AccountId__c           = newAccount.Id;
        newPricebook.Lista__c               = newAccount.Lista__c;
		newPricebook.Name                   = 'isTest ' + key;
		newPricebook.IsActive               = true;
		insert newPricebook;

        PricebookEntry newStandardPBE = new PricebookEntry(); 
        newStandardPBE.Pricebook2Id = Test.getStandardPricebookId();
        newStandardPBE.Product2Id = newProduct.Id;
        newStandardPBE.UnitPrice = 0;
        newStandardPBE.IsActive = true;
        insert newStandardPBE;
        
        BuyerGroup newBuyerGroup = new BuyerGroup();
		newBuyerGroup.Name      = newPricebook.Name;
		newBuyerGroup.Key__c    = key;
        insert newBuyerGroup;
        
        CommerceEntitlementPolicy newCommerceEntitlementPolicy = new CommerceEntitlementPolicy();
		newCommerceEntitlementPolicy.IsActive = true;
		newCommerceEntitlementPolicy.CanViewPrice = true;
		newCommerceEntitlementPolicy.CanViewProduct = true;
		newCommerceEntitlementPolicy.Name = newPricebook.Name;
		insert newCommerceEntitlementPolicy;

        CommerceEntitlementProduct newCommerceEntitlementProduct = new CommerceEntitlementProduct();
        newCommerceEntitlementProduct.PolicyId = newCommerceEntitlementPolicy.Id;
        newCommerceEntitlementProduct.ProductId = newProduct.Id;
        newCommerceEntitlementProduct.Key__c = key +  '-' + newProduct.CodigoProdutoERP__c;
        insert newCommerceEntitlementProduct;

		WebStore store = [SELECT Id FROM WebStore WHERE Name = 'Avivar LWR'];
       	BuyerGroupPricebook newBuyerGroupPricebook = new BuyerGroupPricebook();
		newBuyerGroupPricebook.BuyerGroupId = newBuyerGroup.Id;
		newBuyerGroupPricebook.Pricebook2Id = newPricebook.Id;
		newBuyerGroupPricebook.IsActive = true;
		insert newBuyerGroupPricebook;
                    
		WebStoreBuyerGroup storeLink = new WebStoreBuyerGroup();
		storeLink.WebStoreId = store.Id;
		storeLink.BuyergroupId = newBuyerGroup.Id;
        insert storeLink;

		BuyerGroupMember newBuyerGroupMember = new BuyerGroupMember();
		newBuyerGroupMember.BuyerGroupId = newBuyerGroup.Id;
		newBuyerGroupMember.BuyerId = newAccount.Id;
		insert newBuyerGroupMember;

		CommerceEntitlementBuyerGroup newCommerceEntitlementBuyerGroup = new CommerceEntitlementBuyerGroup();
		newCommerceEntitlementBuyerGroup.BuyerGroupId = newBuyerGroup.Id;
		newCommerceEntitlementBuyerGroup.PolicyId = newCommerceEntitlementPolicy.Id;
		insert newCommerceEntitlementBuyerGroup;

        
        User newUser = new User(
            Alias 				= 'isTest',
            username 			= 'istest.istest@istest.com',
            EmailEncodingKey 	= 'UTF-8',
            LanguageLocaleKey	= 'pt_BR',
			LocaleSidKey 		= 'pt_BR',
            TimeZoneSidKey 		= 'America/Sao_Paulo',
            ProfileId 			= profileId,
            lastname 			= 'Teste X',
            email 				= 'teste@teste31.com',
            phone 				= '+5571987000346',
            ContactId 			= newContact.Id
        );
        insert newUser;
        }
    }

    // Allefy - Cenário em que NÃO EXISTEM registros de Pricebook2, PricebookEntry, BuyerGroup e CommerceEntitlementPolicy o método B2bGerenciadorPricebook.sincronizar(List<Account> lstAccount) CRIA os registros necessários para o E-commerce.
    @IsTest
    static void sincronizarSuccess01() {
        delete [
            SELECT Id
            FROM CommerceEntitlementPolicy
        ];

        delete [
            SELECT Id
            FROM BuyerGroup
        ];

        delete [
            SELECT Id
            FROM PricebookEntry
        ];

        delete [
            SELECT Id
            FROM Pricebook2
        ];

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
        
        test.startTest();
        B2bGerenciadorPricebook.sincronizar(new List<Account> {account});
        test.stopTest();
    }

    // Allefy - Cenário em que EXISTEM registros de Pricebook2, PricebookEntry, BuyerGroup e CommerceEntitlementPolicy o método B2bGerenciadorPricebook.sincronizar(List<Account> lstAccount) ATUALIZA os registros de PricebookEntry relacionados à Conta.
    @IsTest
    static void sincronizarSuccess02() {
        User us = [
            SELECT Id
            FROM User
            WHERE Alias = 'isTest'
        ];
        
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

        UnidadeProduto__c unidadeProduto = [
            SELECT Id, Key__c, Unidade__c
            FROM UnidadeProduto__c
            LIMIT 1
        ];

        unidadeProduto.Key__c = account.Unidade__c + '-' + product.CodigoProdutoERP__c;
        update unidadeProduto;

        Pricebook2 priceBook = [
            SELECT Id
            FROM Pricebook2
            WHERE Id != :Test.getStandardPricebookId()
            LIMIT 1
        ];

        Decimal preTabela = 1;
        String key = account.CodigoClienteERP__c + '-' + product.CodigoProdutoERP__c;

        PricebookEntry newPricebookEntry = new PricebookEntry();
        newPricebookEntry.Product2Id = product.Id;
        newPricebookEntry.Pricebook2Id = priceBook.Id;
        newPricebookEntry.UnitPrice = preTabela;
        newPricebookEntry.KeyBuyerGroup__c = key;
        insert newPricebookEntry;
        
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo" : 1,"msg_descricao":"Integracao realizada com sucesso","msg_retorno":[{"tp_tabpreco": [ { "emp_empresa": ' + account.Empresa__c + ', "uni_unidade": ' + account.Unidade__c + ', "cli_codigo": ' + account.CodigoClienteERP__c + ', "pre_tiporeg": 1, "pit_tabpre": ' + account.Ocorrencia__c + ', "pro_codigo": ' + product.CodigoProdutoERP__c + ', "pre_dtatual": "2001-01-01", "pre_tabela": ' + preTabela + ', "pre_vlrst": 1, "pre_vfcpst": 1 }]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        test.startTest();
        System.runAs(us) {
        	//B2bGerenciadorPricebook.sincronizar(new List<Account> {account});   
        }
        test.stopTest();
    }

    @IsTest
    static void sincronizarSuccess03() {
        User us = [
            SELECT Id
            FROM User
            WHERE Alias = 'isTest'
        ];
        
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

        UnidadeProduto__c unidadeProduto = [
            SELECT Id, Key__c, Unidade__c
            FROM UnidadeProduto__c
            LIMIT 1
        ];

        unidadeProduto.Key__c = account.Unidade__c + '-' + product.CodigoProdutoERP__c;
        update unidadeProduto;

        Pricebook2 priceBook = [
            SELECT Id
            FROM Pricebook2
            WHERE Id != :Test.getStandardPricebookId()
            LIMIT 1
        ];

        Decimal preTabela = 0;
        String key = account.CodigoClienteERP__c + '-' + product.CodigoProdutoERP__c;

        PricebookEntry newPricebookEntry = new PricebookEntry();
        newPricebookEntry.Product2Id = product.Id;
        newPricebookEntry.Pricebook2Id = priceBook.Id;
        newPricebookEntry.UnitPrice = preTabela;
        newPricebookEntry.KeyBuyerGroup__c = key;
        insert newPricebookEntry;
        
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo" : 1,"msg_descricao":"Integracao realizada com sucesso","msg_retorno":[{"tp_tabpreco": [ { "emp_empresa": ' + account.Empresa__c + ', "uni_unidade": ' + account.Unidade__c + ', "cli_codigo": ' + account.CodigoClienteERP__c + ', "pre_tiporeg": 1, "pit_tabpre": ' + account.Ocorrencia__c + ', "pro_codigo": ' + product.CodigoProdutoERP__c + ', "pre_dtatual": "2001-01-01", "pre_tabela": ' + preTabela + ', "pre_vlrst": 1, "pre_vfcpst": 1 }]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        test.startTest();
        System.runAs(us) {
        	//B2bGerenciadorPricebook.sincronizar(new List<Account> {account});   
        }
        test.stopTest();
    }
 
    @IsTest
    static void sincronizarError01() {
        User us = [
            SELECT Id
            FROM User
            WHERE Alias = 'isTest'
        ];
        
        Account account = [
            SELECT Id, Name, CodigoClienteERP__c, Unidade__c
            FROM Account
            LIMIT 1
        ];

        account.Name = 'isError';
        update account;

        Boolean occurredError = false;

        test.startTest();
        System.runAs(us) {
            try {
                B2bGerenciadorPricebook.sincronizar(new List<Account> {account});   
            } catch (Exception ex) {
                occurredError = true;
            }
        }
        test.stopTest();

        System.assertEquals(true, occurredError);
    }
    
    @IsTest
    static void sincronizarAccountSuccess01() {        
        Integer msgCodigo = 1;

        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo" :' + msgCodigo + ',"msg_descricao":"Integracao realizada com sucesso","msg_retorno":[{"tp_tabpreco": [ { "emp_empresa": 1, "uni_unidade": 1, "cli_codigo": 10101, "pre_tiporeg": 27, "pit_tabpre": 1, "pro_codigo": 10101, "pre_dtatual": "2001-01-01", "pre_tabela": 1, "pre_vlrst": 1, "pre_vfcpst": 1 } ]} ]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        test.startTest();
		B2bGerenciadorPricebook.sincronizarAccount([SELECT Id, Name, CodigoClienteERP__c, Unidade__c FROM Account LIMIT 1]);
        test.stopTest();
    }

    @IsTest
    static void sincronizarAccountError01() {        
        Integer msgCodigo = 0;

        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo" :' + msgCodigo + ',"msg_descricao":"Ocorreu um erro na integracao","msg_retorno":[{"tp_tabpreco": [ { "emp_empresa": 1, "uni_unidade": 1, "cli_codigo": 10101, "pre_tiporeg": 27, "pit_tabpre": 1, "pro_codigo": 10101, "pre_dtatual": "2001-01-01", "pre_tabela": 1, "pre_vlrst": 1, "pre_vfcpst": 1 } ]} ]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        Boolean occurredError = false;

        test.startTest();
        try {
            B2bGerenciadorPricebook.sincronizarAccount([SELECT Id, Name, CodigoClienteERP__c, Unidade__c FROM Account LIMIT 1]);
        } catch (Exception ex) {
            occurredError = true;
        }
        test.stopTest();

        System.assertEquals(true, occurredError);
    }

}