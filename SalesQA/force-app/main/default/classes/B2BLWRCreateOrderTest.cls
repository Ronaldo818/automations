@IsTest
global with sharing class B2BLWRCreateOrderTest {

    @TestSetup 
    static void makeData() {
        String profileId = [
            SELECT Id
            FROM Profile
            WHERE Name = 'Customer Community Plus User'
            LIMIT 1
        ].Id;
        
        Profile perfil = [
            SELECT Id
            FROM Profile
            WHERE Name = 'Customer Community Plus User'
            LIMIT 1
        ];
   
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
	
		String uniqueUserName = 'istest.istest' + DateTime.now().getTime() + '@istest.avivar.com';
        User newUser = new User(
            Alias 				= 'isTest',
            username 			= uniqueUserName,
            EmailEncodingKey 	= 'UTF-8',
            LanguageLocaleKey	= 'pt_BR',
			LocaleSidKey 		= 'pt_BR',
            TimeZoneSidKey 		= 'America/Sao_Paulo',
            ProfileId 			= profileId,
            lastname 			= 'isTest',
            email 				= 'istest.istest@istest.avivar.com',
            phone 				= '+5511979797979',
            ContactId 			= newContact.Id
        );
        insert newUser;
		
        WebStore newStore = new WebStore(
             Name = 'isTest'
         );
		insert newStore;
        
        WebCart newCart = new WebCart(
        	AccountId 	= newAccount.Id,
        	Name 		= 'isTest',
        	WebStoreId	= newStore.Id,
        	Type 		= 'Cart',
        	Status 		= 'Active'
        );
        insert newCart;
        
        CartCheckoutSession newCheckoutSession = new CartCheckoutSession(
            WebCartId	= newCart.Id,
            Name 		= newCart.Id,
            State 		= 'Start'
        );
        insert newCheckoutSession;
        
        CartDeliveryGroup newDeliveryGroup = new CartDeliveryGroup(
        	Name 	= 'isTest',
        	CartId	= newCart.Id
        );
		insert newDeliveryGroup;
        
		OrderDeliveryMethod newDeliveryMethod = new OrderDeliveryMethod(
            Name 		= 'Entrega Padrão',
            IsActive	= true
        );
        insert newDeliveryMethod;
        
        Product2 newProduct = new Product2 (
            Name = 'isTest',
            ProductCode = '1010101',
            Family = '035',
            IsActive = true,
            Description='isTest',
            CodigoProdutoERP__c = '1010101',
            PesoUnitario__c = 1,
            UnidadeMedida__c = 'Kg',
            StockKeepingUnit='1010101'
        );
        insert newProduct;
        
        /*
        UnidadeMedida__c newUnidadeMedida = new UnidadeMedida__c(
        	Name 					= 'KG',
			Descricao__c 			= 'KG',
            CodigoUnidadeMedida__c	= 'KG',
            TipoMedida__c 			= true
        );
        insert newUnidadeMedida;
		*/
        
        UnidadeProduto__c newUnidadeProduto = new UnidadeProduto__c(
            Key__c = '1-1',
        	Unidade__c = 1,
            //UnidadePedido__c = newUnidadeMedida.Id,
            Product__c = newProduct.Id
        );
        insert newUnidadeProduto;
        
        CartItem newCartItem = new CartItem(
            CartId = newCart.Id,
            Sku = '10101',
            SalesPrice = 10,
            Quantity = 1,
            TotalPrice = 5,
            Type = 'Product',
            Name='isTest',
            TotalAdjustmentAmount = 1,
            CartDeliveryGroupId = newDeliveryGroup.Id,
            Product2Id = newProduct.Id
        );
        insert newCartItem;
    }

    @IsTest
    public static void B2BLWRCreateOrderSuccess01() {
        WebCart cart = [
            SELECT Id
            FROM WebCart
            LIMIT 1
        ];
        
        WebStore store = [
            SELECT Id
            From WebStore
            LIMIT 1
        ];
        
        List<CartCheckoutSession> lstCartSession = [
            SELECT Id
            FROM CartCheckoutSession
            WHERE WebCartId = :cart.Id
            AND IsDeleted = false
            AND State = 'Start'
        ];

        Test.startTest();            
        B2BLWRCreateOrder b2b = new B2BLWRCreateOrder(lstCartSession[0].Id, cart);
        b2b.cartToOrder();
        Test.stopTest();
    }
    
}