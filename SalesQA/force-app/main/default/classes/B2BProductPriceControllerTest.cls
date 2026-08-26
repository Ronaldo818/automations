@IsTest
public class B2BProductPriceControllerTest {
	
    @TestSetup
    static void makeData() {
        Profile profile = [SELECT Id FROM Profile WHERE Name='Commerce User' LIMIT 1];
        
        Account newAccount = new Account(
            Name = 'TestAccount',
            empresa__c = 1,
            CodigoClienteERP__c = '234',
            Unidade__c = 1,
        	Lista__c = 27,
            Ocorrencia__c = 1);
        insert newAccount;

        Contact newContact = new Contact(
            LastName = 'isTest',
            AccountId = newAccount.id,
            Email = 'teste.teste@teste.com'
        );
        insert newContact;
        
        User newUser = new User(Alias='istest', username='istest@istest.com', EmailEncodingKey='UTF-8', LanguageLocaleKey='pt_BR',
                                LocaleSidKey='pt_BR', TimeZoneSidKey='America/Sao_Paulo', ProfileId = profile.Id, lastname='isTest',
                                email='teste@teste.com', phone='+5511999999999', ContactId = newContact.Id);
        insert newUser;
        
        WebStore newWebStore = new WebStore(Name='isTest');
        insert newWebStore;
        
        WebCart newCart = new WebCart(
            Name= 'Cart',
            WebStoreId= newWebStore.Id,
            AccountId= newAccount.Id,
            OwnerId = newUser.Id,
        	Status = 'Checkout');
        insert newCart;
        
        Product2 newProduct = new Product2 (
            Name = 'test',
            ProductCode = '1101100100',
            IsActive = true,
            CodigoProdutoERP__c = '1101100100',
            PesoUnitario__c = 5,
            StockKeepingUnit = '1101100100',
            UnidadeMedida__c = 'Kg');
        insert newProduct;
        
        UnidadeMedida__c newUnidadeMedida = new UnidadeMedida__c(
            TipoMedida__c = false
        );
        insert newUnidadeMedida;
        
        UnidadeProduto__c newUnidadeProduto = new UnidadeProduto__c(
            Key__c = '1-1',
        	Unidade__c = 1,
        	Product__c = newProduct.Id,
        	UnidadePedido__c = newUnidadeMedida.Id
        );
        insert newUnidadeProduto;
        
        CartDeliveryGroup newDeliveryGroup = new CartDeliveryGroup(
            CartId = newCart.Id,
            Name='isTest');
        insert newDeliveryGroup;
        
        CartItem newCartItem = new CartItem(
            CartId = newCart.Id,
            Sku = '1101100100',
            SalesPrice = 10.00,
            Quantity = 5,
            TotalPrice = 5,
            Type = 'Product',
            Name = 'isTest',
            CartDeliveryGroupId = newDeliveryGroup.Id,
            Product2Id = newProduct.Id,
            TotalAdjustmentAmount = 1.0);
        insert newCartItem;
    }
    
    @IsTest
    static void getProductPriceSuccess() {
        User thisUser = [SELECT Id FROM User WHERE Alias= 'istest' LIMIT 1];
        String productId = [SELECT Id FROM Product2 LIMIT 1].Id;
        
        Test.startTest();
       	System.runAs(thisUser) {
            String retorno = B2BProductPriceController.getProductPrice(productId);
            
            System.assertEquals('20', retorno);
        }
        Test.stopTest();
    }
    
    @IsTest
    static void getProductPriceError() {
        User thisUser = [SELECT Id FROM User WHERE Alias= 'istest' LIMIT 1];
        Product2 product = [SELECT Id, Name FROM Product2 LIMIT 1];
        product.Name = 'isTest';
        update product;
        
        Test.startTest();
        Boolean error = false;
       	System.runAs(thisUser) {
            try {
            	String retorno = B2BProductPriceController.getProductPrice(product.Id);
            }	catch(Exception ex) {
                error = true;
            }
            
            System.assertEquals(true, error);
        }
        Test.stopTest();
    }
}