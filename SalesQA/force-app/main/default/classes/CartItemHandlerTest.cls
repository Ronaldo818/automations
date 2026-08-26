@IsTest
public class CartItemHandlerTest {
	
    @TestSetup
    static void makeData() {
        
    	Product2 newProduct = new Product2(
            Name = 'isTest',
        	PesoUnitario__c = 1,
            ShippingWeight = 1,
        	StockKeepingUnit = '1',
        	CodigoProdutoERP__c = '1',
        	UnidadeMedida__c = 'KG'
        );
		insert newProduct;
        
        Pricebook2 newPriceBook = new Pricebook2(
            Name = 'isTest'
        );
		insert newPriceBook;
        
        PricebookEntry standardPrice = new PricebookEntry(
            Pricebook2Id = Test.getStandardPricebookId(),
            Product2Id = newProduct.Id,
            UnitPrice = 0,
            IsActive = true
        );
        insert standardPrice;
        
        PricebookEntry newPricebokEntry = new PricebookEntry(
			UnitPrice = 1,
    		Pricebook2Id = newPricebook.Id,
			Product2Id = newProduct.Id
        );
		insert newPricebokEntry;
        
        UnidadeMedida__c newUnidadeMedida = new UnidadeMedida__c(
        	Name 					= 'CAIXA',
			Descricao__c 			= 'CAIXA',
            CodigoUnidadeMedida__c	= 'CX',
            TipoMedida__c 			= true
        );
        insert newUnidadeMedida;

        UnidadeProduto__c newUnidadeProduto = new UnidadeProduto__c(
            Key__c = '1-1',
        	Unidade__c = 1,
            UnidadePedido__c = newUnidadeMedida.Id,
            Product__c = newProduct.Id
        );
        insert newUnidadeProduto;
        
        Account newAccount = new Account(
        	Name = 'isTest',
        	Unidade__c = 1,
        	Empresa__c = 1,
        	CodigoClienteERP__c = '1',
        	Ocorrencia__c = 1,
        	Lista__c = 1
        );
        insert newAccount;
        
        WebStore newStore = new WebStore(
            Name = 'isTest'
        );
		insert newStore;
        
        WebCart newCart = new WebCart(
        	AccountId = newAccount.Id,
        	Name = 'isTest',
        	WebStoreId = newStore.Id,
        	Type = 'Cart',
        	Status = 'Active'
        );
        insert newCart;
        
        CartDeliveryGroup newDeliveryGroup = new CartDeliveryGroup(
        	Name = 'isTest',
        	CartId = newCart.Id
        );
		insert newDeliveryGroup;	
       	
        CartItem newCartItem01 = new CartItem(
        	Name = 'oldCartItem',
        	CartId = newCart.Id,
        	CartDeliveryGroupId = newDeliveryGroup.Id,
        	Product2Id = newProduct.Id,
        	Quantity = 1,
        	Type = 'Product',
        	SalesPrice = 2,
        	Sku = '1',
        	ListPrice = 3,
        	TotalPrice = 4
        );
        insert newCartItem01;
        
        CartItem newCartItem02 = new CartItem(
        	Name = 'newCartItem',
        	CartId = newCart.Id,
        	CartDeliveryGroupId = newDeliveryGroup.Id,
        	Product2Id = newProduct.Id,
        	Quantity = 1,
        	Type = 'Product',
        	SalesPrice = 1,
        	Sku = '1',
        	ListPrice = 3,
        	TotalPrice = 1
        );
        insert newCartItem02;
    }
    
    @IsTest
    static void verificarMultiploQuantidadeSuccess01() {
        List<CartItem> oldList = new List<CartItem>();
   		List<CartItem> newList = new List<CartItem>();
    	Map<Id,CartItem> oldMap = new Map<Id,CartItem>();
    	Map<Id,CartItem> newMap = new Map<Id,CartItem>();
        
        for(CartItem item : [SELECT Id, Name, CartId, Product2Id, Quantity FROM CartItem]) {
       		if(item.Name.equalsIgnoreCase('oldcartitem')) {
            	oldList.add(item);
                oldMap.put(item.Id, item);
            }
            else if(item.Name.equalsIgnoreCase('newcartitem')) {
            	newList.add(item);
                newMap.put(item.Id, item);    
            }
        }
        
        test.startTest();
		CartItemHandler newCartItemHandler = new CartItemHandler(oldList, newList, oldMap, newMap);
		newCartItemHandler.verificarMultiploQuantidade();
        test.stopTest();
    }
    /*
    @IsTest
    static void verificarMultiploQuantidadeSuccess02() {
        List<sObject> lstUpdateSObject = new List<sObject>();
        
        Product2 produ = [
        	SELECT Id, PesoUnitario__c
            FROM Product2
            LIMIT 1
        ];
        produ.PesoUnitario__c = 5;
        
        UnidadeMedida__c unidadeMedida = [
        	SELECT Id, TipoMedida__c
            FROM UnidadeMedida__c
            LIMIT 1
        ];
        unidadeMedida.TipoMedida__c = false;
        
        lstUpdateSObject.add(produ);
        lstUpdateSObject.add(unidadeMedida);
        update lstUpdateSObject;
        
        List<CartItem> oldList = new List<CartItem>();
   		List<CartItem> newList = new List<CartItem>();
    	Map<Id,CartItem> oldMap = new Map<Id,CartItem>();
    	Map<Id,CartItem> newMap = new Map<Id,CartItem>();
        
        for(CartItem item : [SELECT Id, Name, CartId, Product2Id, Quantity FROM CartItem]) {
       		if(item.Name.equalsIgnoreCase('oldcartitem')) {
            	oldList.add(item);
                oldMap.put(item.Id, item);
            }
            else if(item.Name.equalsIgnoreCase('newcartitem')) {
            	newList.add(item);
                newMap.put(item.Id, item);    
            }
        }
        
        test.startTest();
        Boolean occurredError = false;
        
        try {
			CartItemHandler newCartItemHandler = new CartItemHandler(oldList, newList, oldMap, newMap);
			newCartItemHandler.verificarMultiploQuantidade();   
        }
        catch(Exception ex) {
        	occurredError = true;	    
        }
        
        System.assertEquals(true, occurredError);
        test.stopTest();
    }
	*/
}