@isTest
public with sharing class CommerceCieloIntegrationTest {
    @TestSetup
    static void makeData(){
        TestUtil util = new TestUtil();
        util.genStoreData();
    }

    public static void checkError(){
        List<LogErro__c> eLog = [SELECT Id, Nome__c, TextoErro__c FROM LogErro__c LIMIT 1];
        if(!eLog.isEmpty()){
            System.assertEquals(true, eLog.isEmpty(), 'Unexpected error log: ' +  eLog[0].Nome__c + ' - ' +  eLog[0].TextoErro__c);
        }
    }

    @IsTest
    static void testVerifyCardValid(){
        Test.setMock(HttpCalloutMock.class, new CieloIntegrationMock());

        Test.startTest();
            CommerceCieloIntegration.Card c = new CommerceCieloIntegration.Card();
            c.CardNumber = '5502095822650000';
            c.ExpirationDate = '01/2099';
            c.Holder = 'TEST';
            c.SecurityCode = '123';
            c.CardType = 'CreditCard';

            CommerceCieloIntegration verifyIntegration = new CommerceCieloIntegration('CieloVerifyCard');
            verifyIntegration.cardInput = c;
            verifyIntegration.createPayload();
            verifyIntegration.sendPayload();
            verifyIntegration.processFinalDMLs();
        Test.stopTest();
        checkError();
        
    }

    @IsTest
    static void testVerifyCardInvalid(){
        Test.setMock(HttpCalloutMock.class, new CieloIntegrationMock('39'));

        Test.startTest();
            CommerceCieloIntegration.Card c = new CommerceCieloIntegration.Card();
            c.CardNumber = '5502095822650000';
            c.ExpirationDate = '01/2099';
            c.Holder = 'TEST';
            c.SecurityCode = '123';
            c.CardType = 'CreditCard';

            CommerceCieloIntegration verifyIntegration = new CommerceCieloIntegration('CieloVerifyCard');
            verifyIntegration.cardInput = c;
            verifyIntegration.createPayload();
            verifyIntegration.sendPayload();
            verifyIntegration.processFinalDMLs();
        Test.stopTest();
        checkError();
        
    }

    @IsTest
    static void testCreateCreditCardPayment(){
        WebCart cart = [SELECT Id, AccountId, GrandTotalAmount, ValorCashback__c, InitialOrderReferenceNumber, PaymentMethodId, CurrencyIsoCode FROM WebCart LIMIT 1];
        if(cart.ValorCashback__c == null){
            cart.ValorCashback__c = 0;
        }
        Account account = [
            SELECT Id, Email__c, cleanCpfCnpj__c, Name, ShippingStreet, ShippingPostalCode, ShippingCity, ShippingState, ShippingCountry, ComplementoEntrega__c, NumeroEntrega__c, BairroEntrega__c, CNPJ__c, CodigoClienteERP__c, PaisEnderecoEntrega__c, RuaEntrega__c, CEPEntrega__c, CidadeEntrega__r.Name, CidadeEntrega__r.UF__c
            FROM Account
            WHERE Id = :cart.AccountId
        ];

        CartDeliveryGroup cdg = [SELECT Id, DeliverToCity, DeliverToCountry, DeliverToState, DeliverToStreet, DeliverToPostalCode FROM CartDeliveryGroup WHERE CartId = :cart.Id LIMIT 1];

        cart.MetodoPagamento__c = 'cc';

        Test.setMock(HttpCalloutMock.class, new CieloIntegrationMock());

        Test.startTest();
            CommerceCieloIntegration.Card c = new CommerceCieloIntegration.Card();
            c.CardNumber = '5502095822650000';
            c.ExpirationDate = '01/2099';
            c.Holder = 'TEST';
            c.SecurityCode = '123';
            c.CardType = 'CreditCard';

            CommerceCieloIntegration integration = new CommerceCieloIntegration('CieloCreatePayment');
            integration.cardInput = c;
            integration.installments = 1;
            integration.cart = cart;
            integration.acc = account;
            integration.cdg = cdg;
            integration.process();
        Test.stopTest();
        checkError();
    }

    @IsTest
    static void testCreatePixPayment(){
        WebCart cart = [SELECT Id, AccountId, GrandTotalAmount, ValorCashback__c, InitialOrderReferenceNumber, PaymentMethodId, CurrencyIsoCode FROM WebCart LIMIT 1];
        if(cart.ValorCashback__c == null){
            cart.ValorCashback__c = 0;
        }
        Account account = [
            SELECT Id, Email__c, cleanCpfCnpj__c, Name, ShippingStreet, ShippingPostalCode, ShippingCity, ShippingState, ShippingCountry, ComplementoEntrega__c, NumeroEntrega__c, BairroEntrega__c, CNPJ__c, CodigoClienteERP__c, PaisEnderecoEntrega__c, RuaEntrega__c, CEPEntrega__c, CidadeEntrega__r.Name, CidadeEntrega__r.UF__c
            FROM Account
            WHERE Id = :cart.AccountId
        ];

        CartDeliveryGroup cdg = [SELECT Id, DeliverToCity, DeliverToCountry, DeliverToState, DeliverToStreet, DeliverToPostalCode FROM CartDeliveryGroup WHERE CartId = :cart.Id LIMIT 1];

        cart.MetodoPagamento__c = 'pix';

        Test.setMock(HttpCalloutMock.class, new CieloIntegrationMock('pix_success'));

        Test.startTest();
            CommerceCieloIntegration integration = new CommerceCieloIntegration('CieloCreatePayment');
            integration.cart = cart;
            integration.acc = account;
            integration.cdg = cdg;
            integration.process();
        Test.stopTest();
        checkError();
    }

    @IsTest
    static void testRefund(){

        Test.setMock(HttpCalloutMock.class, new CieloIntegrationMock());

        Test.startTest();
            CommerceCieloIntegration integration = new CommerceCieloIntegration('CieloRefund', [SELECT GatewayRefNumber FROM Payment LIMIT 1][0].GatewayRefNumber);
            integration.process();
        Test.stopTest();
        checkError();
    }

    @IsTest
    static void testGetPayment(){

        Test.setMock(HttpCalloutMock.class, new CieloIntegrationMock());

        Test.startTest();
            CommerceCieloIntegration integration = new CommerceCieloIntegration('CieloGetPayment', [SELECT GatewayRefNumber FROM Payment LIMIT 1][0].GatewayRefNumber);
            integration.process();
        Test.stopTest();
        checkError();
    }

    @IsTest
    static void testInvocable(){

        Test.setMock(HttpCalloutMock.class, new CieloIntegrationMock());

        Map<String,Object> data = new Map<String,Object>{
                'cardParameter' => new Map<String,Object>{
                    'numeroCartao' => '5502095822650000',
                    'mesCartao' => '1',
                    'anoCartao' => '2099',
                    'nomeCartao' => 'TESTE',
                    'cvvCartao' => '123',
                    'parcelas' => '1x'
                }
            };

        WebCart cart = [SELECT Id, AccountId, GrandTotalAmount, ValorCashback__c, InitialOrderReferenceNumber, PaymentMethodId, CurrencyIsoCode FROM WebCart LIMIT 1];
        if(cart.ValorCashback__c == null){
            cart.ValorCashback__c = 0;
        }
        Account account = [
            SELECT Id, Email__c, cleanCpfCnpj__c, Name, ShippingStreet, ShippingPostalCode, ShippingCity, ShippingState, ShippingCountry, ComplementoEntrega__c, NumeroEntrega__c, BairroEntrega__c, CNPJ__c, CodigoClienteERP__c, PaisEnderecoEntrega__c, RuaEntrega__c, CEPEntrega__c, CidadeEntrega__r.Name, CidadeEntrega__r.UF__c
            FROM Account
            WHERE Id = :cart.AccountId
        ];

        CartDeliveryGroup cdg = [SELECT Id, DeliverToCity, DeliverToCountry, DeliverToState, DeliverToStreet, DeliverToPostalCode FROM CartDeliveryGroup WHERE CartId = :cart.Id LIMIT 1];

        
        Test.startTest();
            CommerceCieloIntegration.Card c = new CommerceCieloIntegration.Card();
            c.CardNumber = '5502095822650000';
            c.ExpirationDate = '01/2099';
            c.Holder = 'TEST';
            c.SecurityCode = '123';
            c.CardType = 'CreditCard';

            CommerceCieloIntegration.FlowInput i1= new CommerceCieloIntegration.FlowInput();
            i1.operation = 'CieloVerifyCard';
            i1.cardInput = c;

            cart.MetodoPagamento__c = 'cc';
            CommerceCieloIntegration.FlowInput i2= new CommerceCieloIntegration.FlowInput();
            i2.operation = 'CieloCreatePayment';
            i2.acc = account;
            i2.cart = cart;
            i2.cdg = cdg;
            i2.cardInput = c;
            i2.installments = 1;

            CommerceCieloIntegration.invocableExecuteIntegration(new List<CommerceCieloIntegration.FlowInput>{i1,i2});
        Test.stopTest();
        checkError();
    }

    @IsTest
    static void testInvocableAsync(){

        Test.setMock(HttpCalloutMock.class, new CieloIntegrationMock());

        WebCart cart = [SELECT Id, AccountId, GrandTotalAmount, ValorCashback__c, InitialOrderReferenceNumber, PaymentMethodId, CurrencyIsoCode FROM WebCart LIMIT 1];
        if(cart.ValorCashback__c == null){
            cart.ValorCashback__c = 0;
        }
        Account account = [
            SELECT Id, Email__c, cleanCpfCnpj__c, Name, ShippingStreet, ShippingPostalCode, ShippingCity, ShippingState, ShippingCountry, ComplementoEntrega__c, NumeroEntrega__c, BairroEntrega__c, CNPJ__c, CodigoClienteERP__c, PaisEnderecoEntrega__c, RuaEntrega__c, CEPEntrega__c, CidadeEntrega__r.Name, CidadeEntrega__r.UF__c
            FROM Account
            WHERE Id = :cart.AccountId
        ];

        CartDeliveryGroup cdg = [SELECT Id, DeliverToCity, DeliverToCountry, DeliverToState, DeliverToStreet, DeliverToPostalCode FROM CartDeliveryGroup WHERE CartId = :cart.Id LIMIT 1];

        
        Test.startTest();
            CommerceCieloIntegration.Card c = new CommerceCieloIntegration.Card();
            c.CardNumber = '5502095822650000';
            c.ExpirationDate = '01/2099';
            c.Holder = 'TEST';
            c.SecurityCode = '123';
            c.CardType = 'CreditCard';

            cart.MetodoPagamento__c = 'cc';
            CommerceCieloIntegration.FlowInput i2= new CommerceCieloIntegration.FlowInput();
            i2.operation = 'CieloCreatePayment';
            i2.acc = account;
            i2.cart = cart;
            i2.cdg = cdg;
            i2.async = true;
            i2.cardInput = c;
            i2.installments = 1;

            CommerceCieloIntegration.invocableExecuteIntegration(new List<CommerceCieloIntegration.FlowInput>{i2});
        Test.stopTest();
        checkError();
    }
}