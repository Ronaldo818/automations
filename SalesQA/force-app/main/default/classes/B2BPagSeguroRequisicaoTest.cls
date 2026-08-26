@IsTest
public class B2BPagSeguroRequisicaoTest {
	
    @TestSetup
    static void makeData() {
    	TestUtil util = new TestUtil('b2b');
        util.genAllData();
    }
    
    @Istest
    static void parseSuccessPix() {
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{ "id": "ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "reference_id": "ex-00001", "created_at": "2021-01-19T09:30:12.197-03:00", "customer": { "name": "José da Silva", "email": "jose@email.com", "tax_id": "12345678901", "phones": [ { "country": "55", "area": "11", "number": "999999999", "type": "MOBILE" } ] }, "items": [ { "reference_id": "referencia do item", "name": "nome do item", "quantity": "1", "unit_amount": "100" } ], "amount": { "currency": "BRL", "additional": 100, "discount": 100 }, "shipping": { "address": { "street": "Avenida Brigadeiro Faria Lima", "number": "1384", "complement": "apto 12", "locality": "Pinheiros", "city": "São Paulo", "region_code": "SP", "country": "BRA", "postal_code": "01452002" } }, "qr_codes": [ { "id": "QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA", "expiration_date": "2021-08-29T20:15:59-03:00", "amount": { "value": 100 }, "text": "00020101021226830014br.gov.bcb.pix2561api.pagseguro.com/pix/v2/9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA27600016BR.COM.PAGSEGURO01369E13BFE1-35C3-4DFD-B499-9B110AC0E1BA52045697530398654041.005802BR5925Leticia Oliveira Porto La6007Barueri62070503***6304658F", "links": [ { "rel": "QRCODE.PNG", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/png", "media": "image/png", "type": "GET" }, { "rel": "QRCODE.BASE64", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/base64", "media": "text/plain", "type": "GET" } ] } ], "charges": [], "links": [ { "rel": "SELF", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "media": "application/json", "type": "GET" }, { "rel": "PAY", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31/pay", "media": "application/json", "type": "POST" } ] }');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        B2BPagSeguroRetorno retorno = new B2BPagSeguroRetorno();
        B2BLWRCheckoutControler.PaymentParameter parameter = new B2BLWRCheckoutControler.PaymentParameter();
		parameter.cartId = [SELECT Id FROM WebCart LIMIT 1].Id;
		parameter.parcelas = '1';
		parameter.formaPagamento = 'pix';
		parameter.valorCashback = '0';
		
        WebCart cart = [
            SELECT Id, GrandTotalAmount, ValorCashback__c, AccountId
            FROM WebCart
            LIMIT 1
        ];
        
        Account account = [
            SELECT Id, Name, ShippingStreet, ShippingPostalCode, ShippingCity, ShippingState, ShippingCountry, ComplementoEntrega__c, NumeroEntrega__c, BairroEntrega__c, CNPJ__c, CodigoClienteERP__c, PaisEnderecoEntrega__c, RuaEntrega__c, CEPEntrega__c, CidadeEntrega__r.Name, CidadeEntrega__r.UF__c
            FROM Account
            WHERE Id = :cart.AccountId
        ];

        B2BPagSeguro.ApexInput input = new B2BPagSeguro.ApexInput(parameter);
        input.cart = cart;
        input.account = account;

        Test.startTest();
            System.runAs([SELECT Id FROM User WHERE username = 'test.b2b@test.com'][0]){
                List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
                retorno = B2BPagSeguro.executar(input);
                List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
                System.assertEquals(0, lstLogBefore.size());
                System.assertEquals(1, lstLogAfter.size());
            }

        Test.stopTest();
    }
    
    @Istest
    static void parseSuccessCreditCard() {
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{ "id": "ORDE_1E38BD2E-F2CC-4D9C-9727-787CDFBCA7CE", "reference_id": "ex-00001", "created_at": "2023-02-08T15:15:11.408-03:00", "customer": { "name": "Jose da Silva", "email": "email@test.com", "tax_id": "12345678909", "phones": [ { "type": "MOBILE", "country": "55", "area": "11", "number": "999999999" } ] }, "items": [ { "reference_id": "referencia do item", "name": "nome do item", "quantity": 1, "unit_amount": 500 } ], "shipping": { "address": { "street": "Avenida Brigadeiro Faria Lima", "number": "1384", "complement": "apto 12", "locality": "Pinheiros", "city": "São Paulo", "region_code": "SP", "country": "BRA", "postal_code": "01452002" } }, "charges": [ { "id": "CHAR_67FC568B-00D8-431D-B2E7-755E3E6C66A0", "reference_id": "referencia da cobranca", "status": "PAID", "created_at": "2023-02-08T15:15:11.881-03:00", "paid_at": "2023-02-08T15:15:12.000-03:00", "description": "descricao da cobranca", "amount": { "value": 500, "currency": "BRL", "summary": { "total": 500, "paid": 500, "refunded": 0 } }, "payment_response": { "code": "20000", "message": "SUCESSO", "reference": "032416400102" }, "payment_method": { "type": "CREDIT_CARD", "installments": 1, "capture": true, "card": { "brand": "visa", "first_digits": "411111", "last_digits": "1111", "exp_month": "12", "exp_year": "2026", "holder": { "name": "Joãozinho da Silva", "tax_id": "65544332211" }, "store": false }, "soft_descriptor": "IntegracaoPagsegu" }, "links": [ { "rel": "SELF", "href": "https://sandbox.api.pagseguro.com/charges/CHAR_67FC568B-00D8-431D-B2E7-755E3E6C66A0", "media": "application/json", "type": "GET" }, { "rel": "CHARGE.CANCEL", "href": "https://sandbox.api.pagseguro.com/charges/CHAR_67FC568B-00D8-431D-B2E7-755E3E6C66A0/cancel", "media": "application/json", "type": "POST" } ] } ], "notification_urls": [ "https://meusite.com/notificacoes" ], "links": [ { "rel": "SELF", "href": "https://sandbox.api.pagseguro.com/orders/ORDE_1E38BD2E-F2CC-4D9C-9727-787CDFBCA7CE", "media": "application/json", "type": "GET" }, { "rel": "PAY", "href": "https://sandbox.api.pagseguro.com/orders/ORDE_1E38BD2E-F2CC-4D9C-9727-787CDFBCA7CE/pay", "media": "application/json", "type": "POST" } ] }');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        B2BPagSeguroRetorno retorno = new B2BPagSeguroRetorno();
        B2BLWRCheckoutControler.PaymentParameter parameter = new B2BLWRCheckoutControler.PaymentParameter();
		parameter.cartId = [SELECT Id FROM WebCart LIMIT 1].Id;
		parameter.parcelas = '1';
		parameter.cvvCartao = '000';
		parameter.anoCartao = '99';
		parameter.mesCartao = '01';
		parameter.numeroCartao = '0000000000';
		parameter.nomeCartao = 'isTest';
		parameter.formaPagamento = 'cc';
		parameter.valorCashback = '0';
		
        WebCart cart = [
            SELECT Id, GrandTotalAmount, ValorCashback__c, AccountId
            FROM WebCart
            LIMIT 1
        ];
        
        Account account = [
            SELECT Id, Name, ShippingStreet, ShippingPostalCode, ShippingCity, ShippingState, ShippingCountry, ComplementoEntrega__c, NumeroEntrega__c, BairroEntrega__c, CNPJ__c, CodigoClienteERP__c, PaisEnderecoEntrega__c, RuaEntrega__c, CEPEntrega__c, CidadeEntrega__r.Name, CidadeEntrega__r.UF__c
            FROM Account
            WHERE Id = :cart.AccountId
        ];

        B2BPagSeguro.ApexInput input = new B2BPagSeguro.ApexInput(parameter);
        input.cart = cart;
        input.account = account;

        Test.startTest();
        System.runAs([SELECT Id FROM User WHERE username = 'test.b2b@test.com'][0]){
            List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
            retorno = B2BPagSeguro.executar(input);
            List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
             System.assertEquals(0, lstLogBefore.size());
            System.assertEquals(1, lstLogAfter.size());
        }
        Test.stopTest();
    }
    
    @Istest
    static void parseSuccessBoleto() {
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{ "id": "ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "reference_id": "ex-00001", "created_at": "2021-01-19T09:30:12.197-03:00", "customer": { "name": "José da Silva", "email": "jose@email.com", "tax_id": "12345678901", "phones": [ { "country": "55", "area": "11", "number": "999999999", "type": "MOBILE" } ] }, "items": [ { "reference_id": "referencia do item", "name": "nome do item", "quantity": "1", "unit_amount": "100" } ], "amount": { "currency": "BRL", "additional": 100, "discount": 100 }, "shipping": { "address": { "street": "Avenida Brigadeiro Faria Lima", "number": "1384", "complement": "apto 12", "locality": "Pinheiros", "city": "São Paulo", "region_code": "SP", "country": "BRA", "postal_code": "01452002" } }, "qr_codes": [ { "id": "QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA", "expiration_date": "2021-08-29T20:15:59-03:00", "amount": { "value": 100 }, "text": "00020101021226830014br.gov.bcb.pix2561api.pagseguro.com/pix/v2/9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA27600016BR.COM.PAGSEGURO01369E13BFE1-35C3-4DFD-B499-9B110AC0E1BA52045697530398654041.005802BR5925Leticia Oliveira Porto La6007Barueri62070503***6304658F", "links": [ { "rel": "QRCODE.PNG", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/png", "media": "image/png", "type": "GET" }, { "rel": "QRCODE.BASE64", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/base64", "media": "text/plain", "type": "GET" } ] } ], "charges": [], "links": [ { "rel": "SELF", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "media": "application/json", "type": "GET" }, { "rel": "PAY", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31/pay", "media": "application/json", "type": "POST" } ] }');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        B2BPagSeguroRetorno retorno = new B2BPagSeguroRetorno();
        B2BLWRCheckoutControler.PaymentParameter parameter = new B2BLWRCheckoutControler.PaymentParameter();
		parameter.cartId = [SELECT Id FROM WebCart LIMIT 1].Id;
		parameter.parcelas = '1';
		parameter.formaPagamento = 'boleto';
		parameter.valorCashback = '0';
		
        WebCart cart = [
            SELECT Id, GrandTotalAmount, ValorCashback__c, AccountId
            FROM WebCart
            LIMIT 1
        ];
        
        Account account = [
            SELECT Id, Name, ShippingStreet, ShippingPostalCode, ShippingCity, ShippingState, ShippingCountry, ComplementoEntrega__c, NumeroEntrega__c, BairroEntrega__c, CNPJ__c, CodigoClienteERP__c, PaisEnderecoEntrega__c, RuaEntrega__c, CEPEntrega__c, CidadeEntrega__r.Name, CidadeEntrega__r.UF__c
            FROM Account
            WHERE Id = :cart.AccountId
        ];

        B2BPagSeguro.ApexInput input = new B2BPagSeguro.ApexInput(parameter);
        input.cart = cart;
        input.account = account;

        Test.startTest();
        System.runAs([SELECT Id FROM User WHERE username = 'test.b2b@test.com'][0]){
            List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
            retorno = B2BPagSeguro.executar(input);
            List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
            System.assertEquals(0, lstLogBefore.size());
            System.assertEquals(1, lstLogAfter.size());
        }
        Test.stopTest();
    }
    
    @Istest
    static void parseError01() {
        Account acc = [
            SELECT Id, Name
            FROM Account
            LIMIT 1
        ];
        
        acc.Name = 'error';
        update acc;
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{ "id": "ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "reference_id": "ex-00001", "created_at": "2021-01-19T09:30:12.197-03:00", "customer": { "name": "José da Silva", "email": "jose@email.com", "tax_id": "12345678901", "phones": [ { "country": "55", "area": "11", "number": "999999999", "type": "MOBILE" } ] }, "items": [ { "reference_id": "referencia do item", "name": "nome do item", "quantity": "1", "unit_amount": "100" } ], "amount": { "currency": "BRL", "additional": 100, "discount": 100 }, "shipping": { "address": { "street": "Avenida Brigadeiro Faria Lima", "number": "1384", "complement": "apto 12", "locality": "Pinheiros", "city": "São Paulo", "region_code": "SP", "country": "BRA", "postal_code": "01452002" } }, "qr_codes": [ { "id": "QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA", "expiration_date": "2021-08-29T20:15:59-03:00", "amount": { "value": 100 }, "text": "00020101021226830014br.gov.bcb.pix2561api.pagseguro.com/pix/v2/9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA27600016BR.COM.PAGSEGURO01369E13BFE1-35C3-4DFD-B499-9B110AC0E1BA52045697530398654041.005802BR5925Leticia Oliveira Porto La6007Barueri62070503***6304658F", "links": [ { "rel": "QRCODE.PNG", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/png", "media": "image/png", "type": "GET" }, { "rel": "QRCODE.BASE64", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/base64", "media": "text/plain", "type": "GET" } ] } ], "charges": [], "links": [ { "rel": "SELF", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "media": "application/json", "type": "GET" }, { "rel": "PAY", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31/pay", "media": "application/json", "type": "POST" } ] }');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        B2BPagSeguroRetorno retorno = new B2BPagSeguroRetorno();
        B2BLWRCheckoutControler.PaymentParameter parameter = new B2BLWRCheckoutControler.PaymentParameter();
		parameter.cartId = [SELECT Id FROM WebCart LIMIT 1].Id;
		parameter.parcelas = '1';
		parameter.formaPagamento = 'boleto';
		parameter.valorCashback = '0';
		
        WebCart cart = [
            SELECT Id, GrandTotalAmount, ValorCashback__c, AccountId
            FROM WebCart
            LIMIT 1
        ];
        
        Account account = [
            SELECT Id, Name, ShippingStreet, ShippingPostalCode, ShippingCity, ShippingState, ShippingCountry, ComplementoEntrega__c, NumeroEntrega__c, BairroEntrega__c, CNPJ__c, CodigoClienteERP__c, PaisEnderecoEntrega__c, RuaEntrega__c, CEPEntrega__c, CidadeEntrega__r.Name, CidadeEntrega__r.UF__c
            FROM Account
            WHERE Id = :cart.AccountId
        ];

        B2BPagSeguro.ApexInput input = new B2BPagSeguro.ApexInput(parameter);
        input.cart = cart;
        input.account = account;

        Test.startTest();
        System.runAs([SELECT Id FROM User WHERE username = 'test.b2b@test.com'][0]){
            Boolean occurredError = false;
            try {
                retorno = B2BPagSeguro.executar(input);
            } catch (Exception ex) {
                occurredError = true;
            }

            System.assertEquals(true, occurredError);
        }
        Test.stopTest();
    }
    
}