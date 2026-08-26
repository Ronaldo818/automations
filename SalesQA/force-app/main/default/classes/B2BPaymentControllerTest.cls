@isTest
public class B2BPaymentControllerTest {	
    @TestSetup
    static void makeData() {
        TestUtil util = new TestUtil('b2b');
        util.genAllData();
    }

    public static void checkError(){
        List<LogErro__c> eLog = [SELECT Id, Nome__c, TextoErro__c FROM LogErro__c LIMIT 1];
        if(!eLog.isEmpty()){
            System.assertEquals(true, eLog.isEmpty(), 'Unexpected error log: ' +  eLog[0].Nome__c + ' - ' +  eLog[0].TextoErro__c);
        }
    }
    
    @IsTest
    static void getCashbackSuccess01() {
        test.startTest();
        	B2BPaymentController.getCashback();   
        test.stopTest();
    }

    @IsTest
    static void setCashbackSuccess01() {
        test.startTest();
        	B2BPaymentController.setCashBackValue(0);   
        test.stopTest();
    }
    
    @IsTest
    static void setCashbackError01() {
        test.startTest();
        	B2BPaymentController.setCashBackValue(1);   
        test.stopTest();
    }

    @IsTest
    static void paySuccesPix() {
        /*
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{ "id": "ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "reference_id": "ex-00001", "created_at": "2021-01-19T09:30:12.197-03:00", "customer": { "name": "José da Silva", "email": "jose@email.com", "tax_id": "12345678901", "phones": [ { "country": "55", "area": "11", "number": "999999999", "type": "MOBILE" } ] }, "items": [ { "reference_id": "referencia do item", "name": "nome do item", "quantity": "1", "unit_amount": "100" } ], "amount": { "currency": "BRL", "additional": 100, "discount": 100 }, "shipping": { "address": { "street": "Avenida Brigadeiro Faria Lima", "number": "1384", "complement": "apto 12", "locality": "Pinheiros", "city": "São Paulo", "region_code": "SP", "country": "BRA", "postal_code": "01452002" } }, "qr_codes": [ { "id": "QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA", "expiration_date": "2021-08-29T20:15:59-03:00", "amount": { "value": 100 }, "text": "00020101021226830014br.gov.bcb.pix2561api.pagseguro.com/pix/v2/9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA27600016BR.COM.PAGSEGURO01369E13BFE1-35C3-4DFD-B499-9B110AC0E1BA52045697530398654041.005802BR5925Leticia Oliveira Porto La6007Barueri62070503***6304658F", "links": [ { "rel": "QRCODE.PNG", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/png", "media": "image/png", "type": "GET" }, { "rel": "QRCODE.BASE64", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/base64", "media": "text/plain", "type": "GET" } ] } ], "charges": [], "links": [ { "rel": "SELF", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "media": "application/json", "type": "GET" }, { "rel": "PAY", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31/pay", "media": "application/json", "type": "POST" } ] }');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        */
        Test.setMock(HttpCalloutMock.class, new CieloIntegrationMock('pix_success'));
        
        Map<String,Object> parameter = new Map<String,Object>{
            'cartId' => [SELECT Id FROM WebCart WHERE Status != 'Closed' LIMIT 1].Id,
            'formaPagamento' => 'pix',
            'valorCashback' => '0',
            'cardParameter' => new Map<String,Object>{
                'parcelas' => '1'
            }
        };

        User b2bUser = [SELECT Id FROM User WHERE username = 'test.b2b@test.com' LIMIT 1][0];
        
        
        test.startTest();

        B2BPaymentController.pay(System.JSON.serialize(parameter));

        test.stopTest();

        List<IntegrationLog__c> errorLogs = [SELECT Id, MessageError__c FROM IntegrationLog__c WHERE OperationStatus__c = 'Error'];
        String errorMessage = '';
        if(!errorLogs.isEmpty()){
            errorMessage = errorLogs[0].MessageError__c;
        }
        System.assertEquals('', errorMessage, 'Unexpected error message');

        checkError();
    }

    @IsTest
    static void paySuccessCC() {
        /*
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{ "id": "ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "reference_id": "ex-00001", "created_at": "2021-01-19T09:30:12.197-03:00", "customer": { "name": "José da Silva", "email": "jose@email.com", "tax_id": "12345678901", "phones": [ { "country": "55", "area": "11", "number": "999999999", "type": "MOBILE" } ] }, "items": [ { "reference_id": "referencia do item", "name": "nome do item", "quantity": "1", "unit_amount": "100" } ], "amount": { "currency": "BRL", "additional": 100, "discount": 100 }, "shipping": { "address": { "street": "Avenida Brigadeiro Faria Lima", "number": "1384", "complement": "apto 12", "locality": "Pinheiros", "city": "São Paulo", "region_code": "SP", "country": "BRA", "postal_code": "01452002" } }, "qr_codes": [ { "id": "QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA", "expiration_date": "2021-08-29T20:15:59-03:00", "amount": { "value": 100 }, "text": "00020101021226830014br.gov.bcb.pix2561api.pagseguro.com/pix/v2/9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA27600016BR.COM.PAGSEGURO01369E13BFE1-35C3-4DFD-B499-9B110AC0E1BA52045697530398654041.005802BR5925Leticia Oliveira Porto La6007Barueri62070503***6304658F", "links": [ { "rel": "QRCODE.PNG", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/png", "media": "image/png", "type": "GET" }, { "rel": "QRCODE.BASE64", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/base64", "media": "text/plain", "type": "GET" } ] } ], "charges": [], "links": [ { "rel": "SELF", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "media": "application/json", "type": "GET" }, { "rel": "PAY", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31/pay", "media": "application/json", "type": "POST" } ] }');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        */
        Test.setMock(HttpCalloutMock.class, new CieloIntegrationMock());
        
        Map<String,Object> parameter = new Map<String,Object>{
            'cartId' => [SELECT Id FROM WebCart WHERE Status != 'Closed' LIMIT 1].Id,
            'formaPagamento' => 'cc',
            'valorCashback' => '0',
            'cardParameter' => new Map<String,Object>{
                'numeroCartao' => '5502095822650000',
                'cvvCartao' => '123',
                'anoCartao' => '2099',
                'mesCartao' => '1',
                'nomeCartao' => 'TEST',
                'parcelas' => '1x'
            }
        };

        User b2bUser = [SELECT Id FROM User WHERE username = 'test.b2b@test.com' LIMIT 1][0];
        
        
        test.startTest();
            System.runAs(b2bUser){
        	    B2BPaymentController.pay(System.JSON.serialize(parameter));
            }
        test.stopTest();

        List<IntegrationLog__c> errorLogs = [SELECT Id, MessageError__c FROM IntegrationLog__c WHERE OperationStatus__c = 'Error'];
        String errorMessage = '';
        if(!errorLogs.isEmpty()){
            errorMessage = errorLogs[0].MessageError__c;
        }
        System.assertEquals('', errorMessage, 'Unexpected error message');
    }

    @IsTest
    static void payInvalidCC() {
        /*
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{ "id": "ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "reference_id": "ex-00001", "created_at": "2021-01-19T09:30:12.197-03:00", "customer": { "name": "José da Silva", "email": "jose@email.com", "tax_id": "12345678901", "phones": [ { "country": "55", "area": "11", "number": "999999999", "type": "MOBILE" } ] }, "items": [ { "reference_id": "referencia do item", "name": "nome do item", "quantity": "1", "unit_amount": "100" } ], "amount": { "currency": "BRL", "additional": 100, "discount": 100 }, "shipping": { "address": { "street": "Avenida Brigadeiro Faria Lima", "number": "1384", "complement": "apto 12", "locality": "Pinheiros", "city": "São Paulo", "region_code": "SP", "country": "BRA", "postal_code": "01452002" } }, "qr_codes": [ { "id": "QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA", "expiration_date": "2021-08-29T20:15:59-03:00", "amount": { "value": 100 }, "text": "00020101021226830014br.gov.bcb.pix2561api.pagseguro.com/pix/v2/9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA27600016BR.COM.PAGSEGURO01369E13BFE1-35C3-4DFD-B499-9B110AC0E1BA52045697530398654041.005802BR5925Leticia Oliveira Porto La6007Barueri62070503***6304658F", "links": [ { "rel": "QRCODE.PNG", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/png", "media": "image/png", "type": "GET" }, { "rel": "QRCODE.BASE64", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/base64", "media": "text/plain", "type": "GET" } ] } ], "charges": [], "links": [ { "rel": "SELF", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "media": "application/json", "type": "GET" }, { "rel": "PAY", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31/pay", "media": "application/json", "type": "POST" } ] }');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        */
        Test.setMock(HttpCalloutMock.class, new CieloIntegrationMock('39'));
        
        Map<String,Object> parameter = new Map<String,Object>{
            'cartId' => [SELECT Id FROM WebCart WHERE Status != 'Closed' LIMIT 1].Id,
            'formaPagamento' => 'cc',
            'valorCashback' => '0',
            'cardParameter' => new Map<String,Object>{
                'numeroCartao' => '5502095822650000',
                'cvvCartao' => '123',
                'anoCartao' => '2099',
                'mesCartao' => '1',
                'nomeCartao' => 'TEST',
                'parcelas' => '1x'
            }
        };

        User b2bUser = [SELECT Id FROM User WHERE username = 'test.b2b@test.com' LIMIT 1][0];
        
        
        test.startTest();
            System.runAs(b2bUser){
        	    B2BPaymentController.pay(System.JSON.serialize(parameter));
            }
        test.stopTest();

        List<IntegrationLog__c> errorLogs = [SELECT Id, MessageError__c FROM IntegrationLog__c WHERE OperationStatus__c = 'Error'];
        String errorMessage = '';
        if(!errorLogs.isEmpty()){
            errorMessage = errorLogs[0].MessageError__c;
        }
        System.assertEquals('', errorMessage, 'Unexpected error message');
    }

    @IsTest
    static void paySuccessBoleto() {
        /*
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{ "id": "ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "reference_id": "ex-00001", "created_at": "2021-01-19T09:30:12.197-03:00", "customer": { "name": "José da Silva", "email": "jose@email.com", "tax_id": "12345678901", "phones": [ { "country": "55", "area": "11", "number": "999999999", "type": "MOBILE" } ] }, "items": [ { "reference_id": "referencia do item", "name": "nome do item", "quantity": "1", "unit_amount": "100" } ], "amount": { "currency": "BRL", "additional": 100, "discount": 100 }, "shipping": { "address": { "street": "Avenida Brigadeiro Faria Lima", "number": "1384", "complement": "apto 12", "locality": "Pinheiros", "city": "São Paulo", "region_code": "SP", "country": "BRA", "postal_code": "01452002" } }, "qr_codes": [ { "id": "QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA", "expiration_date": "2021-08-29T20:15:59-03:00", "amount": { "value": 100 }, "text": "00020101021226830014br.gov.bcb.pix2561api.pagseguro.com/pix/v2/9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA27600016BR.COM.PAGSEGURO01369E13BFE1-35C3-4DFD-B499-9B110AC0E1BA52045697530398654041.005802BR5925Leticia Oliveira Porto La6007Barueri62070503***6304658F", "links": [ { "rel": "QRCODE.PNG", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/png", "media": "image/png", "type": "GET" }, { "rel": "QRCODE.BASE64", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/base64", "media": "text/plain", "type": "GET" } ] } ], "charges": [], "links": [ { "rel": "SELF", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "media": "application/json", "type": "GET" }, { "rel": "PAY", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31/pay", "media": "application/json", "type": "POST" } ] }');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        */
        Test.setMock(HttpCalloutMock.class, new CieloIntegrationMock());
        
        Map<String,Object> parameter = new Map<String,Object>{
            'cartId' => [SELECT Id FROM WebCart WHERE Status != 'Closed' LIMIT 1].Id,
            'formaPagamento' => 'boleto',
            'valorCashback' => '0'
        };

        User b2bUser = [SELECT Id FROM User WHERE username = 'test.b2b@test.com' LIMIT 1][0];
        
        
        test.startTest();
            System.runAs(b2bUser){
        	    B2BPaymentController.pay(System.JSON.serialize(parameter));
            }
        test.stopTest();

        List<IntegrationLog__c> errorLogs = [SELECT Id, MessageError__c FROM IntegrationLog__c WHERE OperationStatus__c = 'Error'];
        String errorMessage = '';
        if(!errorLogs.isEmpty()){
            errorMessage = errorLogs[0].MessageError__c;
        }
        System.assertEquals('', errorMessage, 'Unexpected error message');
    }
    
    @IsTest
    static void payError01() {
        Account acc = [
            SELECT Id, Name
            FROM Account
            LIMIT 1
        ];
        
        acc.Name = 'error01';
        update acc;
        /*
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{ "id": "ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "reference_id": "ex-00001", "created_at": "2021-01-19T09:30:12.197-03:00", "customer": { "name": "José da Silva", "email": "jose@email.com", "tax_id": "12345678901", "phones": [ { "country": "55", "area": "11", "number": "999999999", "type": "MOBILE" } ] }, "items": [ { "reference_id": "referencia do item", "name": "nome do item", "quantity": "1", "unit_amount": "100" } ], "amount": { "currency": "BRL", "additional": 100, "discount": 100 }, "shipping": { "address": { "street": "Avenida Brigadeiro Faria Lima", "number": "1384", "complement": "apto 12", "locality": "Pinheiros", "city": "São Paulo", "region_code": "SP", "country": "BRA", "postal_code": "01452002" } }, "qr_codes": [ { "id": "QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA", "expiration_date": "2021-08-29T20:15:59-03:00", "amount": { "value": 100 }, "text": "00020101021226830014br.gov.bcb.pix2561api.pagseguro.com/pix/v2/9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA27600016BR.COM.PAGSEGURO01369E13BFE1-35C3-4DFD-B499-9B110AC0E1BA52045697530398654041.005802BR5925Leticia Oliveira Porto La6007Barueri62070503***6304658F", "links": [ { "rel": "QRCODE.PNG", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/png", "media": "image/png", "type": "GET" }, { "rel": "QRCODE.BASE64", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/base64", "media": "text/plain", "type": "GET" } ] } ], "charges": [], "links": [ { "rel": "SELF", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "media": "application/json", "type": "GET" }, { "rel": "PAY", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31/pay", "media": "application/json", "type": "POST" } ] }');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        */
        Test.setMock(HttpCalloutMock.class, new CieloIntegrationMock());
        
        Map<String,Object> parameter = new Map<String,Object>{
            'cartId' => [SELECT Id FROM WebCart WHERE Status != 'Closed' LIMIT 1].Id,
            'formaPagamento' => 'pix'
        };
        
        User b2bUser = [SELECT Id FROM User WHERE username = 'test.b2b@test.com' LIMIT 1][0];
        
        test.startTest();
            System.runAs(b2bUser){
        	    B2BPaymentController.pay(System.JSON.serialize(parameter));
            }
        test.stopTest();
    }
}