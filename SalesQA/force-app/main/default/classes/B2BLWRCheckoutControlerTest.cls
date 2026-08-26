@IsTest
public class B2BLWRCheckoutControlerTest {
	
    @TestSetup
    static void makeData() {
        TestUtil util = new TestUtil();
        util.genAllData();
    }
    
    @IsTest
    static void getCashbackSuccess01() {
        User us = [
            SELECT Id
            FROM User
            WHERE Alias = 'isTest'
        ];
        
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"","msg_retorno":[{"tp_extverba":[{"cli_codigo":0000,"hvr_dtmov":"2000-01-01","hvr_tpverba":0,"hvr_tpmov":false,"hvr_tpdoc":"SaldodoDia","hvr_docto":"","hvr_valor":0.0,"ite_codigo":0,"hvr_saldo":1.00,"hvr_seq":1}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        test.startTest();
        System.runAs(us) {
        	B2BLWRCheckoutControler.getCashback();   
        }
        test.stopTest();
    }
    
    @IsTest
    static void getCashbackError01() {
        Account acc = [
            SELECT Id, Name
            FROM Account
            LIMIT 1
        ];
        
        acc.Name = 'error01';
        update acc;
            
        User us = [
            SELECT Id
            FROM User
            WHERE Alias = 'isTest'
        ];
        
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"","msg_retorno":[{"tp_extverba":[{"cli_codigo":0000,"hvr_dtmov":"2000-01-01","hvr_tpverba":0,"hvr_tpmov":false,"hvr_tpdoc":"SaldodoDia","hvr_docto":"","hvr_valor":0.0,"ite_codigo":0,"hvr_saldo":1.00,"hvr_seq":1}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        test.startTest();
        System.runAs(us) {
        	B2BLWRCheckoutControler.getCashback();   
        }
        test.stopTest();
    }
    
    @IsTest
    static void getCashbackError02() {
        Account acc = [
            SELECT Id, Name
            FROM Account
            LIMIT 1
        ];
        
        acc.Name = 'error02';
        update acc;
            
        User us = [
            SELECT Id
            FROM User
            WHERE Alias = 'isTest'
        ];
        
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"","msg_retorno":[{"tp_extverba":[{"cli_codigo":0000,"hvr_dtmov":"2000-01-01","hvr_tpverba":0,"hvr_tpmov":false,"hvr_tpdoc":"SaldodoDia","hvr_docto":"","hvr_valor":0.0,"ite_codigo":0,"hvr_saldo":1.00,"hvr_seq":1}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        test.startTest();
        System.runAs(us) {
        	B2BLWRCheckoutControler.getCashback();   
        }
        test.stopTest();
    }

        
    @IsTest
    static void setCashback() {        
        test.startTest();
        	B2BLWRCheckoutControler.setCashBackValue(0.1);   
        test.stopTest();
    }

        @IsTest
    static void setHighCashback() {        
        test.startTest();
        	B2BLWRCheckoutControler.setCashBackValue(999999999999999.0);   
        test.stopTest();
    }
    /*
    @IsTest
    static void createOrderSuccess01() {
        User us = [
            SELECT Id
            FROM User
            WHERE Alias = 'isTest'
        ];
        
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"","msg_retorno":[{"tp_extverba":[{"cli_codigo":10101,"hvr_dtmov":"2001-01-01","hvr_tpverba":0,"hvr_tpmov":false,"hvr_tpdoc":"SaldodoDia","hvr_docto":"","hvr_valor":0.0,"ite_codigo":0,"hvr_saldo":1006.05,"hvr_seq":1}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        test.startTest();
        System.runAs(us) {
        	B2BLWRCheckoutControler.createOrder('10');   
        }
        test.stopTest();
    }	
    
    @IsTest
    static void createOrderError01() {
        ContactPointAddress cpa = [
            SELECT Id
            FROM ContactPointAddress
            WHERE AddressType = 'Shipping'
            AND Name = 'isTest'
            LIMIT 1
        ];
        delete cpa;
        
        User us = [
            SELECT Id
            FROM User
            WHERE Alias = 'isTest'
        ];
        
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"","msg_retorno":[{"tp_extverba":[{"cli_codigo":10101,"hvr_dtmov":"2001-01-01","hvr_tpverba":0,"hvr_tpmov":false,"hvr_tpdoc":"SaldodoDia","hvr_docto":"","hvr_valor":0.0,"ite_codigo":0,"hvr_saldo":1006.05,"hvr_seq":1}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        test.startTest();
        System.runAs(us) {
        	B2BLWRCheckoutControler.createOrder('10');   
        }
        test.stopTest();
    }
    */
    @IsTest
    static void paySuccess01() {
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{ "id": "ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "reference_id": "ex-00001", "created_at": "2021-01-19T09:30:12.197-03:00", "customer": { "name": "José da Silva", "email": "jose@email.com", "tax_id": "12345678901", "phones": [ { "country": "55", "area": "11", "number": "999999999", "type": "MOBILE" } ] }, "items": [ { "reference_id": "referencia do item", "name": "nome do item", "quantity": "1", "unit_amount": "100" } ], "amount": { "currency": "BRL", "additional": 100, "discount": 100 }, "shipping": { "address": { "street": "Avenida Brigadeiro Faria Lima", "number": "1384", "complement": "apto 12", "locality": "Pinheiros", "city": "São Paulo", "region_code": "SP", "country": "BRA", "postal_code": "01452002" } }, "qr_codes": [ { "id": "QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA", "expiration_date": "2021-08-29T20:15:59-03:00", "amount": { "value": 100 }, "text": "00020101021226830014br.gov.bcb.pix2561api.pagseguro.com/pix/v2/9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA27600016BR.COM.PAGSEGURO01369E13BFE1-35C3-4DFD-B499-9B110AC0E1BA52045697530398654041.005802BR5925Leticia Oliveira Porto La6007Barueri62070503***6304658F", "links": [ { "rel": "QRCODE.PNG", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/png", "media": "image/png", "type": "GET" }, { "rel": "QRCODE.BASE64", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/base64", "media": "text/plain", "type": "GET" } ] } ], "charges": [], "links": [ { "rel": "SELF", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "media": "application/json", "type": "GET" }, { "rel": "PAY", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31/pay", "media": "application/json", "type": "POST" } ] }');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        B2BLWRCheckoutControler.PaymentParameter parameter = new B2BLWRCheckoutControler.PaymentParameter();
		parameter.cartId= [SELECT Id FROM WebCart LIMIT 1].Id;
		parameter.parcelas = '1';
		parameter.formaPagamento = 'pix';
		parameter.valorCashback = '0';
        
        User us = [
            SELECT Id
            FROM User
            WHERE Alias = 'isTest'
        ];
        
        test.startTest();
        System.runAs(us) {
        	B2BLWRCheckoutControler.pay(System.JSON.serialize(parameter));   
        }
        test.stopTest();
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
        
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{ "id": "ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "reference_id": "ex-00001", "created_at": "2021-01-19T09:30:12.197-03:00", "customer": { "name": "José da Silva", "email": "jose@email.com", "tax_id": "12345678901", "phones": [ { "country": "55", "area": "11", "number": "999999999", "type": "MOBILE" } ] }, "items": [ { "reference_id": "referencia do item", "name": "nome do item", "quantity": "1", "unit_amount": "100" } ], "amount": { "currency": "BRL", "additional": 100, "discount": 100 }, "shipping": { "address": { "street": "Avenida Brigadeiro Faria Lima", "number": "1384", "complement": "apto 12", "locality": "Pinheiros", "city": "São Paulo", "region_code": "SP", "country": "BRA", "postal_code": "01452002" } }, "qr_codes": [ { "id": "QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA", "expiration_date": "2021-08-29T20:15:59-03:00", "amount": { "value": 100 }, "text": "00020101021226830014br.gov.bcb.pix2561api.pagseguro.com/pix/v2/9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA27600016BR.COM.PAGSEGURO01369E13BFE1-35C3-4DFD-B499-9B110AC0E1BA52045697530398654041.005802BR5925Leticia Oliveira Porto La6007Barueri62070503***6304658F", "links": [ { "rel": "QRCODE.PNG", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/png", "media": "image/png", "type": "GET" }, { "rel": "QRCODE.BASE64", "href": "https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/base64", "media": "text/plain", "type": "GET" } ] } ], "charges": [], "links": [ { "rel": "SELF", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31", "media": "application/json", "type": "GET" }, { "rel": "PAY", "href": "https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31/pay", "media": "application/json", "type": "POST" } ] }');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        B2BLWRCheckoutControler.PaymentParameter parameter = new B2BLWRCheckoutControler.PaymentParameter();
		parameter.cartId= [SELECT Id FROM WebCart LIMIT 1].Id;
		parameter.parcelas = '1';
		parameter.formaPagamento = 'pix';
		parameter.valorCashback = '0';
        
        User us = [
            SELECT Id
            FROM User
            WHERE Alias = 'isTest'
        ];
        
        test.startTest();
        System.runAs(us) {
        	B2BLWRCheckoutControler.pay(System.JSON.serialize(parameter));   
        }
        test.stopTest();
    }
    /*
    @IsTest
    static void closeWebCartPixSuccess01() {
        B2BPagSeguroRetorno pagSeguroRetorno = B2BPagSeguroRetorno.parse('{"id":"ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31","reference_id":"ex-00001","created_at":"2021-01-19T09:30:12.197-03:00","customer":{"name":"JosédaSilva","email":"jose@email.com","tax_id":"12345678901","phones":[{"country":"55","area":"11","number":"999999999","type":"MOBILE"}]},"items":[{"reference_id":"referenciadoitem","name":"nomedoitem","quantity":"1","unit_amount":"100"}],"amount":{"currency":"BRL","additional":100,"discount":100},"shipping":{"address":{"street":"AvenidaBrigadeiroFariaLima","number":"1384","complement":"apto12","locality":"Pinheiros","city":"SãoPaulo","region_code":"SP","country":"BRA","postal_code":"01452002"}},"qr_codes":[{"id":"QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA","expiration_date":"2021-08-29T20:15:59-03:00","amount":{"value":100},"text":"00020101021226830014br.gov.bcb.pix2561api.pagseguro.com/pix/v2/9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA27600016BR.COM.PAGSEGURO01369E13BFE1-35C3-4DFD-B499-9B110AC0E1BA52045697530398654041.005802BR5925LeticiaOliveiraPortoLa6007Barueri62070503***6304658F","links":[{"rel":"QRCODE.PNG","href":"https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/png","media":"image/png","type":"GET"},{"rel":"QRCODE.BASE64","href":"https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/base64","media":"text/plain","type":"GET"}]}],"charges":[],"links":[{"rel":"SELF","href":"https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31","media":"application/json","type":"GET"},{"rel":"PAY","href":"https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31/pay","media":"application/json","type":"POST"}]}');
        B2BLWRCheckoutControler.PaymentReturn parameter = new B2BLWRCheckoutControler.PaymentReturn();
		parameter.cartId= [SELECT Id FROM WebCart LIMIT 1].Id;
		parameter.parcelas = '1';
		parameter.formaPagamento = 'pix';
		parameter.valorCashback = '0';
        parameter.pagSeguroReturn = pagSeguroRetorno;
        
        User us = [
            SELECT Id
            FROM User
            WHERE Alias = 'isTest'
        ];
        
        test.startTest();
        System.runAs(us) {
        	B2BLWRCheckoutControler.closeWebCart(System.JSON.serialize(parameter));   
        }
        test.stopTest();
    }
    
    @IsTest
    static void closeWebCartCartaoCreditoSuccess01() {
        B2BPagSeguroRetorno pagSeguroRetorno = B2BPagSeguroRetorno.parse('{"id":"ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31","reference_id":"ex-00001","created_at":"2021-01-19T09:30:12.197-03:00","customer":{"name":"JosédaSilva","email":"jose@email.com","tax_id":"12345678901","phones":[{"country":"55","area":"11","number":"999999999","type":"MOBILE"}]},"items":[{"reference_id":"referenciadoitem","name":"nomedoitem","quantity":"1","unit_amount":"100"}],"amount":{"currency":"BRL","additional":100,"discount":100},"shipping":{"address":{"street":"AvenidaBrigadeiroFariaLima","number":"1384","complement":"apto12","locality":"Pinheiros","city":"SãoPaulo","region_code":"SP","country":"BRA","postal_code":"01452002"}},"qr_codes":[{"id":"QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA","expiration_date":"2021-08-29T20:15:59-03:00","amount":{"value":100},"text":"00020101021226830014br.gov.bcb.pix2561api.pagseguro.com/pix/v2/9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA27600016BR.COM.PAGSEGURO01369E13BFE1-35C3-4DFD-B499-9B110AC0E1BA52045697530398654041.005802BR5925LeticiaOliveiraPortoLa6007Barueri62070503***6304658F","links":[{"rel":"QRCODE.PNG","href":"https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/png","media":"image/png","type":"GET"},{"rel":"QRCODE.BASE64","href":"https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/base64","media":"text/plain","type":"GET"}]}],"charges":[],"links":[{"rel":"SELF","href":"https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31","media":"application/json","type":"GET"},{"rel":"PAY","href":"https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31/pay","media":"application/json","type":"POST"}]}');
        B2BLWRCheckoutControler.PaymentReturn parameter = new B2BLWRCheckoutControler.PaymentReturn();
		parameter.cartId= [SELECT Id FROM WebCart LIMIT 1].Id;
		parameter.parcelas = '1';
		parameter.formaPagamento = 'cc';
		parameter.valorCashback = '0';
        parameter.pagSeguroReturn = pagSeguroRetorno;
        
        User us = [
            SELECT Id
            FROM User
            WHERE Alias = 'isTest'
        ];
        
        test.startTest();
        System.runAs(us) {
        	B2BLWRCheckoutControler.closeWebCart(System.JSON.serialize(parameter));   
        }
        test.stopTest();
    }
    
    @IsTest
    static void closeWebCartBoletoSuccess01() {
        B2BPagSeguroRetorno pagSeguroRetorno = B2BPagSeguroRetorno.parse('{"id":"ORDE_E8A44EB1-89BC-4261-9004-739A09B53A8C","reference_id":"ex-00001","created_at":"2023-02-08T15:10:15.763-03:00","customer":{"name":"JosedaSilva","email":"email@test.com","tax_id":"12345678909","phones":[{"type":"MOBILE","country":"55","area":"11","number":"999999999"}]},"items":[{"reference_id":"referenciadoitem","name":"nomedoitem","quantity":1,"unit_amount":500}],"shipping":{"address":{"street":"AvenidaBrigadeiroFariaLima","number":"1384","complement":"apto12","locality":"Pinheiros","city":"SãoPaulo","region_code":"SP","country":"BRA","postal_code":"01452002"}},"charges":[{"id":"CHAR_BF10AE3B-68D2-430D-AED6-0D5E2C7E180E","reference_id":"referenciadacobranca","status":"WAITING","created_at":"2023-02-08T15:10:15.901-03:00","description":"descricaodacobranca","amount":{"value":500,"currency":"BRL","summary":{"total":500,"paid":0,"refunded":0}},"payment_response":{"code":"20000","message":"SUCESSO"},"payment_method":{"type":"BOLETO","boleto":{"id":"0E243EE1-736A-45EF-B153-CC9B9F0A838F","barcode":"03399853012970000024227020901016278150000015630","formatted_barcode":"03399.8530129700.00024227020.901016278150000015630","due_date":"2023-06-20","instruction_lines":{"line_1":"PagamentoprocessadoparaDESCFatura","line_2":"ViaPagSeguro"},"holder":{"name":"JosedaSilva","tax_id":"12345679891","email":"jose@email.com","address":{"region":"SãoPaulo","city":"SaoPaulo","postal_code":"01452002","street":"AvenidaBrigadeiroFariaLima","number":"1384","locality":"Pinheiros","country":"Brasil","region_code":"SP"}}}},"links":[{"rel":"SELF","href":"https://boleto.sandbox.pagseguro.com.br/0e243ee1-736a-45ef-b153-cc9b9f0a838f.pdf","media":"application/pdf","type":"GET"},{"rel":"SELF","href":"https://boleto.sandbox.pagseguro.com.br/0e243ee1-736a-45ef-b153-cc9b9f0a838f.png","media":"image/png","type":"GET"},{"rel":"SELF","href":"https://sandbox.api.pagseguro.com/charges/CHAR_BF10AE3B-68D2-430D-AED6-0D5E2C7E180E","media":"application/json","type":"GET"}]}],"notification_urls":["https://meusite.com/notificacoes"],"links":[{"rel":"SELF","href":"https://sandbox.api.pagseguro.com/orders/ORDE_E8A44EB1-89BC-4261-9004-739A09B53A8C","media":"application/json","type":"GET"},{"rel":"PAY","href":"https://sandbox.api.pagseguro.com/orders/ORDE_E8A44EB1-89BC-4261-9004-739A09B53A8C/pay","media":"application/json","type":"POST"}]}');
        B2BLWRCheckoutControler.PaymentReturn parameter = new B2BLWRCheckoutControler.PaymentReturn();
		parameter.cartId= [SELECT Id FROM WebCart LIMIT 1].Id;
		parameter.parcelas = '1';
		parameter.formaPagamento = 'boleto';
		parameter.valorCashback = '0';
        parameter.pagSeguroReturn = pagSeguroRetorno;
        
        User us = [
            SELECT Id
            FROM User
            WHERE Alias = 'isTest'
        ];
        
        test.startTest();
        System.runAs(us) {
        	B2BLWRCheckoutControler.closeWebCart(System.JSON.serialize(parameter));   
        }
        test.stopTest();
    }
    
    @IsTest
    static void closeWebCartPixError01() {
        B2BPagSeguroRetorno pagSeguroRetorno = B2BPagSeguroRetorno.parse('{"id":"ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31","reference_id":"ex-00001","created_at":"2021-01-19T09:30:12.197-03:00","customer":{"name":"JosédaSilva","email":"jose@email.com","tax_id":"12345678901","phones":[{"country":"55","area":"11","number":"999999999","type":"MOBILE"}]},"items":[{"reference_id":"referenciadoitem","name":"nomedoitem","quantity":"1","unit_amount":"100"}],"amount":{"currency":"BRL","additional":100,"discount":100},"shipping":{"address":{"street":"AvenidaBrigadeiroFariaLima","number":"1384","complement":"apto12","locality":"Pinheiros","city":"SãoPaulo","region_code":"SP","country":"BRA","postal_code":"01452002"}},"qr_codes":[{"id":"QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA","expiration_date":"2021-08-29T20:15:59-03:00","amount":{"value":100},"text":"00020101021226830014br.gov.bcb.pix2561api.pagseguro.com/pix/v2/9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA27600016BR.COM.PAGSEGURO01369E13BFE1-35C3-4DFD-B499-9B110AC0E1BA52045697530398654041.005802BR5925LeticiaOliveiraPortoLa6007Barueri62070503***6304658F","links":[{"rel":"QRCODE.PNG","href":"https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/png","media":"image/png","type":"GET"},{"rel":"QRCODE.BASE64","href":"https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/base64","media":"text/plain","type":"GET"}]}],"charges":[],"links":[{"rel":"SELF","href":"https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31","media":"application/json","type":"GET"},{"rel":"PAY","href":"https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31/pay","media":"application/json","type":"POST"}]}');
        B2BLWRCheckoutControler.PaymentReturn parameter = new B2BLWRCheckoutControler.PaymentReturn();
		parameter.nomeCartao = 'errortest';
        parameter.cartId= [SELECT Id FROM WebCart LIMIT 1].Id;
		parameter.parcelas = '1';
		parameter.formaPagamento = 'pix';
		parameter.valorCashback = '0';
        parameter.pagSeguroReturn = pagSeguroRetorno;
        
        User us = [
            SELECT Id
            FROM User
            WHERE Alias = 'isTest'
        ];
        
        test.startTest();
        System.runAs(us) {
        	B2BLWRCheckoutControler.closeWebCart(System.JSON.serialize(parameter));   
        }
        test.stopTest();
    }
    
    @IsTest
    static void closeWebCartCartaoCreditoError01() {
        B2BPagSeguroRetorno pagSeguroRetorno = B2BPagSeguroRetorno.parse('{"id":"ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31","reference_id":"ex-00001","created_at":"2021-01-19T09:30:12.197-03:00","customer":{"name":"JosédaSilva","email":"jose@email.com","tax_id":"12345678901","phones":[{"country":"55","area":"11","number":"999999999","type":"MOBILE"}]},"items":[{"reference_id":"referenciadoitem","name":"nomedoitem","quantity":"1","unit_amount":"100"}],"amount":{"currency":"BRL","additional":100,"discount":100},"shipping":{"address":{"street":"AvenidaBrigadeiroFariaLima","number":"1384","complement":"apto12","locality":"Pinheiros","city":"SãoPaulo","region_code":"SP","country":"BRA","postal_code":"01452002"}},"qr_codes":[{"id":"QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA","expiration_date":"2021-08-29T20:15:59-03:00","amount":{"value":100},"text":"00020101021226830014br.gov.bcb.pix2561api.pagseguro.com/pix/v2/9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA27600016BR.COM.PAGSEGURO01369E13BFE1-35C3-4DFD-B499-9B110AC0E1BA52045697530398654041.005802BR5925LeticiaOliveiraPortoLa6007Barueri62070503***6304658F","links":[{"rel":"QRCODE.PNG","href":"https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/png","media":"image/png","type":"GET"},{"rel":"QRCODE.BASE64","href":"https://api.pagseguro.com/qrcode/QRCO_9E13BFE1-35C3-4DFD-B499-9B110AC0E1BA/base64","media":"text/plain","type":"GET"}]}],"charges":[],"links":[{"rel":"SELF","href":"https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31","media":"application/json","type":"GET"},{"rel":"PAY","href":"https://api.pagseguro.com/orders/ORDE_9BBD677F-863E-4D46-BAAF-2EECEE49FF31/pay","media":"application/json","type":"POST"}]}');
        B2BLWRCheckoutControler.PaymentReturn parameter = new B2BLWRCheckoutControler.PaymentReturn();
		parameter.nomeCartao = 'errortest';
        parameter.cartId= [SELECT Id FROM WebCart LIMIT 1].Id;
		parameter.parcelas = '1';
		parameter.formaPagamento = 'cc';
		parameter.valorCashback = '0';
        parameter.pagSeguroReturn = pagSeguroRetorno;
        
        User us = [
            SELECT Id
            FROM User
            WHERE Alias = 'isTest'
        ];
        
        test.startTest();
        System.runAs(us) {
        	B2BLWRCheckoutControler.closeWebCart(System.JSON.serialize(parameter));   
        }
        test.stopTest();
    }

    @IsTest
    static void closeWebCartError01() {
        Account acc = [
            SELECT Id, Name
            FROM Account
            LIMIT 1
        ];
        acc.Name = 'error01';
        update acc;

        B2BLWRCheckoutControler.PaymentReturn parameter = new B2BLWRCheckoutControler.PaymentReturn();
		parameter.cartId= [SELECT Id FROM WebCart LIMIT 1].Id;
		parameter.parcelas = '1';
		parameter.formaPagamento = 'pix';
		parameter.valorCashback = '0';
        parameter.pagSeguroReturn = null;
        
        User us = [
            SELECT Id
            FROM User
            WHERE Alias = 'isTest'
        ];
        
        test.startTest();
        System.runAs(us) {
        	B2BLWRCheckoutControler.closeWebCart(System.JSON.serialize(parameter));   
        }
        test.stopTest();
    }
    */
}