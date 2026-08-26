@IsTest
public class B2BPagSeguroRetornoTest {
	@IsTest
    static void parseSucess01() {
        B2BPagSeguroRetorno retorno = new B2BPagSeguroRetorno();
        
        Test.startTest();
		retorno = B2BPagSeguroRetorno.parse('{ "id":"isTest", "reference_id":"isTest", "created_at":"2024-04-30T14:31:20.783-03:00", "customer":{ "name":"isTest", "email":"istest.istest@istest.com.br", "tax_id":"istest" }, "items":[ { "reference_id":"", "name":"isTest", "quantity":1, "unit_amount":1 } ], "shipping":{ "address":{ "street":"isTest", "number":"010", "locality":"isTest", "city":"isTest", "region_code":"SP", "country":"BRA", "postal_code":"10101010" } }, "charges":[ { "id":"isTest", "reference_id":"isTest", "status":"PAID", "created_at":"2024-04-30T14:31:21.402-03:00", "paid_at":"2024-04-30T14:31:23.000-03:00", "description":"isTest", "amount":{ "value":1, "currency":"BRL", "summary":{ "total":1, "paid":1, "refunded":0 } }, "payment_response":{ "code":"20000", "message":"SUCESSO", "reference":"isTest", "raw_data":{ "authorization_code":"isTest", "nsu":"isTest", "reason_code":"00" } }, "payment_method":{ "type":"CREDIT_CARD", "installments":1, "capture":true, "card":{ "id":"isTest", "brand":"visa", "first_digits":"000000", "last_digits":"0000", "exp_month":"0", "exp_year":"0000", "holder":{ "name":"isTest" } }, "soft_descriptor":"sellervirtual" }, "links":[ { "rel":"SELF", "href":"isTest", "media":"application/json", "type":"GET" }, { "rel":"CHARGE.CANCEL", "href":"isTest", "media":"application/json", "type":"POST" } ], "metadata":{ } } ], "notification_urls":[ "isTest" ], "links":[ { "rel":"SELF", "href":"isTest", "media":"application/json", "type":"GET" }, { "rel":"PAY", "href":"isTest", "media":"application/json", "type":"POST" } ] }');
    	System.assertEquals(true, retorno != null);
        Test.stopTest();
    }
}